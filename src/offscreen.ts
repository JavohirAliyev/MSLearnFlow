/**
 * Offscreen worker for MS LearnFlow.
 *
 * This page runs hidden (chrome.offscreen API) so work continues even when
 * the user closes the extension popup.  It has full DOM access, which is
 * required by DOMParser-based HTML cleaning and html2canvas PDF rendering.
 *
 * Lifecycle:
 *   1. Background service worker creates this page when a job starts.
 *   2. On load, we read the job spec from chrome.storage.session.mlfPendingJob.
 *   3. We run the job (parallel fetch → clean → PDF render → download).
 *   4. Progress is written to chrome.storage.session.mlfJob so the popup can
 *      display live updates when it is open.
 *   5. When done (or on error), we notify the background to close this page.
 */

import { fetchAndCleanUnit } from './services/scraper';
import { generatePDFBlob } from './services/pdfEngine';

// ─── Types ────────────────────────────────────────────────────────────────────

interface PendingJob {
  units: { title: string; url: string }[];
  docTitle: string;
  filename: string;
}

export interface JobState {
  status: 'idle' | 'fetching' | 'rendering' | 'done' | 'error';
  pct: number;
  label: string;
  docTitle: string;
  totalUnits: number;
  error?: string;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const CONCURRENCY = 6;
/** Fraction of total progress bar that the fetch phase occupies (0-100). */
const FETCH_PCT = 65;
/** Hard timeout for the entire job to prevent the offscreen page living forever. */
const JOB_TIMEOUT_MS = 5 * 60_000; // 5 minutes
/** Per-unit hard timeout — if a single unit takes longer, skip it and move on. */
const PER_UNIT_TIMEOUT_MS = 45_000; // 45 seconds

// ─── Helpers ──────────────────────────────────────────────────────────────────

/** Write a complete JobState — proxied through the background service worker
 *  because chrome.storage may not be available in offscreen documents. */
function writeJob(state: JobState): void {
  chrome.runtime.sendMessage({ type: 'MLF_WRITE_JOB', target: 'background', state }).catch(() => {});
}

// ─── Main job ─────────────────────────────────────────────────────────────────

async function runJob(job: PendingJob): Promise<void> {
  const { units, docTitle, filename } = job;
  const total = units.length;

  // ── Phase 1: parallel fetch + clean (0 → FETCH_PCT %) ───────────────────

  writeJob({ status: 'fetching', pct: 0, label: `Fetching 0 / ${total} units…`, docTitle, totalUnits: total });

  const results: ({ title: string; url: string; html: string } | null)[] = new Array(total).fill(null);
  let doneCount = 0;
  let nextIdx = 0;

  // Up to CONCURRENCY workers run simultaneously.
  const workers = Array.from({ length: Math.min(CONCURRENCY, total) }, async () => {
    while (nextIdx < total) {
      const i = nextIdx++;
      try {
        // Race each unit against a hard per-unit timeout so one slow page
        // can never block the entire pipeline.
        results[i] = await Promise.race([
          fetchAndCleanUnit(units[i].url),
          new Promise<null>((resolve) =>
            setTimeout(() => resolve(null), PER_UNIT_TIMEOUT_MS),
          ),
        ]);
      } catch {
        results[i] = null;
      }
      doneCount++;
      const pct = Math.round((doneCount / total) * FETCH_PCT);
      writeJob({
        status: 'fetching',
        pct,
        label: `Fetching ${doneCount} / ${total} units…`,
        docTitle,
        totalUnits: total,
      });
    }
  });

  await Promise.all(workers);

  const cleaned = results.filter((r): r is NonNullable<typeof r> => r !== null);

  if (cleaned.length === 0) {
    writeJob({
      status: 'error',
      pct: 0,
      label: 'Error: All unit pages failed to load. Check your network connection.',
      error: 'No unit content could be fetched.',
      docTitle,
      totalUnits: total,
    });
    chrome.runtime.sendMessage({ type: 'MLF_JOB_ERROR', target: 'background' }).catch(() => {});
    return;
  }

  // ── Phase 2: PDF render (FETCH_PCT → 98 %) ────────────────────────────────

  writeJob({ status: 'rendering', pct: FETCH_PCT, label: 'Rendering document…', docTitle, totalUnits: total });

  const pdfBytes = await generatePDFBlob(
    cleaned,
    filename,
    ({ step, total: pages }) => {
      const pct = Math.round(FETCH_PCT + (step / pages) * (100 - FETCH_PCT - 2));
      writeJob({ status: 'rendering', pct, label: `Building PDF… (page ${step} / ${pages})`, docTitle, totalUnits: total });
    },
    docTitle,
    (phase) => {
      const pct =
        phase === 'Rendering document…' ? FETCH_PCT :
        phase === 'Building PDF…'       ? FETCH_PCT + 5 :
        98;
      writeJob({ status: 'rendering', pct, label: phase, docTitle, totalUnits: total });
    },
  );

  // ── Phase 3: trigger OS download (98 → 100 %) ────────────────────────────

  writeJob({ status: 'rendering', pct: 99, label: 'Preparing download…', docTitle, totalUnits: total });

  // Convert PDF bytes to a base64 data-URL so we can send it to the
  // background worker (object URLs don't cross execution contexts).
  const blob = new Blob([pdfBytes.buffer as ArrayBuffer], { type: 'application/pdf' });
  const dataUrl: string = await new Promise((resolve) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result as string);
    reader.readAsDataURL(blob);
  });

  // Proxy the download through the background service worker because
  // chrome.downloads may be unavailable in offscreen documents.
  await new Promise<void>((resolve, reject) => {
    chrome.runtime.sendMessage(
      { type: 'MLF_DOWNLOAD', target: 'background', dataUrl, filename },
      (resp) => {
        if (resp?.ok) resolve();
        else reject(new Error(resp?.error ?? 'Download failed'));
      },
    );
  });

  writeJob({ status: 'done', pct: 100, label: 'Download complete!', docTitle, totalUnits: total });

  // Tell the background worker to close this page.
  chrome.runtime.sendMessage({ type: 'MLF_JOB_COMPLETE', target: 'background' }).catch(() => {});
}

// ─── Entry point ──────────────────────────────────────────────────────────────

(async () => {
  // Read the pending job from the background service worker (which has full
  // chrome.storage access — offscreen documents may not).
  const mlfPendingJob: PendingJob | null = await new Promise((resolve) => {
    chrome.runtime.sendMessage(
      { type: 'MLF_GET_PENDING_JOB', target: 'background' },
      (response) => resolve(response ?? null),
    );
  });
  if (!mlfPendingJob) return; // nothing to do (shouldn't happen in normal flow)

  runJob(mlfPendingJob as PendingJob)
    .catch(async (err) => {
      const error = err instanceof Error ? err.message : String(err);
      writeJob({
        status: 'error',
        pct: 0,
        label: 'Error: ' + error,
        error,
        docTitle: (mlfPendingJob as PendingJob).docTitle,
        totalUnits: (mlfPendingJob as PendingJob).units.length,
      });
      chrome.runtime.sendMessage({ type: 'MLF_JOB_ERROR', target: 'background' }).catch(() => {});
    });

  // Hard timeout safeguard — kills this page if the job hangs
  setTimeout(() => {
    writeJob({
      status: 'error',
      pct: 0,
      label: 'Error: Job timed out after 5 minutes.',
      error: 'Job timed out.',
      docTitle: (mlfPendingJob as PendingJob).docTitle,
      totalUnits: (mlfPendingJob as PendingJob).units.length,
    });
    chrome.runtime.sendMessage({ type: 'MLF_JOB_ERROR', target: 'background' }).catch(() => {});
  }, JOB_TIMEOUT_MS);
})();
