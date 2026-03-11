/**
 * Background service worker for MS LearnFlow.
 *
 * Responsibilities:
 *   - Create the offscreen document when a PDF job starts (popup may be closed mid-job)
 *   - Close the offscreen document when the job finishes or is cancelled
 *
 * The offscreen document reads its job spec directly from chrome.storage.session,
 * so there is no timing race between "document created" and "message received".
 */

const OFFSCREEN_URL = chrome.runtime.getURL('offscreen.html');

async function ensureOffscreen(): Promise<void> {
  // Always tear down any stale offscreen document first.
  // This avoids the race where a leftover page from a previous (failed) job
  // already consumed the mlfPendingJob key and won't pick up the new one.
  await closeOffscreen();
  try {
    await chrome.offscreen.createDocument({
      url: OFFSCREEN_URL,
      // DOM_PARSER: we parse raw HTML with DOMParser
      // DOM_SCRAPING: we manipulate the parsed DOM tree
      reasons: ['DOM_PARSER', 'DOM_SCRAPING'] as unknown as chrome.offscreen.Reason[],
      justification:
        'Parse and clean Microsoft Learn HTML pages, then render a PDF with html2canvas.',
    });
  } catch (err: unknown) {
    // Chrome throws when an offscreen document already exists – that is fine.
    const msg = err instanceof Error ? err.message : String(err);
    if (!msg.includes('single') && !msg.includes('already')) throw err;
  }
}

async function closeOffscreen(): Promise<void> {
  try {
    await chrome.offscreen.closeDocument();
  } catch {
    /* no offscreen doc open – ignore */
  }
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  // Only react to messages explicitly targeted at the background.
  if (msg?.target !== 'background') return false;

  if (msg.type === 'MLF_START_JOB') {
    // Job spec is already stored in chrome.storage.session.mlfPendingJob by the popup.
    // Just create the offscreen document – it will read the spec and start working.
    ensureOffscreen()
      .then(() => sendResponse({ ok: true }))
      .catch((err) => sendResponse({ ok: false, error: String(err) }));
    return true; // will respond asynchronously
  }

  // ── Storage proxy for the offscreen document ──────────────────────────
  // Offscreen pages may not have access to chrome.storage in all Chrome
  // versions, so we act as a proxy.

  if (msg.type === 'MLF_GET_PENDING_JOB') {
    chrome.storage.session.get('mlfPendingJob', (result) => {
      const job = result?.mlfPendingJob ?? null;
      // Clear the key so a reload of the offscreen page won't re-run the job.
      if (job) chrome.storage.session.remove('mlfPendingJob');
      sendResponse(job);
    });
    return true; // async
  }

  if (msg.type === 'MLF_WRITE_JOB') {
    chrome.storage.session.set({ mlfJob: msg.state });
    return false;
  }

  if (msg.type === 'MLF_DOWNLOAD') {
    chrome.downloads.download(
      { url: msg.dataUrl, filename: msg.filename, saveAs: false },
      (downloadId) => {
        if (chrome.runtime.lastError) {
          sendResponse({ ok: false, error: chrome.runtime.lastError.message });
        } else {
          sendResponse({ ok: true, downloadId });
        }
      },
    );
    return true; // async
  }

  if (msg.type === 'MLF_CANCEL_JOB') {
    closeOffscreen().then(() =>
      chrome.storage.session.set({
        mlfJob: { status: 'idle', pct: 0, label: '', docTitle: '', totalUnits: 0 },
      })
    );
    sendResponse({ ok: true });
    return false;
  }

  if (msg.type === 'MLF_JOB_COMPLETE' || msg.type === 'MLF_JOB_ERROR') {
    closeOffscreen();
    return false;
  }

  return false;
});
