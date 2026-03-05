import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

/**
 * Generate a single PDF from multiple cleaned HTML segments.
 * - segments: array with { title, url, html }
 * - filename: output filename
 * - onProgress: optional callback(progress: { step: number, total: number })
 */
export async function generatePDF(
  segments: { title: string; url: string; html: string }[],
  filename = 'learnflow.pdf',
  onProgress?: (p: { step: number; total: number }) => void,
  moduleTitle?: string,
) {
  // Prepare PDF sizes in points
  const pdf = new jsPDF({ unit: 'pt', format: 'a4', compress: true });
  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = pdf.internal.pageSize.getHeight();

  // Render all segments into a single off-screen container so content flows naturally
  // and we can slice pages while avoiding splitting protected elements.
  const container = document.createElement('div');
  container.style.width = '794px'; // visual width for rendering (approx A4 @ 96dpi)
  container.style.background = '#ffffff';
  container.style.color = '#111827';
  container.style.padding = '0';
  container.style.boxSizing = 'border-box';
  container.style.fontFamily = "'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif";
  container.style.fontSize = '13px';
  container.style.lineHeight = '1.45';
  container.style.maxWidth = '734px';

  // Insert module / learning-path title at the very top
  if (moduleTitle) {
    const banner = document.createElement('div');
    banner.style.fontSize = '26px';
    banner.style.fontWeight = '800';
    banner.style.color = '#1e3a8a';
    banner.style.padding = '18px 30px 10px 30px';
    banner.style.lineHeight = '1.25';
    banner.style.letterSpacing = '0.01em';
    banner.textContent = moduleTitle;
    container.appendChild(banner);

    // Thin accent line under the title
    const rule = document.createElement('hr');
    rule.style.border = 'none';
    rule.style.height = '3px';
    rule.style.background = 'linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%)';
    rule.style.margin = '0 30px 6px 30px';
    rule.style.borderRadius = '2px';
    container.appendChild(rule);
  }

  // Append each segment as a panel inside the single container
  segments.forEach((seg) => {
    const panel = document.createElement('div');
    panel.className = 'panel';
    panel.style.margin = '8px 0';
    panel.style.background = '#fff';
    panel.style.maxWidth = '734px';
    panel.style.boxShadow = 'none';
    panel.style.border = 'none';
    panel.style.borderRadius = '10px';
    panel.style.padding = '28px 30px';

    const header = document.createElement('div');
    header.className = 'header';
    header.style.fontSize = '22px';
    header.style.fontWeight = '700';
    header.style.marginBottom = '18px';
    header.style.color = '#1e3a8a';
    header.style.letterSpacing = '0.01em';
    header.style.lineHeight = '1.2';
    header.style.textAlign = 'left';
    header.textContent = seg.title || '';

    const content = document.createElement('div');
    content.className = 'unit';
    content.innerHTML = seg.html || '';

    // Remove any h1/h2/h3 headers at the very top of the content (to avoid duplicate headers)
    while (content.firstChild && content.firstChild.nodeType === 1 && ['H1', 'H2', 'H3'].includes((content.firstChild as HTMLElement).tagName)) {
      content.removeChild(content.firstChild);
    }

    // Small niceties for readability and tables/code styling
    content.style.color = '#0b1220';
    content.style.marginBottom = '0';
    content.querySelectorAll('img').forEach((img) => {
      (img as HTMLImageElement).style.maxWidth = '100%';
      (img as HTMLImageElement).style.height = 'auto';
      (img as HTMLImageElement).style.borderRadius = '6px';
      (img as HTMLImageElement).style.boxShadow = '0 2px 8px rgba(16,24,40,0.08)';
      (img as HTMLImageElement).style.margin = '12px 0';
    });
    content.querySelectorAll('p').forEach((p) => { (p as HTMLElement).style.margin = '12px 0'; });
    content.querySelectorAll('ul,ol').forEach((el) => { (el as HTMLElement).style.margin = '12px 0 12px 28px'; });
    content.querySelectorAll('h1,h2,h3,h4,h5').forEach((h) => { (h as HTMLElement).style.margin = '20px 0 10px 0'; (h as HTMLElement).style.color = '#1e3a8a'; });
    content.querySelectorAll('pre,code').forEach((el) => { (el as HTMLElement).style.margin = '12px 0'; });
    content.querySelectorAll('table').forEach((table) => {
      const t = table as HTMLTableElement;
      t.style.width = '100%';
      t.style.borderCollapse = 'collapse';
      t.style.margin = '16px 0';
      t.style.fontSize = '12px';
      t.style.pageBreakInside = 'avoid';
      t.style.breakInside = 'avoid';
    });
    content.querySelectorAll('th').forEach((th) => { const h = th as HTMLTableCellElement; h.style.background = '#1e3a8a'; h.style.color = '#fff'; h.style.padding = '10px 12px'; h.style.textAlign = 'left'; h.style.fontWeight = '600'; h.style.border = '1px solid #e6eef8'; });
    content.querySelectorAll('td').forEach((td) => { const d = td as HTMLTableCellElement; d.style.padding = '10px 12px'; d.style.border = '1px solid #e6eef8'; d.style.verticalAlign = 'top'; });
    content.querySelectorAll('tr:nth-child(even)').forEach((tr) => { (tr as HTMLTableRowElement).style.background = '#f8fafc'; });

    // Remove footer/feedback remnants inside content
    content.querySelectorAll('[data-bi-name*="feedback"], [data-bi-name*="rating"], [class*="feedback"], [id*="feedback"], .thumb-rating-button, .xp-tag, hr, .links, .locale-selector-link, ul.links, .margin-block-xs').forEach((el) => el.remove());

    panel.appendChild(header);
    panel.appendChild(content);
    container.appendChild(panel);
  });

  // Place off-screen and render the whole document at once
  container.style.position = 'fixed';
  container.style.left = '-9999px';
  document.body.appendChild(container);

  // Wait for fonts and images to load so layout stabilizes before rendering
  try { await (document as any).fonts?.ready; } catch {}
  const imgs = Array.from(container.querySelectorAll('img')) as HTMLImageElement[];
  await Promise.all(imgs.map((img) => new Promise((res) => { if (img.complete) return res(null); img.onload = img.onerror = () => res(null); })));

  const canvas = await html2canvas(container, { scale: 2, useCORS: true, allowTaint: true });

  // Compute ratio and page height in px
  const ratio = canvas.width / pdfWidth;
  const rawPageHeightPx = Math.floor(pdfHeight * ratio);

  // Page margins in PDF points – keeps content away from top/bottom edges
  const PAGE_MARGIN_PT = 24; // ~8 mm
  const PAGE_MARGIN_PX = Math.round(PAGE_MARGIN_PT * ratio); // same margin in canvas px
  const pageHeightPx = rawPageHeightPx - 2 * PAGE_MARGIN_PX; // usable content height per page

  // ---------------------------------------------------------------------------
  // Compute protected element positions in CANVAS pixel coordinates.
  // html2canvas rendered at scale=2, so canvas coords = DOM offset * 2.
  // We use getBoundingClientRect which is reliable since the container is in
  // the DOM (position:fixed).  We subtract the container's own rect.top.
  // ---------------------------------------------------------------------------
  const h2cScale = 2; // must match the scale passed to html2canvas above
  const containerRect = container.getBoundingClientRect();

  const protectedSelectors = [
    // Block-level text elements – prevents slicing through the middle of a line
    'p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'dt', 'dd',
    // Larger compound elements
    'table', 'pre', 'code', 'img', 'blockquote',
    '.quiz-question', '.header', 'figure', 'svg',
  ];

  interface PRect { top: number; bottom: number; height: number; el: HTMLElement }
  const protectedRects: PRect[] = [];

  protectedSelectors.forEach((sel) => {
    container.querySelectorAll(sel).forEach((node) => {
      const el = node as HTMLElement;
      const r = el.getBoundingClientRect();
      const top  = Math.round((r.top  - containerRect.top) * h2cScale);
      const bottom = Math.round((r.bottom - containerRect.top) * h2cScale);
      if (bottom <= top) return; // skip zero‑height
      protectedRects.push({ top, bottom, height: bottom - top, el });
    });
  });

  // Deduplicate / merge overlapping rects (e.g. thead inside table)
  protectedRects.sort((a, b) => a.top - b.top || b.height - a.height);
  const merged: PRect[] = [];
  for (const r of protectedRects) {
    const last = merged[merged.length - 1];
    if (last && r.top >= last.top && r.bottom <= last.bottom) continue; // fully inside previous
    merged.push(r);
  }

  // For tables taller than a page, collect per‑row boundaries so we can split
  // between rows rather than through the middle of one.
  const tableRowMap = new Map<HTMLElement, { top: number; bottom: number }[]>();
  merged.forEach((r) => {
    if (r.el.tagName.toLowerCase() !== 'table') return;
    if (r.height < pageHeightPx) return; // fits on one page – no need
    const rows: { top: number; bottom: number }[] = [];
    r.el.querySelectorAll('tr').forEach((tr) => {
      const rr = (tr as HTMLElement).getBoundingClientRect();
      rows.push({
        top:    Math.round((rr.top    - containerRect.top) * h2cScale),
        bottom: Math.round((rr.bottom - containerRect.top) * h2cScale),
      });
    });
    if (rows.length) tableRowMap.set(r.el, rows);
  });

  // ---------------------------------------------------------------------------
  // Slice the canvas into PDF pages. CRITICAL RULE: a slice must NEVER exceed
  // pageHeightPx, otherwise the resulting image overflows the fixed‑size PDF
  // page and gets clipped.  When a protected element straddles the page
  // boundary we pull the boundary BACK (before the element) so that the
  // element starts cleanly on the next page.
  // ---------------------------------------------------------------------------
  const pagesInfo: { y: number; height: number }[] = [];
  let y = 0;
  const totalHeight = canvas.height;

  while (y < totalHeight) {
    let end = Math.min(y + pageHeightPx, totalHeight);

    // Iteratively pull `end` back to avoid splitting any protected rect.
    // We loop because pulling back for one rect might uncover another.
    let changed = true;
    let iters = 0;
    while (changed && iters < 20) {
      changed = false;
      iters++;
      for (const r of merged) {
        if (r.bottom <= y) continue;   // entirely above current page
        if (r.top >= end) break;       // beyond current boundary (sorted)
        if (r.bottom <= end) continue; // fully within current page – OK

        // r straddles the boundary (r.top < end && r.bottom > end)
        if (r.height <= pageHeightPx && r.top > y) {
          // Element fits on a page → cut BEFORE it so it moves to next page.
          end = r.top;
          changed = true;
          break; // re‑check from the start with new end
        }

        // Element is taller than a page (e.g. big table) → try row split
        if (tableRowMap.has(r.el)) {
          const rows = tableRowMap.get(r.el)!;
          let candidate = -1;
          for (const row of rows) {
            if (row.top >= end) break;
            if (row.bottom <= end && row.bottom > y) candidate = row.bottom;
          }
          if (candidate > y) {
            end = candidate;
            // don't set changed – we found a clean row boundary
          }
        }
        // Element doesn't fit and no rows → allow the split (unavoidable)
        break;
      }
    }

    // Don't create absurdly small pages (< 15 % of a full page).
    // If pulling back made the slice too small, revert to full page height
    // and accept the split.
    const minSlice = Math.round(pageHeightPx * 0.15);
    if (end - y < minSlice && y + pageHeightPx < totalHeight) {
      end = Math.min(y + pageHeightPx, totalHeight);
    }

    // Absolute safety: guarantee forward progress
    if (end <= y) end = Math.min(y + pageHeightPx, totalHeight);

    pagesInfo.push({ y, height: end - y });
    y = end;
  }

  // Total pages
  const totalPages = pagesInfo.length;
  let currentStep = 0;

  // Render each page slice into PDF
  for (let p = 0; p < pagesInfo.length; p++) {
    const { y: sliceY, height: sliceH } = pagesInfo[p];
    const pageCanvas = document.createElement('canvas');
    pageCanvas.width = canvas.width;
    pageCanvas.height = sliceH;
    const ctx = pageCanvas.getContext('2d');
    if (!ctx) throw new Error('Canvas context not available');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
    ctx.drawImage(canvas, 0, sliceY, canvas.width, sliceH, 0, 0, pageCanvas.width, pageCanvas.height);

    const imgData = pageCanvas.toDataURL('image/png');
    const imgProps = { width: pdfWidth, height: pageCanvas.height / ratio };
    if (p > 0) pdf.addPage();
    // Place image inset by the page margin so content never touches the edge
    pdf.addImage(imgData, 'PNG', 0, PAGE_MARGIN_PT, imgProps.width, imgProps.height);

    currentStep += 1;
    onProgress?.({ step: currentStep, total: totalPages });
  }

  // cleanup container
  document.body.removeChild(container);

  pdf.save(filename);
}

export default generatePDF;
