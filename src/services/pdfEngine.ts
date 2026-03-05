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
  onProgress?: (p: { step: number; total: number }) => void
) {
  // Prepare PDF sizes in points
  const pdf = new jsPDF({ unit: 'pt', format: 'a4', compress: true });
  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = pdf.internal.pageSize.getHeight();

  // Render each segment individually to canvases so each unit starts on a new PDF page.
  const rendered: { canvas: HTMLCanvasElement; pages: number }[] = [];

  // Prepare a style element to inject popup CSS for consistent styling
  let popupCss: string | null = null;
  try {
    const res = await fetch(chrome.runtime.getURL('src/popup/popup.css'));
    if (res.ok) popupCss = await res.text();
  } catch {}

  for (let i = 0; i < segments.length; i++) {
    const seg = segments[i];

    // Create off-DOM container for this segment
    const container = document.createElement('div');
    container.style.width = '794px'; // visual width for rendering (approx A4 @ 96dpi)
    container.style.background = '#f4f7fb';
    container.style.color = '#111827';
    container.style.padding = '0';
    container.style.boxSizing = 'border-box';
    container.style.fontFamily = "'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif";
    container.style.fontSize = '13px';
    container.style.lineHeight = '1.45';
    container.style.maxWidth = '734px';

    // Inject popup CSS for consistent styling if available
    if (popupCss) {
      const style = document.createElement('style');
      style.textContent = popupCss;
      container.appendChild(style);
    }

    // Panel wrapper for segment
    const panel = document.createElement('div');
    panel.className = 'panel';
    panel.style.margin = '32px auto';
    panel.style.background = '#fff';
    panel.style.borderRadius = '10px';
    panel.style.boxShadow = '0 6px 18px rgba(16,24,40,0.06)';
    panel.style.border = '1px solid #e6eef8';
    panel.style.padding = '28px 30px';
    panel.style.maxWidth = '734px';

    // Header for unit (always at the top, visually separated)
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

    // Insert content
    const content = document.createElement('div');
    content.className = 'unit';
    content.innerHTML = seg.html || '';

    // Check if this is an assessment/quiz page
    const isAssessment = seg.url.includes('knowledge-check') || 
                         seg.url.includes('assessment') || 
                         seg.url.includes('check-your-knowledge') ||
                         seg.title.toLowerCase().includes('knowledge check') ||
                         seg.title.toLowerCase().includes('assessment');

    // For assessment pages, style quiz questions nicely
    if (isAssessment) {
      content.querySelectorAll('.question, [class*="question"], [class*="quiz"]').forEach((q) => {
        (q as HTMLElement).style.background = '#f8fafc';
        (q as HTMLElement).style.padding = '16px';
        (q as HTMLElement).style.borderRadius = '8px';
        (q as HTMLElement).style.marginBottom = '16px';
        (q as HTMLElement).style.border = '1px solid #e6eef8';
      });
    }

    // Remove any h1/h2/h3 headers at the very top of the content (to avoid duplicate headers)
    while (content.firstChild && content.firstChild.nodeType === 1 &&
      ['H1', 'H2', 'H3'].includes((content.firstChild as HTMLElement).tagName)) {
      content.removeChild(content.firstChild);
    }

    // Small niceties for readability
    content.style.color = '#0b1220';
    content.style.marginBottom = '0';
    content.querySelectorAll('img').forEach((img) => {
      (img as HTMLImageElement).style.maxWidth = '100%';
      (img as HTMLImageElement).style.height = 'auto';
      (img as HTMLImageElement).style.borderRadius = '6px';
      (img as HTMLImageElement).style.boxShadow = '0 2px 8px rgba(16,24,40,0.08)';
      (img as HTMLImageElement).style.margin = '12px 0';
    });
    // Add spacing between paragraphs
    content.querySelectorAll('p').forEach((p) => {
      (p as HTMLElement).style.margin = '12px 0';
    });
    // Add spacing for lists
    content.querySelectorAll('ul,ol').forEach((el) => {
      (el as HTMLElement).style.margin = '12px 0 12px 28px';
    });
    // Add spacing for headers
    content.querySelectorAll('h1,h2,h3,h4,h5').forEach((h) => {
      (h as HTMLElement).style.margin = '20px 0 10px 0';
      (h as HTMLElement).style.color = '#1e3a8a';
    });
    // Add spacing for code blocks
    content.querySelectorAll('pre,code').forEach((el) => {
      (el as HTMLElement).style.margin = '12px 0';
    });

    // Style tables for proper rendering and prevent page breaks
    content.querySelectorAll('table').forEach((table) => {
      const t = table as HTMLTableElement;
      t.style.width = '100%';
      t.style.borderCollapse = 'collapse';
      t.style.margin = '16px 0';
      t.style.fontSize = '12px';
      t.style.pageBreakInside = 'avoid';
      t.style.breakInside = 'avoid';
    });
    content.querySelectorAll('th').forEach((th) => {
      const h = th as HTMLTableCellElement;
      h.style.background = '#1e3a8a';
      h.style.color = '#fff';
      h.style.padding = '10px 12px';
      h.style.textAlign = 'left';
      h.style.fontWeight = '600';
      h.style.border = '1px solid #e6eef8';
    });
    content.querySelectorAll('td').forEach((td) => {
      const d = td as HTMLTableCellElement;
      d.style.padding = '10px 12px';
      d.style.border = '1px solid #e6eef8';
      d.style.verticalAlign = 'top';
    });
    content.querySelectorAll('tr:nth-child(even)').forEach((tr) => {
      (tr as HTMLTableRowElement).style.background = '#f8fafc';
    });

    // Remove any remaining feedback or navigation elements that slipped through
    content.querySelectorAll('[data-bi-name*="feedback"], [data-bi-name*="rating"], [class*="feedback"], [id*="feedback"], .thumb-rating-button, .xp-tag').forEach((el) => el.remove());

    // Remove page-ending elements that may have slipped through
    content.querySelectorAll('hr, [class*="footer"], [class*="Footer"], .links, .locale-selector-link, ul.links, .margin-block-xs').forEach((el) => el.remove());

    // Remove sections that look like page endings by their content
    content.querySelectorAll('section, div').forEach((el) => {
      const text = (el.textContent || '').trim().toLowerCase();
      if (text.startsWith('was this page helpful') ||
          text.startsWith('feedback') ||
          text === 'feedback' ||
          (text.includes('privacy') && text.includes('terms of use')) ||
          text.includes('© microsoft') ||
          text.includes('your privacy choices') ||
          (text.includes('previous versions') && text.includes('blog')) ||
          (text.includes('ai disclaimer') && text.includes('contribute'))) {
        el.remove();
      }
    });

    // Remove empty elements
    content.querySelectorAll('div, span, p').forEach((el) => {
      if (!el.textContent?.trim() && !el.querySelector('img, table, pre, code')) {
        el.remove();
      }
    });

    // Clean up any remaining style elements that might have MS Learn styles
    content.querySelectorAll('style').forEach((el) => el.remove());

    panel.appendChild(header);
    panel.appendChild(content);
    container.appendChild(panel);

    // Place off-screen and render
    container.style.position = 'fixed';
    container.style.left = '-9999px';
    document.body.appendChild(container);

    const canvas = await html2canvas(container, { scale: 2, useCORS: true, allowTaint: true });

    // Compute ratio and how many PDF pages this canvas will occupy
    const ratio = canvas.width / pdfWidth;
    const pageHeightPx = Math.floor(pdfHeight * ratio);
    const pages = Math.max(1, Math.ceil(canvas.height / pageHeightPx));

    rendered.push({ canvas, pages });

    // cleanup container
    document.body.removeChild(container);
  }

  // Total pages across all segments
  const totalPages = rendered.reduce((s, r) => s + r.pages, 0);
  let currentStep = 0;

  // Add each rendered canvas to PDF, slicing per page as needed
  for (let i = 0; i < rendered.length; i++) {
    const { canvas, pages } = rendered[i];
    const ratio = canvas.width / pdfWidth;
    const pageHeightPx = Math.floor(pdfHeight * ratio);

    for (let p = 0; p < pages; p++) {
      const y = p * pageHeightPx;
      const pageCanvas = document.createElement('canvas');
      pageCanvas.width = canvas.width;
      pageCanvas.height = Math.min(pageHeightPx, canvas.height - y);
      const ctx = pageCanvas.getContext('2d');
      if (!ctx) throw new Error('Canvas context not available');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
      ctx.drawImage(canvas, 0, y, canvas.width, pageCanvas.height, 0, 0, pageCanvas.width, pageCanvas.height);

      const imgData = pageCanvas.toDataURL('image/png');
      const imgProps = { width: pdfWidth, height: pageCanvas.height / ratio };
      if (currentStep > 0) pdf.addPage();
      pdf.addImage(imgData, 'PNG', 0, 0, imgProps.width, imgProps.height);

      currentStep += 1;
      onProgress?.({ step: currentStep, total: totalPages });
    }
  }

  pdf.save(filename);
}

export default generatePDF;
