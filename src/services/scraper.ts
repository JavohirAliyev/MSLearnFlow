/*
  scraper.ts
  - Fetches Microsoft Learn module metadata via the Catalog API (best-effort)
  - Fetches unit pages and cleans them using DOMParser
  - Inlines CSS and converts <img> sources to base64 data URLs
  - Understands the full MS Learn content hierarchy:
      Course → Learning Paths → Modules → Units

  Note: The Microsoft Learn Catalog API surface can vary; this implementation
  uses a flexible "search by path" approach and attempts to find the module
  and its child units. You may need to adapt the parsing to the exact API
  responses you observe in production.
*/

// ─── Hierarchy types ────────────────────────────────────────────────────────

export type ContentType = 'course' | 'learning-path' | 'module' | 'unit' | 'unknown';

export interface UnitNode {
  title: string;
  url: string;
}

export interface ModuleNode {
  uid?: string;
  title: string;
  url: string;
  units: UnitNode[];
  unitsLoaded: boolean;
}

export interface LearningPathNode {
  uid?: string;
  title: string;
  url: string;
  modules: ModuleNode[];
  modulesLoaded: boolean;
}

export interface CourseNode {
  title: string;
  url: string;
  learningPaths: LearningPathNode[];
}

// ─── Legacy types (kept for PDF generation pipeline) ────────────────────────

export type ModuleUnit = {
  title: string;
  url: string;
};

export type ModuleMetadata = {
  title: string | null;
  units: ModuleUnit[];
  debug?: any;
};

// ─── Content-type detection ──────────────────────────────────────────────────

/** Detect the MS Learn content type from a URL. */
export function detectContentType(url: string): ContentType {
  try {
    const { pathname } = new URL(url);
    if (/\/training\/courses\//i.test(pathname)) return 'course';
    if (/\/training\/paths\//i.test(pathname)) return 'learning-path';
    // Unit pages: /training/modules/{module-slug}/{unit-slug}  (has two path segments after "modules")
    if (/\/training\/modules\/[^/]+\/[^/]+/i.test(pathname)) return 'unit';
    if (/\/training\/modules\//i.test(pathname)) return 'module';
    return 'unknown';
  } catch {
    return 'unknown';
  }
}

/** Strip the common "learn.wwl." / "learn.azure." etc. UID prefix to get the slug. */
function uidToSlug(uid: string): string {
  return uid.replace(/^learn\.[^.]+\./, '');
}

/** Friendly title from a slug (fallback before real title is loaded). */
function slugToTitle(slug: string): string {
  return slug
    .replace(/-/g, ' ')
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

/** For a unit URL, return the parent module URL (strip the last path segment). */
export function moduleUrlFromUnitUrl(unitUrl: string): string {
  try {
    const u = new URL(unitUrl);
    const parts = u.pathname.replace(/\/$/, '').split('/');
    parts.pop(); // remove unit slug
    u.pathname = parts.join('/') + '/';
    return u.toString();
  } catch {
    return unitUrl;
  }
}

// ─── Structure fetchers ──────────────────────────────────────────────────────

/** Fetch a Course page and return the top-level CourseNode with LP stubs. */
export async function fetchCourseStructure(url: string): Promise<CourseNode> {
  const res = await fetch(url);
  const text = await res.text();
  const parser = new DOMParser();
  const doc = parser.parseFromString(text, 'text/html');

  const h1 = doc.querySelector('h1')?.textContent?.trim();
  const metaTitle = doc
    .querySelector('meta[property="og:title"]')
    ?.getAttribute('content')
    ?.replace(/\s*[-|]\s*Training.*$/i, '')
    .trim();
  const title = h1 || metaTitle || 'Course';

  // <meta name="learn_item" content="learn.wwl.{slug}"> tags list the LP UIDs
  const lpMetas = Array.from(doc.querySelectorAll('meta[name="learn_item"]'));
  const learningPaths: LearningPathNode[] = lpMetas.map((meta) => {
    const uid = meta.getAttribute('content') || '';
    const slug = uidToSlug(uid);
    return {
      uid,
      title: slugToTitle(slug), // placeholder; real title loaded on expand
      url: `https://learn.microsoft.com/en-us/training/paths/${slug}/`,
      modules: [],
      modulesLoaded: false,
    };
  });

  return { title, url, learningPaths };
}

/** Fetch a Learning-Path page and return the LearningPathNode with Module stubs. */
export async function fetchLearningPathStructure(url: string): Promise<LearningPathNode> {
  const res = await fetch(url);
  const text = await res.text();
  const parser = new DOMParser();
  const doc = parser.parseFromString(text, 'text/html');

  const h1 = doc.querySelector('h1.title, h1')?.textContent?.trim();
  const metaTitle = doc
    .querySelector('meta[property="og:title"]')
    ?.getAttribute('content')
    ?.replace(/\s*[-|]\s*Training.*$/i, '')
    .trim();
  const title = h1 || metaTitle || 'Learning Path';
  const uid = doc.querySelector('meta[name="uid"]')?.getAttribute('content') ?? undefined;

  // Each module is a <div data-bi-name="module"> block
  const moduleBoxes = Array.from(doc.querySelectorAll('[data-bi-name="module"]'));
  const modules: ModuleNode[] = moduleBoxes
    .map((box) => {
      // The module title is in a semibold anchor that links to the module page
      const titleLink = box.querySelector(
        'a[href*="modules/"].font-weight-semibold, a[href*="modules/"][class*="font-weight-semibold"], .column.is-auto a[href*="modules/"]'
      ) as HTMLAnchorElement | null;
      // Fallback: any anchor with modules/ in href
      const anyLink = titleLink ||
        (box.querySelector('a[href*="modules/"]') as HTMLAnchorElement | null);
      const moduleTitle =
        (titleLink?.textContent || anyLink?.textContent || '').trim().replace(/\s+/g, ' ') ||
        'Module';
      const href = anyLink?.getAttribute('href') || '';
      let moduleUrl = '';
      try {
        moduleUrl = href ? new URL(href, url).toString() : '';
      } catch {
        moduleUrl = href;
      }
      const boxId = (box as HTMLElement).id || undefined;
      return {
        uid: boxId,
        title: moduleTitle,
        url: moduleUrl,
        units: [],
        unitsLoaded: false,
      };
    })
    .filter((m) => !!m.url);

  return { uid, title, url, modules, modulesLoaded: true };
}

/** Fetch a Module page (via Catalog API + HTML fallback) and return ModuleNode with Units. */
export async function fetchModuleStructure(url: string): Promise<ModuleNode> {
  const meta = await fetchModuleCatalog(url);
  const title = meta?.title || 'Module';
  const units: UnitNode[] = (meta?.units || []).map((u) => ({
    title: u.title,
    url: u.url,
  }));
  return { title, url, units, unitsLoaded: true };
}

// Attempt to extract the module slug/path from the full URL
export function parseModulePathFromUrl(url: string): string {
  try {
    const u = new URL(url);
    // For Learn URLs, module pages often look like: /en-us/learn/modules/<slug>/
    // We'll remove the locale prefix if present and return the path.
    const parts = u.pathname.split('/').filter(Boolean);
    // If path contains 'modules' or 'learning', return from that segment onward
    const idx = parts.findIndex((p) => p === 'modules' || p === 'learning' || p === 'path');
    if (idx >= 0) return '/' + parts.slice(idx).join('/');
    return u.pathname;
  } catch (e) {
    return url;
  }
}

// Fetch module metadata from the Catalog API using the path as a search term.
export async function fetchModuleCatalog(pathOrUrl: string): Promise<ModuleMetadata | null> {
  // Prefer searching the catalog by the module path (e.g. /en-us/training/modules/<slug>)
  const path = parseModulePathFromUrl(pathOrUrl);
  const search = encodeURIComponent(path.replace(/^\//, ''));
  const api = `https://learn.microsoft.com/api/catalog?search=${search}&language=en-US`;
    try {
    const res = await fetch(api);
    const resOk = res.ok;
    let json: any = null;
    try {
      json = await res.json();
    } catch (e) {
      // ignore non-JSON responses
    }
    if (!resOk) {
      // continue to HTML fallback below
    }

    // The Catalog API responses differ; try common shapes:
    //  - json.items[] with type/module and children
    //  - json.results
    const items = (json && (json.items || json.results)) || json;
    // resolved items info omitted in production
    if (!Array.isArray(items)) {
      // don't return early; we'll attempt HTML parsing fallback later
    }

    // Find the best module-like item
    let moduleItem: any = null;
    if (Array.isArray(items)) {
      moduleItem = items.find((it: any) => {
        const t = (it.type || it.content_type || it.contentType || '').toLowerCase();
        return t.includes('module') || t.includes('learning') || t.includes('path') || t.includes('module_collection');
      }) || items[0];
    } else if (items && typeof items === 'object') {
      // API sometimes returns a single object instead of an array
      moduleItem = items;
    }

    if (!moduleItem) {
      // no module-like item found
    }

    const title = (moduleItem && (moduleItem.title || moduleItem.name)) || 'Module';

    // Try to extract child units
    let units: ModuleUnit[] = [];
    // Some responses include a `items` or `children` array
    const children = moduleItem.children || moduleItem.items || moduleItem.units || moduleItem.contents || [];
    if (Array.isArray(children) && children.length > 0) {
      units = children.map((c: any) => ({
        title: c.title || c.name || c.displayName || 'Unit',
        url: c.url || c.site_url || c.href || (c.path ? `https://learn.microsoft.com${c.path}` : '')
      })).filter((u: ModuleUnit) => !!u.url);
    }

    // If no children found, try to extract from search hits
    if (units.length === 0 && Array.isArray(items)) {
      units = items.map((it: any) => ({
        title: it.title || it.name || 'Unit',
        url: it.url || it.site_url || it.href || (it.path ? `https://learn.microsoft.com${it.path}` : '')
      })).filter((u: ModuleUnit) => !!u.url);
      // units from search hits
    }

    // If still no units (or catalog returned nothing useful), try fetching the page HTML
    // and parse the unit list directly (robust for module pages served as HTML).
    if (units.length === 0) {
      try {
        const pageRes = await fetch(pathOrUrl);
        if (pageRes.ok) {
          const txt = await pageRes.text();
          const parser = new DOMParser();
          const doc = parser.parseFromString(txt, 'text/html');

          // Try common title locations
          const h1 = doc.querySelector('h1.title')?.textContent?.trim() || doc.querySelector('h1')?.textContent?.trim();
          const metaTitle = doc.querySelector('meta[property="og:title"]')?.getAttribute('content') || doc.querySelector('title')?.textContent;
          const finalTitle = (h1 || metaTitle || title || '').trim();

          const unitAnchors = Array.from(doc.querySelectorAll('#unit-list a.unit-title, #unit-list a')) as HTMLAnchorElement[];
          const parsedUnits: ModuleUnit[] = unitAnchors.map((a) => {
            const href = a.getAttribute('href') || a.href || '';
            try {
              const url = new URL(href, pathOrUrl).toString();
              return { title: (a.textContent || 'Unit').trim(), url };
            } catch (e) {
              return { title: (a.textContent || 'Unit').trim(), url: href };
            }
          }).filter((u) => !!u.url);

          // parsedUnits from HTML
          if (parsedUnits.length > 0) {
            return { title: finalTitle || title, units: parsedUnits };
          }
        }
      } catch (e) {
        // ignore page-parse failures, we'll fall through to returning what we have
      }
    }

    // returning title and units counts
    return { title, units };
  } catch (err) {
    // Catalog fetch failed
    return { title: null, units: [], debug: { error: String(err) } };
  }
}

async function toDataUrl(src: string): Promise<string> {
  try {
    const r = await fetch(src);
    const blob = await r.blob();
    return await new Promise<string>((resolve, reject) => {
      const fr = new FileReader();
      fr.onloadend = () => resolve(fr.result as string);
      fr.onerror = reject;
      fr.readAsDataURL(blob);
    });
  } catch (e) {
    // Failed to inline image – leave original src
    return src;
  }
}

// Clean a unit page's HTML: extract main content, remove navs/feedback, inline assets, inline styles
export async function fetchAndCleanUnit(url: string): Promise<{ title: string; url: string; html: string } | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const text = await res.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, 'text/html');

    let title = doc.querySelector('title')?.textContent || url;
    // Clean up the title - remove common suffixes
    title = title.replace(/\s*[-|]\s*Training\s*\|?\s*Microsoft Learn/gi, '').trim();
    title = title.replace(/\s*[-|]\s*Microsoft Learn/gi, '').trim();
    title = title.replace(/\s*[-|]\s*Microsoft Docs/gi, '').trim();

    // Detect if this is an assessment/quiz page
    const isAssessment = url.includes('knowledge-check') || 
                         url.includes('assessment') || 
                         url.includes('check-your-knowledge') ||
                         doc.querySelector('meta[name="module_assessment"][content="true"]') !== null ||
                         doc.querySelector('#question-container') !== null;

    let clone: HTMLElement;

    if (isAssessment) {
      // For assessment pages, build content from quiz elements specifically
      clone = document.createElement('div');
      
      // Get the quiz title
      const quizTitle = doc.querySelector('#quiz-title');
      if (quizTitle) {
        const titleClone = quizTitle.cloneNode(true) as HTMLElement;
        titleClone.removeAttribute('hidden');
        titleClone.removeAttribute('aria-hidden');
        titleClone.style.fontSize = '20px';
        titleClone.style.fontWeight = '600';
        titleClone.style.marginBottom = '16px';
        titleClone.style.color = '#1e3a8a';
        clone.appendChild(titleClone);
      }

      // Get all quiz questions
      const questions = doc.querySelectorAll('.quiz-question');
      questions.forEach((q, idx) => {
        const questionDiv = document.createElement('div');
        questionDiv.style.marginBottom = '24px';
        questionDiv.style.padding = '16px';
        questionDiv.style.background = '#f8fafc';
        questionDiv.style.borderRadius = '8px';
        questionDiv.style.border = '1px solid #e2e8f0';

        // Get question text
        const questionLabel = q.querySelector('.field-label');
        if (questionLabel) {
          const questionText = document.createElement('div');
          questionText.style.fontWeight = '600';
          questionText.style.marginBottom = '12px';
          questionText.style.color = '#1e3a8a';
          questionText.innerHTML = questionLabel.innerHTML;
          questionDiv.appendChild(questionText);
        }

        // Get answer choices
        const choices = q.querySelectorAll('.quiz-choice');
        choices.forEach((choice) => {
          const choiceDiv = document.createElement('div');
          choiceDiv.style.padding = '10px 14px';
          choiceDiv.style.margin = '6px 0';
          choiceDiv.style.background = '#fff';
          choiceDiv.style.borderRadius = '6px';
          choiceDiv.style.border = '1px solid #e2e8f0';
          
          // Get the text content from the label, excluding input elements
          const labelText = choice.querySelector('.radio-label-text, .checkbox-label-text');
          if (labelText) {
            choiceDiv.innerHTML = labelText.innerHTML;
          } else {
            // Fallback: get text content directly
            const textContent = (choice.textContent || '').trim();
            choiceDiv.textContent = textContent;
          }
          
          questionDiv.appendChild(choiceDiv);
        });

        clone.appendChild(questionDiv);
      });

      // If no questions found, try to get module-assessment form
      if (questions.length === 0) {
        const assessmentForm = doc.querySelector('#module-assessment-questions-form');
        if (assessmentForm) {
          const formClone = assessmentForm.cloneNode(true) as HTMLElement;
          formClone.removeAttribute('hidden');
          formClone.removeAttribute('aria-hidden');
          clone.appendChild(formClone);
        }
      }
    } else {
      // For regular pages, target the most specific content container
      const main = doc.querySelector('#module-unit-content') ||
                   doc.querySelector('#unit-inner-section') ||
                   doc.querySelector('.unit-section') ||
                   doc.querySelector('#main-column') ||
                   doc.querySelector('main') ||
                   doc.body;
      clone = main ? main.cloneNode(true) as HTMLElement : document.createElement('div');
    }

    // Remove known noisy selectors and interactive/style-breaking elements (MS Learn specific)
    const removeSelectors = [
      // MS Learn specific navigation and UI elements
      '#unit-nav-dropdown',
      '#module-menu',
      '#article-header',
      '#article-header-breadcrumbs',
      '#article-header-page-actions',
      '#module-unit-metadata',
      '#completion-nav',
      '#next-section',
      '#ms--unit-user-feedback',
      '#site-user-feedback-footer',
      '#ms--unit-support-message',
      '#module-unit-notification-container',
      '#module-unit-module-assessment-message-container',
      '.feedback-section',
      '.xp-tag',
      '.xp-tag-hexagon',
      'bread-crumbs',
      '.popover',
      '.dropdown',
      '.dropdown-menu',
      '.dropdown-trigger',
      '[data-bi-name="unit-menu"]',
      '[data-bi-name="module-nav"]',
      '[data-ask-learn-modal-entry]',
      '[data-ask-learn-flyout-entry]',
      // Generic navigation and chrome
      'header',
      'nav',
      'aside',
      'footer',
      '.next-prev',
      '.next-previous',
      // Feedback elements
      '.feedback',
      'form.feedback',
      '.feedback-form',
      '.ms-Feedback',
      '.survey',
      '.rating',
      '.vote',
      '.comment',
      '.thumb-rating-button',
      '[data-binary-rating-response]',
      // Interactive elements that break styling
      '.az-try-it',
      '.try-it',
      '.interactive',
      '.sandbox',
      'button',
      'input',
      'select',
      'textarea',
      'label',
      'form',
      // Actions and controls
      '.actions',
      '.action-bar',
      '.buttons',
      '.controls',
      // Skip links and hidden elements
      '.skip-to-main-content',
      '.visually-hidden',
      '.visually-hidden-until-focused',
      '.sr-only',
      '.skip-link',
      '.skipnav',
      '.skip-navigation',
      '.skip',
      '.hidden',
      '.hide',
      '.d-none',
      '.display-none',
      '.display-none-print',
      '[hidden]',
      '[aria-hidden="true"]',
      // Breadcrumbs and navigation
      '.breadcrumbs',
      '.global-nav',
      '.side-nav',
      '.toc',
      '.toc-nav',
      '.nav-buttons',
      '.nav-controls',
      '.pagination',
      // Banners and ads
      '.banner',
      '.ad',
      '.ads',
      '.cookie-banner',
      '.consent-banner',
      '.newsletter',
      '.newsletter-signup',
      // Related and external content
      '.related-content',
      '.related-links',
      '.external-links',
      '.footer-links',
      // Print-specific classes
      '.no-print',
      '.print-only',
      '.print-hide',
      '.print-button',
      // Page ending and section styling elements
      '.modular-content-container.margin-block-xs',
      '.section.is-full-height',
      '.is-hidden-interactive',
      '.layout-body-footer',
      '.footer-layout',
      '.uhf-container',
      '[data-bi-name="footer"]',
      '[data-bi-name="footerlinks"]',
      '[data-bi-name="layout-footer"]',
      '.locale-selector-link',
      '.ccpa-privacy-link',
      '.theme-dropdown-trigger',
      '.dropdown.has-caret-up',
      '.links',
      'hr.hr',
      'hr',
      // MS Learn specific ending elements
      '.is-uniform.position-relative > .xp-tag',
      '.margin-top-md.display-none-print',
      '[data-test-id="site-user-feedback-footer"]',
      '.font-weight-semibold.display-none',
      // Feedback section parents
      '.margin-block-xs',
      '[class*="feedback"]',
      '[id*="feedback"]',
      // Empty decorative sections
      '.box:empty',
      '.section:empty',
      'div.modular-content-container:empty',
    ];
    removeSelectors.forEach((sel) => {
      try {
        clone.querySelectorAll(sel).forEach((n) => n.remove());
      } catch (e) {
        // Ignore invalid selectors
      }
    });

    // Handle interactive elements - for assessment pages, we already extracted text above
    if (isAssessment) {
      // Just remove any remaining interactive elements
      clone.querySelectorAll('button, input, select, textarea').forEach((n) => n.remove());
    } else {
      // For non-assessment pages, remove all interactive elements
      clone.querySelectorAll('button, input, select, textarea, label, form').forEach((n) => n.remove());
    }

    // Remove empty list items (like "1." "2." navigation remnants)
    clone.querySelectorAll('li').forEach((li) => {
      const text = (li.textContent || '').trim();
      // Remove if it's just a number, empty, or very short navigation text
      if (!text || /^\d+\.?$/.test(text) || text.length < 3) {
        li.remove();
      }
    });

    // Remove empty uls and ols after cleaning
    clone.querySelectorAll('ul, ol').forEach((list) => {
      if (list.children.length === 0) {
        list.remove();
      }
    });

    // Remove any elements with data attributes related to feedback or navigation
    clone.querySelectorAll('[data-bi-name*="feedback"], [data-bi-name*="rating"], [data-bi-name*="nav"]').forEach((n) => n.remove());

    // Replace interactive sandboxes with a placeholder link
    clone.querySelectorAll('.sandbox, .interactive, .az-try-it, .try-it').forEach((node) => {
      const a = doc.createElement('a');
      a.textContent = 'Link to Lab';
      a.href = url;
      a.style.display = 'block';
      a.style.padding = '12px';
      a.style.background = '#f5f5f5';
      a.style.border = '1px solid #ddd';
      node.replaceWith(a);
    });

    // Inline images
    const imgs = Array.from(clone.querySelectorAll('img')) as HTMLImageElement[];
    await Promise.all(imgs.map(async (img) => {
      const src = img.src;
      if (!src) return;
      try {
        const data = await toDataUrl(src);
        img.src = data;
      } catch (e) {
        // leave original src if conversion fails
      }
    }));

    // DO NOT inline MS Learn CSS - it causes page-ending styles to appear mid-document
    // Instead, use minimal clean styles that we control
    const combinedStyles = '';

    // Preserve code block inline styles where possible (fallback styles included)
    Array.from(clone.querySelectorAll('pre, code')).forEach((el) => {
      // Provide a safe inline style so html2canvas keeps background/colors
      (el as HTMLElement).style.background = (el as HTMLElement).style.background || '#f6f8fa';
      (el as HTMLElement).style.color = (el as HTMLElement).style.color || '#24292e';
      (el as HTMLElement).style.fontFamily = (el as HTMLElement).style.fontFamily || 'ui-monospace, SFMono-Regular, Menlo, Monaco, "Roboto Mono", "Segoe UI Mono", "Helvetica Neue", monospace';
      (el as HTMLElement).style.padding = (el as HTMLElement).style.padding || '8px';
      (el as HTMLElement).style.borderRadius = (el as HTMLElement).style.borderRadius || '6px';
      (el as HTMLElement).style.overflow = 'auto';
    });

    // Remove any remaining page-ending elements by inspecting content
    clone.querySelectorAll('section, div').forEach((el) => {
      const text = (el.textContent || '').trim().toLowerCase();
      // Remove feedback/footer sections by content
      if (text.startsWith('was this page helpful') ||
          text.startsWith('feedback') ||
          text === 'feedback' ||
          (text.includes('previous versions') && text.includes('blog')) ||
          (text.includes('privacy') && text.includes('terms of use')) ||
          text.includes('© microsoft') ||
          text.includes('your privacy choices')) {
        el.remove();
      }
    });

    // Remove <hr> elements (usually section dividers before footers)
    clone.querySelectorAll('hr').forEach((hr) => hr.remove());

    // Remove any elements with footer-like styling classes
    clone.querySelectorAll('[class*="footer"], [class*="Footer"], [class*="feedback"], [id*="feedback"]').forEach((el) => el.remove());

    // Remove elements containing footer links pattern
    clone.querySelectorAll('ul.links, [data-bi-name="footerlinks"]').forEach((el) => el.remove());

    // Build final HTML fragment with clean, minimal styles
    const wrapper = document.createElement('div');
    const styleEl = document.createElement('style');
    styleEl.textContent = `
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial; color: #111; }
      * { box-sizing: border-box; }
      img { max-width: 100%; height: auto; }
      pre, code { background: #f6f8fa; padding: 8px; border-radius: 6px; overflow: auto; font-family: ui-monospace, monospace; }
      table { width: 100%; border-collapse: collapse; margin: 16px 0; }
      th { background: #1e3a8a; color: #fff; padding: 10px 12px; text-align: left; border: 1px solid #e6eef8; }
      td { padding: 10px 12px; border: 1px solid #e6eef8; vertical-align: top; }
      tr:nth-child(even) { background: #f8fafc; }
      ul, ol { margin: 12px 0 12px 28px; }
      p { margin: 12px 0; }
      h1, h2, h3, h4, h5 { margin: 20px 0 10px 0; color: #1e3a8a; }
    `;
    wrapper.appendChild(styleEl);
    wrapper.appendChild(clone);

    // Serialize
    const html = wrapper.innerHTML;
    return { title: title || url, url, html };
  } catch (err) {
    // Failed to fetch/clean unit
    return null;
  }
}
