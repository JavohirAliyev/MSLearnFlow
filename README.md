# MS LearnFlow

A Chrome extension that lets you export Microsoft Learn modules and learning paths as a single, clean PDF — directly from your browser, with no accounts, no uploads, and no data leaving your device.

## What it does

Microsoft Learn is one of the best free resources for learning Azure, Microsoft 365, and other Microsoft technologies. The problem is that there is no built-in way to save a module or learning path as a document you can read offline, annotate, or share with your team.

MS LearnFlow solves that. Open any Microsoft Learn module or learning path, click the extension icon, select the units you want, and download a PDF. Everything happens locally in your browser. Nothing is sent to a server.

## Key features

- Export any Microsoft Learn module or full learning path to PDF
- Select individual units or export an entire course in one click
- All rendering happens inside the browser — no server, no account required
- No data collection of any kind — see the [Privacy Policy](PRIVACY_POLICY.md)
- Works on `learn.microsoft.com`

## Installation

The extension is available on the Chrome Web Store. Search for **MS LearnFlow**, or install it manually:

1. Download or clone this repository
2. Run `npm install` and then `npm run build`
3. Open Chrome and go to `chrome://extensions`
4. Enable Developer mode
5. Click "Load unpacked" and select the `dist` folder

## How to use it

1. Navigate to any module or learning path on [learn.microsoft.com](https://learn.microsoft.com)
2. Click the MS LearnFlow icon in the Chrome toolbar
3. The extension will detect the content and list the available units
4. Select the units you want to include
5. Click "Export PDF" — the file will download automatically

## Tech stack

- React 18 and TypeScript
- Vite for bundling
- jsPDF and html2canvas for PDF generation
- Chrome Manifest V3

## Privacy

MS LearnFlow does not collect, store, or transmit any user data. All processing is performed locally inside your browser. Temporary export progress is kept in `chrome.storage.session` and is cleared when the browser session ends. See [PRIVACY_POLICY.md](PRIVACY_POLICY.md) for full details.

## License

MIT
