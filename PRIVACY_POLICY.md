# Privacy Policy for MS LearnFlow

**Last updated: March 11, 2026**

## Overview

MS LearnFlow is a Chrome extension that allows users to export selected units from Microsoft Learn modules as a single PDF document. This privacy policy explains how the extension handles user data.

## Data Collection

**MS LearnFlow does not collect, store, transmit, or share any personal data.** The extension has no analytics, telemetry, tracking, or external data-collection mechanisms of any kind.

## How the Extension Works

- The extension fetches publicly available content from `https://learn.microsoft.com` when you initiate a PDF export.
- All HTML parsing, cleaning, and PDF rendering happens **entirely within your browser** using Chrome's offscreen document API.
- Generated PDF files are saved to your local device through Chrome's built-in download functionality.
- Temporary job state (e.g., export progress) is stored in `chrome.storage.session`, which is automatically cleared when the browser session ends. No data persists beyond the active session.

## Permissions Used

| Permission | Purpose |
|---|---|
| `activeTab` | Read the current Microsoft Learn page URL to detect module/unit structure. |
| `scripting` | Inject a content script to extract page content for PDF generation. |
| `tabs` | Access the active tab's URL to determine the Microsoft Learn content type. |
| `storage` | Store temporary job progress in session storage (cleared when the browser closes). |
| `offscreen` | Create a hidden document to parse HTML and render PDFs in the background. |
| `downloads` | Save the generated PDF file to your device. |

## Host Permissions

The extension only requests access to `https://learn.microsoft.com/*`. It does not access any other websites.

## Third-Party Services

MS LearnFlow does **not** communicate with any third-party servers, APIs, or services. All processing is performed locally on your device.

## Data Sharing

No data is shared with the developer, any third parties, or any external servers.

## Changes to This Policy

If this privacy policy is updated, the changes will be reflected in the extension's listing and this document. The "Last updated" date at the top will be revised accordingly.

## Contact

If you have questions about this privacy policy, please open an issue on the extension's GitHub repository.
