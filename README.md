# OpenCreative

**A free, independent, open-source creative workspace.**

OpenCreative is an early browser-based MVP that brings together practical starter tools for image drawing, vector graphics, video preview, audio recording and PDF viewing. The long-term goal is a modular creative suite spanning the wider creative workflow.

> OpenCreative is an independent project. It is not Adobe software, is not affiliated with Adobe, and does not contain Adobe source code or proprietary assets.

## What's included in v0.1.0

| Tool | Current capabilities |
| --- | --- |
| Pixel Studio | Open an image, draw, erase, add lines and shapes, add text, export PNG |
| Vector Studio | Create SVG rectangles, ellipses, lines and text; export SVG |
| Video Studio | Preview a local video and attempt an in/out trim export as WebM where browser support permits |
| Audio Studio | Record microphone audio locally and download a WebM recording |
| PDF Desk | Preview a local PDF using the browser's built-in viewer |
| Project Vault | Save and reopen vector projects as portable JSON |

## Run locally

1. Download or clone this repository.
2. Open `index.html` in a modern browser.
3. For microphone access and some browser APIs, use HTTPS or localhost.

No account or subscription is required. Most processing in this MVP happens locally in your browser. Browser codec and device API support varies, especially on mobile.

## Project files

- `index.html` — complete browser app
- `productivity.html`, `productivity-sheet-engine.js` and `productivity-interoperability.js` — local-first productivity suite and Office format helpers
- `portal.html` and `portal-enhancements.js` — independent Yahoo-inspired portal, local drafts/contacts, calendar export and backup tools
- `LICENSE` — MIT License
- `.gitignore` — common local development exclusions
- `OpenCreative-v0.1.0.zip` — packaged v0.1.0 source bundle

## Roadmap

The current release is a starting point, not a complete replacement for professional creative applications. Planned areas include:

- Advanced photo editing, RAW development, layers, masks and color management
- Professional vector paths, typography and boolean operations
- Multitrack video editing, transitions, effects and export formats
- Multitrack audio editing, mixing and mastering
- Motion graphics, compositing, keyframes and tracking
- Multi-page publishing and print workflows
- PDF annotation and page editing
- 2D animation, rigging and sprite tools
- 3D modelling, materials and rendering
- Optional self-hosted sync and collaboration

## Principles

- Free to use, with no subscription gate in this project
- Local-first handling of user-selected files
- Original code, assets and interface design
- Open formats and modular development
- Respect for third-party licenses and intellectual property

## License

Released under the MIT License. See [LICENSE](LICENSE). Third-party libraries and assets added in the future must retain their own notices and compatible licenses.


## Why OpenCreative?

OpenCreative brings useful creative starter tools together in one browser-based workspace. It is designed to be free to use, local-first, built with original code, and developed openly so the project can grow with community contributions.

## Screenshots

Screenshots and product previews will be added here as the interface evolves. You can currently explore the tools by opening the app from this repository.

## How to Get Started

1. Open the repository's `index.html` file, or download/clone the repository.
2. Launch `index.html` in a modern web browser.
3. Choose a tool from the workspace and open a local file when needed.
4. For microphone access and browser features that require a secure context, run the app through HTTPS or localhost.

Some features depend on browser and device support. Keep your original files safe and test exports before relying on them for important work.

## Project Tagline

**OpenCreative — One creative universe. Countless possibilities. Free and open source.**

## Join the Community

OpenCreative grows through feedback and contributions. You do not need to be an expert developer to help.

- **Report a bug:** [Open a bug report](../../issues/new?template=bug_report.md)
- **Suggest an idea:** [Request a feature](../../issues/new?template=feature_request.md)
- **Contribute code or documentation:** Read [CONTRIBUTING.md](CONTRIBUTING.md).
- **Community standards:** Read our [Code of Conduct](CODE_OF_CONDUCT.md).
- **Get help:** See [SUPPORT.md](SUPPORT.md).
- **Discuss the roadmap:** Visit [GitHub Discussions](../../discussions) when enabled.

Good first contributions include improving documentation, testing a tool in another browser, reporting a reproducible bug, and suggesting a focused improvement.


## Productivity Workspace

Open `productivity.html` or choose **Productivity Workspace** from the home page to use the local-first productivity apps:

- **Writer:** rich-text editing with HTML and TXT export.
- **Sheets:** multiple worksheets, cell bold/italic and text/fill colors, alignment, row filtering, sorting by the selected column, formula bar, CSV import/export, and XLSX workbooks with multiple sheets.
- **Tasks:** add, complete, due-date and delete tasks.
- **Notes:** create, edit, search and delete notes.
- **Calendar:** create local events and export an ICS calendar.
- **Forms:** create a one-question form, collect local responses and export responses as CSV.
- **Backup:** export or import a JSON backup of the productivity workspace.

Productivity data is stored in the current browser's local storage. It is not synced to a server. Export backups regularly; clearing browser data can erase local content. This release does not include accounts, cloud sync, realtime collaboration, hosted email/meetings, or full DOCX/XLSX/PPTX compatibility. Use HTTPS or localhost for best browser support.

### Office document interoperability

- **Writer:** import `.docx` files in the browser and export Writer content as `.docx`, HTML, or TXT. DOCX export includes a Normal style, Heading 1–6 style definitions, bold/italic/underline runs, and basic ordered or bulleted list markers.
- **Sheets:** import and export multi-worksheet `.xlsx` workbooks. Import retains cell coordinates even when a worksheet's used range does not begin at A1, keeps formula expressions, and restores common bold/italic, font color, fill color and horizontal alignment when those styles are exposed by the parser.
- XLSX export reopens the generated workbook in SheetJS before download and checks worksheet count, non-empty values and formula expressions. Invalid or duplicate worksheet names are sanitized for the exported copy. **Cell formatting is not written to XLSX yet**; the local OpenCreative workbook retains its formatting.
- DOCX import uses Mammoth.js; DOCX packaging uses JSZip; XLSX processing uses SheetJS Community Edition. These browser libraries load on demand from their public CDNs.
- The selected document/workbook is processed in the browser. OpenCreative does not upload its contents to an OpenCreative server. Loading the libraries does make a request to the relevant CDN.
- **Compatibility limits:** this is not a full Microsoft Office renderer. DOCX import/export focuses on text, paragraphs, headings, basic emphasis and lists; tables, images, hyperlinks and complex layouts are not reliably preserved. XLSX does not support macros, external links, charts, pivot tables, all number/date formats or the full Excel function set. Formula calculations are evaluated by OpenCreative's local formula engine; supported functions include SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, AND, OR, NOT, ROUND, ABS, SQRT, POWER, CONCAT, LEN, UPPER and LOWER, plus cell references, ranges and arithmetic/comparison operators. Imported formula expressions are preserved, but unsupported formulas can display `#ERROR`.
- **Import limits:** DOCX files up to 15 MB; XLSX files up to 25 MB, 50 worksheets, 5,000 rows and 256 columns per sheet, and 250,000 cells across the workbook. These are app-level limits, not guarantees against every malformed or unusually compressed Office file.
- If a CDN is unavailable or the device is offline, existing local Writer/Sheets features continue to work; Office import/export will be unavailable until the library can load.

### Spreadsheet engine

The Sheets workspace supports multiple named worksheets, adding/renaming/deleting sheets, a formula bar, cell bold/italic, text and fill colors, alignment, sorting by the selected column, and live row filtering. The formula parser supports cell references, ranges, arithmetic and comparisons, plus SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, AND, OR, NOT, ROUND, ABS, SQRT, POWER, CONCAT, LEN, UPPER and LOWER. XLSX import/export includes all worksheets and formula expressions. Import restores a subset of common cell styles where available; export currently preserves values and formulas but not cell formatting. Exact Excel style fidelity, charts, pivot tables, macros, external workbook links and advanced Excel functions are not supported. XLSX import/export requires the SheetJS library to load from its public CDN.


### All-in-One Portal

Open `portal.html` for an independent Yahoo-inspired dashboard with news discovery links, a browser-based RSS/Atom reader, web search, live weather lookup via Open-Meteo, finance watchlist shortcuts, mail-provider links, local calendar events and notes, a local contacts book, and an email draft desk. The feed reader supports direct HTTPS RSS/Atom subscriptions, refresh, headline/summary search, a local reading list, and OPML import/export. It does not use a proxy: publishers must allow browser cross-origin (CORS) requests, so some feeds may not load from a static site. Feed URLs are fetched without credentials and are not sent to a third-party proxy. Drafts can be saved on this device, opened in the configured email app with a `mailto:` link, or downloaded as `.eml`; OpenCreative does not access mailboxes or send messages. Calendar events can be exported as `.ics`, and portal data can be exported or merged from a JSON backup. It does not copy Yahoo's proprietary code, provide brokerage quotes, or connect to personal accounts. Weather and feed refresh require an internet connection; personal portal data stays in this browser's localStorage.
