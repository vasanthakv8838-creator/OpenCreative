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
- `package.json` — minimal project metadata and optional local server command
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
