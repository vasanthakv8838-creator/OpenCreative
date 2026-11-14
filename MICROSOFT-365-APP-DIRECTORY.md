# OpenCreative Office Suite — Clean-Room Compatibility Blueprint

This file replaces the Microsoft 365 link directory with an implementation plan for an **independent productivity suite** inside OpenCreative.

## Important scope and naming

OpenCreative will build its own apps and backend from original code and original interface design. It will not decompile Microsoft binaries, copy Microsoft source code, reproduce protected branding/assets, bypass Microsoft subscriptions, or impersonate Microsoft services.

The goal is functional interoperability: create, edit, import, and export common office documents using published standards and documented APIs where permitted. This is **not** Microsoft 365, and it must not be marketed as an official Microsoft product or a complete clone.

## Product suite

Use original OpenCreative names and visual identity for each module.

| OpenCreative module | Office-style capability | Initial feature scope |
|---|---|---|
| OpenCreative Writer | Word-style word processing | Rich text, headings, lists, tables, page layout, images, comments, print/PDF export |
| OpenCreative Sheets | Excel-style spreadsheets | Grid editor, formulas, cell formats, sorting/filtering, charts, CSV import/export |
| OpenCreative Slides | PowerPoint-style presentations | Slides, layouts, text/images/shapes, themes, presenter view, PDF export |
| OpenCreative Mail | Outlook-style email client | Connect user-authorized mail accounts, folders, compose, search, attachments |
| OpenCreative Calendar | Calendar and scheduling | Events, reminders, recurring events, invitations, timezone support |
| OpenCreative Notes | OneNote-style notebooks | Notebooks, sections, pages, rich notes, tags, search |
| OpenCreative Drive | OneDrive-style file workspace | Upload/download, folders, sharing permissions, version history |
| OpenCreative Chat | Teams-style collaboration | Workspaces, channels, messages, file links, meetings integration |
| OpenCreative Sites | SharePoint-style team spaces | Pages, document libraries, lists, access controls |
| OpenCreative Tasks | To Do / Planner-style planning | Personal tasks, team boards, due dates, assignments |
| OpenCreative Forms | Forms-style surveys | Form builder, question types, validation, response summaries |
| OpenCreative Whiteboard | Collaborative canvas | Shapes, sticky notes, connectors, freehand drawing |
| OpenCreative Diagrams | Visio-style diagrams | Flowcharts, connectors, templates, SVG/PDF export |
| OpenCreative Reports | Power BI-style reporting | Data import, charts, filters, dashboards, shareable reports |
| OpenCreative Automations | Power Automate-style workflows | Triggers, actions, conditions, run history, user-approved integrations |
| OpenCreative Database | Access-style relational data tool | Tables, relationships, queries, forms, CSV import/export |
| OpenCreative Publisher | Desktop publishing | Page templates, typography, image placement, PDF export |
| OpenCreative Video | Clipchamp-style editor | Timeline, trims, transitions, titles, audio, video export |
| OpenCreative Design | Designer-style graphics | Templates, shapes, text, image editing, export |
| OpenCreative Admin | Admin and governance console | Users, roles, audit trail, storage quotas, policy settings |

## Interoperability and file formats

Prioritize open, documented formats. Use libraries that support the relevant standards; do not claim full fidelity until compatibility tests prove it.

| Module | Target formats |
|---|---|
| Writer | DOCX import/export where supported; ODT, RTF, TXT, PDF |
| Sheets | XLSX import/export where supported; ODS, CSV, TSV |
| Slides | PPTX import/export where supported; ODP, PDF |
| Mail | MIME / EML; connect through documented IMAP/SMTP or provider APIs with explicit user authorization |
| Calendar | iCalendar (.ics); CalDAV where supported |
| Notes | Markdown, HTML, JSON export |
| Diagrams | SVG, PNG, PDF, JSON project format |
| Database | CSV, JSON, SQLite |

Open XML file extensions include DOCX, XLSX, and PPTX. Support must be implemented and tested per feature; macros, embedded objects, advanced formulas, animations, and proprietary extensions may not be supported. Treat macro-enabled files as untrusted and never execute embedded code.

## Technical architecture

Build a modular web application with a shared shell and independently deployable services only where needed.

- **Frontend:** responsive TypeScript application; accessible navigation; app launcher; command palette; keyboard shortcuts; mobile-friendly layouts.
- **Backend:** documented REST API; modular services for documents, files, identity, collaboration, notifications, and audit events.
- **Database:** PostgreSQL for users, metadata, permissions, comments, tasks, and collaboration state.
- **File storage:** S3-compatible object storage with signed uploads/downloads, versioning, and quotas.
- **Realtime:** WebSocket collaboration with document-level authorization and conflict handling.
- **Search:** indexed metadata and document text with permission-aware results.
- **Authentication:** secure sessions or OAuth/OIDC; no collection or storage of Microsoft passwords.
- **Integrations:** optional Microsoft Graph or other provider APIs only through official documented authorization flows and user consent. Keep provider tokens encrypted and revocable.
- **Deployment:** environment-based configuration, HTTPS, backups, migrations, logging, rate limits, and secret management.

## Security and privacy requirements

1. Enforce authorization on every server-side read and write.
2. Validate uploads, enforce size limits, scan files, and serve downloads safely.
3. Escape/sanitize rendered content and protect against XSS, CSRF, SSRF, and injection.
4. Encrypt secrets and provider tokens; never put credentials in source control.
5. Provide account export and deletion workflows.
6. Keep audit logs for sharing, admin actions, and access changes.
7. Do not execute VBA, macros, scripts, or embedded file code.
8. Clearly disclose third-party services, storage limits, and any costs.

## Phased build plan

### Phase 1 — Working office core
- Shared app shell, account/session handling, file workspace.
- Writer with rich-text editing and DOCX/ODT/PDF export/import where supported.
- Sheets with grid editing, basic formulas, CSV/XLSX import/export where supported.
- Slides with basic slide editing and PDF/PPTX export/import where supported.
- Save/open flows, autosave, responsive UI, tests.

### Phase 2 — Personal productivity
- Mail connection through authorized provider APIs.
- Calendar with ICS import/export and reminders.
- Notes, tasks, forms, diagrams, and whiteboard.

### Phase 3 — Team collaboration
- Shared folders, permissions, comments, presence, collaborative editing.
- Team chat, sites, shared task boards, and notifications.

### Phase 4 — Business and advanced tools
- Reports, automation workflows, relational database UI, admin console.
- SSO, audit exports, retention policies, organization controls.

## Definition of done

A module is not complete merely because its card or button exists. It must have a working user flow, persistent storage, validation and error states, authorization checks, automated tests, and documented limitations. File-format compatibility must be tested against real sample files and reported honestly.

## Technical references

- [Microsoft Open XML and Office standards support](https://learn.microsoft.com/openspecs/office_standards/)
- [Microsoft Office file format reference](https://learn.microsoft.com/office/compatibility/office-file-format-reference)
- [About the Open XML SDK](https://learn.microsoft.com/en-us/office/open-xml/about-the-open-xml-sdk)
- [Microsoft Graph Excel API overview](https://learn.microsoft.com/en-us/graph/excel-concept-overview)

These references describe formats and documented integration surfaces. They do not grant permission to copy Microsoft products or branding. Check the license and terms of every library and API before using it.

## Current implementation status

**Blueprint only.** This document does not mean the modules, backend, authentication, cloud storage, collaboration, or file conversion engines have already been implemented. Build and test them incrementally, preserving existing OpenCreative files and functionality.
