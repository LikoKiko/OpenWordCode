# OpenWordCode 0.1.0

Initial public release of OpenWordCode, an open-source AI document agent for
Microsoft Word created and maintained by Liko.

## Included

- React and Office.js task pane for Microsoft Word.
- Local TypeScript Core with provider routing, model discovery, streaming, and
  protected loopback sessions.
- OpenWordCode Bridge support plus provider OAuth and official local CLI session
  connectors where configured.
- Multiple provider selection with model lists for OpenWordCode Bridge,
  Anthropic, Google Antigravity, xAI, Kimi Code, Nous Portal, and GitHub
  Copilot.
- Document reading for selections, paragraphs, tables, pictures, outline data,
  and supported Word layout information.
- Review, rewrite, summarize, draft, format, table, and picture workflows with
  approval controls and change tracking.
- Selected-table resize and in-place editing safeguards to avoid accidental
  nested tables.
- Image and PDF attachments, drag-and-drop, clipboard image paste, previews,
  and multimodal provider input when supported.
- Web search and bounded read-only workspace inspection.
- AI skills loaded from local skill files, with an upload path for custom
  document workflows.
- Chat history, new tasks, context usage display, automatic conversation
  compaction, interactive clarification questions, answer copying, and change
  revert support.
- Centralized responsive UI styling and a transparent OpenWordCode logo.
- Automated typecheck, tests, and production build checks in GitHub Actions.

## Important limits

Provider authorization, model availability, Office.js support, and usage limits
depend on the user's provider account and installed Word build. The repository
does not contain provider passwords, user tokens, API keys, or a hosted backend.
Read `SECURITY.md`, `docs/AUTHENTICATION.md`, and `docs/PRIVACY.md` before
sharing sensitive documents with a cloud provider.
