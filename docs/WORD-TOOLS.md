# Word tools

The browser/task-pane side sends a simplified `DocumentSnapshot`:

- current selection and a stable target fingerprint;
- bounded document text;
- bounded paragraphs and outline;
- host capabilities.

The Core's read tools operate on this snapshot, so a model cannot request arbitrary Word state. Office.js stays inside `packages/app-word`.

## Office adapter

`OfficeWordAdapter` uses `Word.run()` to load selection, body text, and paragraphs. It supports verified selection and paragraph text replacement, plus allowlisted `Word.Range` mutation proposals. Range operations are dispatched by documented member name only; Office.js objects never cross the Core boundary.

The shared Range catalog includes the documented properties, methods, and comment events from the Microsoft Learn `Word.Range` surface. Mutating proposals can use operations such as `insertText`, `insertHtml`, `insertOoxml`, `insertParagraph`, `insertBreak`, `insertTable`, `insertComment`, `insertBookmark`, `insertContentControl`, picture/file/field/note insertion, `clear`, `delete`, `set`, formatting properties, and selection/highlight operations. Read-only members remain available in the catalog for capability awareness but are rejected as write proposals.

The runtime checks the host's API set and reports a clear failure when a preview-only or unsupported member is unavailable. This matters because the linked page includes members from multiple Word API sets and not every Word desktop build exposes every preview member.

`MemoryWordAdapter` is used outside Word for browser preview and automated tests. It has the same contract and stale-target behavior, with text/table/HTML/paragraph operations rendered into preview text and unsupported visual-only operations reported instead of faked.

Whole-table selections are handled separately because Word can report an empty
`Range.text` for a selected table. The Office adapter loads the selected range's
table collection and parent-table relationship, exposes that metadata to the
agent, and routes `scope:"table"` operations to `Word.Table`, including deletion
and portable built-in styles such as `GridTable4_Accent1`.

## DOCX skill

The attached DOCX skill is stored at `skills/docx/SKILL.md`. The Core loads it automatically when the agent runtime starts and appends it to the document-domain system context. Its standalone-file guidance is retained, while live Word operations are constrained to the Range catalog and the approval protocol above.

## Apply protocol

1. Agent creates a validated change containing `beforeText` and `beforeFingerprint`.
2. Task pane reads the live target again.
3. Core `/api/changes/:id/approve` compares both text and fingerprint.
4. Adapter re-reads the Word target, temporarily disables Word tracking while applying the clean edit, and restores the user's previous tracking setting. Before/After is shown in the OpenWordCode panel instead of being injected as redline markup into the document.
5. Task pane reports success to `/complete` and records the edit in the chat history.

Any mismatch fails closed.

## Optional agent tools

The composer can enable live web search for current information. Search requests
run through the configured loopback OpenWordCode Bridge and return a short answer
with source URLs; they do not receive the Word document unless the user includes
that context in the request.

The composer can also enable a bounded Windows console. It is intentionally
read-only and workspace-scoped: commands such as `Get-Location`, `Get-Content`,
`rg`, and inspection-only `git`/`npm` checks are allowed, while deletion,
downloads, elevation, shell chaining, and arbitrary program execution are
blocked. In manual approval mode, every console action appears as a pending
card before it runs. Auto mode only runs commands that pass the same validator.
The console is for inspecting the workspace; Word edits still go through the
Office adapter so they can be validated against the live document.
