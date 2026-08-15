# Design QA notes

OpenWordCode's task pane is designed for a narrow Word sidebar. The composer,
file menu, model menu, approval menu, and safety notice share the same tokens
and component styles from `apps/word-addin/src/styles.css`.

## Verified behavior

- The composer fills the available pane width and is not user-resizable.
- Uploads are available from the plus menu; connector and plugin actions are
  not rendered.
- The model menu exposes provider-owned model names, effort controls, selected
  state, and usage labels when supplied by the provider.
- The approval menu exposes manual approval, automatic approval, and skip-all
  modes with an explicit unsafe-mode notice.
- The task pane keeps the disclaimer and send controls visible at narrow Word
  widths.
- The UI uses the OpenWordCode design tokens and a single stylesheet rather
  than per-component color systems.

## Local visual checks

Run the development server and open `https://localhost:3000` in a browser. For
Word-pane fidelity, inspect the UI at approximately 520 CSS pixels wide and
verify the same states listed above. Local screenshots are intentionally
ignored by Git (`qa-*.png`) so private or stale screenshots are not published
accidentally.

## Validation

```powershell
npm run typecheck
npm test
npm run build
```
