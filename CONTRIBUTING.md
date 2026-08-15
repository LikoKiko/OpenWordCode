# Contributing to OpenWordCode

Thanks for helping improve OpenWordCode. The project is currently an early
public development build, so small, focused changes are easiest to review.

## Before opening a pull request

1. Create a branch from the current default branch.
2. Do not include API keys, OAuth tokens, private documents, `.env` files,
   Office certificates, or files from a local credential store.
3. Run the checks below:

```powershell
npm ci
npm run typecheck
npm test
npm run build
```

4. If the change touches Office.js, also test the add-in by sideloading the
   manifest into Word Desktop and describe the host/version you tested.

## Development expectations

- Keep Office.js access inside `packages/app-word`.
- Keep provider/authentication code independent from the Word adapter.
- Validate all data crossing the Core HTTP boundary.
- Preserve the approval flow and stale-target checks for document writes.
- Add or update tests when behavior changes.
- Keep secrets and provider-specific credentials out of source and logs.

## Pull requests

Describe what changed, why it changed, how it was tested, and any known
limitations. UI changes should include a screenshot when practical. Changes
to provider authentication must explain the provider's documented or
provider-approved flow and must not rely on scraped cookies or private files.

By contributing, you agree that your contribution is provided under the
project's MIT License.
