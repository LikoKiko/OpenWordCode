# Testing

Run:

```powershell
npm run typecheck
npm test
npm run build
```

The test suite covers:

- read-tool loops and proposal generation;
- manual approval, automatic safe-edit application, and safety-gated actions;
- stale approval and required approval transitions;
- memory Word selection replacement and stale rejection;
- secret redaction and constant-time token comparison;
- Core health, bootstrap auth, origin blocking, streamed demo agent flow, bounded image/PDF attachments, and the first-party OpenWordCode Bridge compatibility routes.

No external bridge is required by the test suite. No cloud model request was made because that can spend provider quota. Real Word desktop automation still requires manual sideload verification.
