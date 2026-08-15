# Architecture

OpenWordCode keeps the shared agent/provider/authentication core independent from Office.js so the same core can later support Excel, PowerPoint, or Outlook adapters.

## Boundaries

```text
React task pane
  → Core HTTP client
    → Agent runtime + context engine
      → Provider runtime
        → OpenAI-compatible / Anthropic / Gemini / local endpoint

React task pane
  → WordApplicationAdapter
    → Office.js
```

The model receives tool definitions and document data. It never receives an Office.js object or arbitrary executable capability. A write request creates a `ProposedChange`; the Core and adapter both validate it before the adapter applies it.

## Core

`apps/core/src/server.ts` is a loopback-only Fastify service. It provides:

- `/health` and `/healthz` for liveness;
- a bootstrap session token for the local task pane;
- provider/auth/model routes;
- an SSE `/api/agent` endpoint;
- a change approval/completion protocol;
- settings, diagnostics, and OpenWordCode Bridge status.

The Core starts the first-party OpenWordCode Bridge alongside the task-pane service. The Bridge owns the loopback OpenAI-compatible data-plane surface (`/v1/models` and `/v1/chat/completions`) and routes requests through the configured OpenWordCode provider/account runtime. OpenWordCode does not import or manage another application's proxy.

## Configuration and secrets

Configuration is stored under the platform config directory, or `OPENWORDCODE_DATA_DIR` when set. Provider config stores a credential reference or environment variable name, never the raw API key. The fallback store encrypts secrets with AES-256-GCM and keeps the key in a separate permission-hardened file. A native Windows Credential Manager/macOS Keychain implementation can replace that class behind the `CredentialStore` interface.

The server binds to `127.0.0.1` and requires a short-lived-process session token on protected routes, with an additional CSRF header on mutations.
