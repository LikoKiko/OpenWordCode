# Security

- Core binds to `127.0.0.1` by default.
- Protected API routes require a process session token obtained through an allowlisted local origin.
- Mutating routes also require the matching CSRF header.
- Unknown tool names and invalid tool arguments fail closed.
- Request bodies, document context, tool results, paragraphs, and proposed changes have size limits.
- Legacy provider credentials are not returned after storage; config stores references only. The public task pane does not accept new API keys.
- Secrets are redacted from user-facing errors and are not logged.
- OpenWordCode Bridge binds to loopback only. An optional `OPENWORDCODE_BRIDGE_TOKEN` protects its compatibility routes in addition to the loopback boundary.
- Office writes require explicit approval and two stale-content checks.
- Document text is treated as untrusted data to reduce prompt-injection impact.

The Core is local, but a local service is still a security boundary. Do not expose port 10200 to a network interface, and keep the allowed-origin list narrow when embedding the task pane elsewhere.
