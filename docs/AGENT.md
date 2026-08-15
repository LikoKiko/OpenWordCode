# Agent runtime

The runtime runs a bounded loop with a default maximum of eight model iterations and sixteen tool calls. Read tools are validated with Zod and return capped JSON results. `propose_change` is the only write-like tool; it creates a validated change object, and the task pane applies it through the stale-checking Word adapter.

The system prompt explicitly treats document content as untrusted data. The context builder includes selection, nearby paragraphs, and bounded document text. It does not silently claim that a truncated document was fully analyzed.

Providers with native OpenAI-compatible tool calling can use the read tools directly. Providers without usable native tool calling are instructed to return a constrained JSON result, which the Core validates before creating proposals.
