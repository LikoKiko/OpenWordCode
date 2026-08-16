# Providers

Provider adapters implement the same `listModels()` and `streamChat()` contract. The agent runtime does not branch on provider names.

## Defaults

| Provider | Transport | Default endpoint | Auth |
| --- | --- | --- | --- |
| OpenAI | OpenAI-compatible | `https://api.openai.com/v1` | Use OpenWordCode Bridge with Codex CLI |
| Anthropic | Messages API | `https://api.anthropic.com/v1` | Claude Code CLI or OAuth |
| OpenRouter | OpenAI-compatible | `https://openrouter.ai/api/v1` | No public account connector |
| Google Gemini | Gemini REST | `https://generativelanguage.googleapis.com/v1beta` | Use Google Antigravity CLI/OAuth |
| xAI | OpenAI-compatible | `https://api.x.ai/v1` | OAuth |
| Kimi Code | OpenAI-compatible | `https://api.kimi.com/coding/v1` | Kimi CLI or OAuth |
Provider endpoints are configurable through the Core API. Model discovery is live and bounded; disconnected providers do not prevent the rest of the UI from loading. Legacy API-key/environment transports remain available to existing Core configurations, but the public Word task pane does not expose them.

Account and CLI connector details are documented in [`AUTHENTICATION.md`](AUTHENTICATION.md). Custom OpenAI-compatible endpoints, including locally hosted ones, can still be added through the Core provider API.

Kimi Code OAuth and the Kimi Code endpoint are separate from the Moonshot API
surface. Provider-specific OAuth details and limitations are documented in
[`AUTHENTICATION.md`](AUTHENTICATION.md).
