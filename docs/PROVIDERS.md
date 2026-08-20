# Providers

Provider adapters implement the same `listModels()` and `streamChat()` contract. The agent runtime does not branch on provider names.

## Defaults

| Provider | Transport | Default endpoint | Auth |
| --- | --- | --- | --- |
| OpenWordCode Bridge | OpenAI-compatible loopback | `http://127.0.0.1:10101/v1` | Codex CLI or OpenWordCode account session |
| Anthropic | Messages API | `https://api.anthropic.com/v1` | Claude Code CLI or OAuth |
| Google Antigravity | Cloud Code Assist | `https://daily-cloudcode-pa.googleapis.com` | Antigravity CLI or Google OAuth |
| xAI | OpenAI-compatible | `https://api.x.ai/v1` | OAuth |
| Kimi Code | OpenAI-compatible | `https://api.kimi.com/coding/v1` | Kimi CLI or OAuth |
| Nous Portal | OpenAI-compatible | `https://inference-api.nousresearch.com/v1` | Device OAuth |
| GitHub Copilot | Copilot API | `https://api.githubcopilot.com` | GitHub device OAuth |

Model discovery is live and bounded; disconnected providers do not prevent the rest of the UI from loading. Fresh installs expose only account or official-CLI connectors. The public Word task pane has no API-key entry flow.

Account and CLI connector details are documented in [`AUTHENTICATION.md`](AUTHENTICATION.md).

Kimi Code OAuth and the Kimi Code endpoint are separate from the Moonshot API
surface. Provider-specific OAuth details and limitations are documented in
[`AUTHENTICATION.md`](AUTHENTICATION.md).
