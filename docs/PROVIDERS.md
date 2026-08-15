# Providers

Provider adapters implement the same `listModels()` and `streamChat()` contract. The agent runtime does not branch on provider names.

## Defaults

| Provider | Transport | Default endpoint | Auth |
| --- | --- | --- | --- |
| OpenAI | OpenAI-compatible | `https://api.openai.com/v1` | API key or `OPENAI_API_KEY` |
| Anthropic | Messages API | `https://api.anthropic.com/v1` | API key or `ANTHROPIC_API_KEY` |
| OpenRouter | OpenAI-compatible | `https://openrouter.ai/api/v1` | API key or `OPENROUTER_API_KEY` |
| Google Gemini | Gemini REST | `https://generativelanguage.googleapis.com/v1beta` | API key or `GEMINI_API_KEY` |
| xAI | OpenAI-compatible | `https://api.x.ai/v1` | API key or `XAI_API_KEY` |
| Kimi Code | OpenAI-compatible | `https://api.kimi.com/coding/v1` | OAuth or `MOONSHOT_API_KEY` |
| Ollama | OpenAI-compatible | `http://127.0.0.1:11434/v1` | None |
| LM Studio | OpenAI-compatible | `http://127.0.0.1:1234/v1` | None |

Provider endpoints are configurable through the Core API. Model discovery is live and bounded; disconnected providers do not prevent the rest of the UI from loading.

OpenAI's [API authentication documentation](https://platform.openai.com/docs/api-reference/authentication) recommends server-side environment variables/key management and Bearer authentication. Anthropic direct requests use `x-api-key` according to its [authentication documentation](https://platform.claude.com/docs/en/build-with-claude/authentication). Gemini supports API-key authentication through its [API key guide](https://ai.google.dev/gemini-api/docs/api-key), while [Ollama's OpenAI compatibility API](https://docs.ollama.com/api/openai-compatibility) and [LM Studio's compatibility endpoint](https://lmstudio.ai/lmstudio/openai-compat-endpoint) cover local transports.

Kimi Code OAuth and the Kimi Code endpoint are separate from the Moonshot API
surface. Provider-specific OAuth details and limitations are documented in
[`AUTHENTICATION.md`](AUTHENTICATION.md).
