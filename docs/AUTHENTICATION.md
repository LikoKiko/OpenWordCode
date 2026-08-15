# Authentication

## Implemented methods

| Method | Providers | Behavior |
| --- | --- | --- |
| API key | OpenAI, Anthropic, OpenRouter, Gemini, xAI, Kimi, custom | Key is accepted by Core, encrypted at rest, and injected only into the provider request. |
| Environment | Same cloud providers | Only the variable name is stored. The Core resolves the value at request time. |
| None | Ollama, LM Studio, Demo | No credential is sent. Endpoint reachability is tested separately. |
| OpenWordCode Bridge | OpenWordCode Bridge | Uses the first-party loopback compatibility API; the Bridge routes through the configured OpenWordCode provider/account runtime. |
| OAuth / account sign-in | Claude, xAI, Kimi Code, Google Antigravity, GitHub Copilot, Nous Portal, and the OpenWordCode account bridge | Starts the provider's browser or device flow, exchanges the code on Core, encrypts the returned token, and refreshes it when possible. |

The UI lists only methods that the provider adapter actually supports. OAuth is provider-specific: the Word pane starts sign-in, but Core owns the loopback callback/device polling and the encrypted credential. The provider password is never entered into Word. The `chatgpt.com/codex/open-app` page is not embedded in the task pane and is not treated as an authorization API; the add-in uses a normal redirect-based OAuth callback instead.

The normal public setup does not require OAuth: OpenAI is selected by default and the user pastes their own API key into the task pane. The key is sent to the local Core process, stored there encrypted, and is never returned to the browser UI. OAuth is optional; select a provider and use its **Sign in** card from the task pane or Settings.

## Supported provider OAuth

The current first-party adapters are:

- **Claude** — Claude Code-compatible OAuth headers and account session identity. This is experimental outside Anthropic's own clients; users should verify their plan and provider terms.
- **xAI** — OpenID Connect authorization-code flow with PKCE.
- **Kimi Code** — device authorization flow against `auth.kimi.com`, using the Kimi coding endpoint rather than the separate Moonshot API-key surface.
- **Google Antigravity** — Google authorization-code flow with PKCE, followed by Cloud Code Assist project discovery. The OAuth access token is not treated as a Gemini API key.
- **GitHub Copilot** — GitHub device flow followed by the Copilot token exchange. This is an unofficial integration and requires an eligible Copilot account; GitHub may change or revoke it.
- **Nous Portal** — device authorization flow with rotating refresh tokens.

**Cursor**, **Kiro**, and **Command Code** are deliberately shown as custom-transport-only in the OAuth catalog but are not enabled as generic bearer-token providers here: each requires a proprietary request protocol, account metadata, or wire adapter. Attaching an OAuth token to a normal `/chat/completions` request would look like a login feature while failing at runtime. They should be added only with a dedicated, provider-authorized adapter.

The OAuth client identifiers used for the provider flows are public-client identifiers, not an OpenWordCode secret. Do not add client secrets, refresh tokens, browser cookies, or provider credential files to the repository. Google Antigravity's client secret is configurable through `GOOGLE_ANTIGRAVITY_CLIENT_SECRET`; use a provider-approved registration for a production distribution.

The built-in callback is `http://localhost:10200/oauth/chatgpt/callback` unless `OPENWORDCODE_OPENAI_OAUTH_REDIRECT_URI` is set. The callback must be registered with the OAuth client, and the Core must be restarted after changing environment variables. The default authorization and token URLs can be overridden with `OPENWORDCODE_OPENAI_OAUTH_AUTHORIZE_URL` and `OPENWORDCODE_OPENAI_OAUTH_TOKEN_URL`; the default scope can be overridden with `OPENWORDCODE_OPENAI_OAUTH_SCOPE`.

This account runtime is not the documented OpenAI Platform API. Platform API requests continue to use API keys. The account endpoint is an experimental provider integration and may change independently of the Platform API; do not ship it as a guaranteed public integration without confirming provider authorization.

## Credential lifecycle

1. The task pane posts an API key to the Core over the configured local origin.
2. The Core validates it is a single-line value and stores it under `credentials.enc.json` encrypted with AES-256-GCM.
3. `config.json` stores only `credentialRef: "provider:<id>"`.
4. Provider adapters resolve the secret immediately before a request and never return it to the task pane.
5. Disconnect deletes the stored entry and restores the provider's environment/no-auth mode when available.

For the OpenWordCode account provider, disconnect removes the encrypted OAuth credential and leaves the Bridge in its sign-in-required state. The task pane never receives the access or refresh token.

The fallback encrypted file is an isolated, honest fallback rather than a claim that it is equivalent to an OS keychain. A native credential backend is the next hardening step for packaged releases.
