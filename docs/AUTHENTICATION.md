# Authentication

## Implemented methods

| Method | Providers | Behavior |
| --- | --- | --- |
| Legacy API-key transport | Existing Core configurations only | Kept for backward compatibility; the public Word task pane no longer exposes API-key entry. |
| Legacy environment transport | Existing Core configurations only | Only the variable name is stored. The Core resolves the value at request time; the public task pane does not expose this setup. |
| None | Legacy/custom local endpoints | No credential is sent. This is not exposed as an account option in Word. |
| OpenWordCode Bridge | OpenWordCode Bridge | Uses the first-party loopback compatibility API; the Bridge routes through the configured OpenWordCode provider/account runtime. |
| Existing Codex CLI session | OpenWordCode Bridge | After the user explicitly chooses **Sign in with Codex CLI** or **Connect existing session**, Core reads the local Codex CLI session metadata/token on demand and never copies it into the OpenWordCode credential store. |
| Existing provider CLI session | Claude, Kimi, Google Antigravity | After the user explicitly chooses the provider's CLI connector, Core reads that official CLI's local session on demand. Refreshes remain in memory; CLI credentials are never copied into OpenWordCode storage. |
| OAuth / account sign-in | Claude, xAI, Kimi Code, Google Antigravity, GitHub Copilot, Nous Portal, and the OpenWordCode account bridge | Starts the provider's browser or device flow, exchanges the code on Core, encrypts the returned token, and refreshes it when possible. |

The UI lists only methods that the provider adapter actually supports. OAuth is provider-specific: the Word pane starts sign-in, but Core owns the loopback callback/device polling and the encrypted credential. The provider password is never entered into Word. The `chatgpt.com/codex/open-app` page is not embedded in the task pane and is not treated as an authorization API; the add-in uses a normal redirect-based OAuth callback instead.

The public setup is account-first: select the OpenWordCode Bridge and choose **Sign in with Codex CLI**, or choose a provider OAuth/CLI flow. Fresh installs do not include API-only providers, and the Word task pane does not expose API-key entry. For ChatGPT subscription access, OpenWordCode starts the official local `codex login` flow and watches for completion; **Connect existing session** remains available after a user runs the command manually. On Windows, Core discovers the accessible per-user Codex executable automatically, and `CODEX_CLI_PATH` can override it. Core reads the official local Codex CLI session only after that explicit action; it does not scrape browser cookies or copy the token into GitHub or the OpenWordCode credential store. If Codex uses a custom `CODEX_HOME`, the Core process must inherit the same value.

For provider subscription access, select Claude, Kimi, or Google Antigravity and use the matching **Sign in with … CLI** or **Connect existing … session** action. Claude Code credentials are read from `CLAUDE_CONFIG_DIR/.credentials.json` (or `~/.claude/.credentials.json`), and Kimi Code credentials are read from `KIMI_SHARE_DIR/credentials/kimi-code.json` (or `~/.kimi/credentials/kimi-code.json`). Antigravity is different: its official CLI uses the operating-system keyring, so Windows builds read the `gemini:antigravity` Credential Manager entry instead of guessing at plaintext files. The Antigravity CLI executable is `agy`; if it is not on PATH, set `ANTIGRAVITY_CLI_PATH`. These connectors do not read browser cookies, passwords, or unrelated application credentials.

## Supported provider OAuth

The current first-party adapters are:

- **Claude** — Claude Code-compatible OAuth headers and account session identity. This is experimental outside Anthropic's own clients; users should verify their plan and provider terms.
- **xAI** — OpenID Connect authorization-code flow with PKCE.
- **Kimi Code** — device authorization flow against `auth.kimi.com`, using the Kimi coding endpoint rather than the separate Moonshot API-key surface.
- **Google Antigravity** — Google authorization-code flow with PKCE, followed by Cloud Code Assist project discovery. The OAuth access token is not treated as a Gemini API key.
- **GitHub Copilot** — GitHub device flow followed by the Copilot token exchange. This is an unofficial integration and requires an eligible Copilot account; GitHub may change or revoke it.
- **Nous Portal** — device authorization flow with rotating refresh tokens.

**Cursor**, **Kiro**, and **Command Code** are deliberately shown as custom-transport-only in the OAuth catalog but are not enabled as generic bearer-token providers here: each requires a proprietary request protocol, account metadata, or wire adapter. Attaching an OAuth token to a normal `/chat/completions` request would look like a login feature while failing at runtime. They should be added only with a dedicated, provider-authorized adapter.

The OAuth client identifiers used for the provider flows are public-client identifiers, not user credentials. Installed-app OAuth registrations can include a bundled client value that cannot be kept confidential in a desktop application. Do not add refresh tokens, access tokens, browser cookies, or provider credential files to the repository. Deployment owners can override the bundled Google Antigravity registration with `GOOGLE_ANTIGRAVITY_CLIENT_ID` and `GOOGLE_ANTIGRAVITY_CLIENT_SECRET`.

The built-in callback is `http://localhost:10200/oauth/chatgpt/callback` unless `OPENWORDCODE_OPENAI_OAUTH_REDIRECT_URI` is set. The callback must be registered with the OAuth client, and the Core must be restarted after changing environment variables. The default authorization and token URLs can be overridden with `OPENWORDCODE_OPENAI_OAUTH_AUTHORIZE_URL` and `OPENWORDCODE_OPENAI_OAUTH_TOKEN_URL`; the default scope can be overridden with `OPENWORDCODE_OPENAI_OAUTH_SCOPE`.

Codex, Claude Code, Kimi Code, and Antigravity CLI sessions are separate provider credentials. Connecting one never authenticates another. Antigravity remains a separate Google Cloud Code Assist provider; its CLI connector relies on the official CLI's keyring session and may be affected by Google's CLI availability, account eligibility, quota, or terms.

xAI may show an authorization code in the browser instead of redirecting back to
Word. Copy that code into the xAI sign-in panel in Settings and choose
**Complete sign-in**. Google Antigravity can also use the official Antigravity
CLI session from Settings when browser OAuth is unavailable for an account.

This account runtime is not the documented OpenAI Platform API. The account
endpoint is an experimental provider integration and may change independently
of the Platform API; do not ship it as a guaranteed public integration without
confirming provider authorization. The public add-in intentionally does not
offer direct API-key setup.

## Credential lifecycle

1. The task pane starts an account or local-CLI connection through the Core over
   the configured local origin.
2. OAuth credentials are stored under `credentials.enc.json` encrypted with
   AES-256-GCM; CLI credentials remain owned by the official CLI and are read
   only on demand.
3. `config.json` stores provider metadata and credential references, never raw
   account tokens.
4. Provider adapters resolve the active account credential immediately before a
   request and never return it to the task pane.
5. Disconnect removes the OpenWordCode OAuth credential or local-session marker
   and returns the provider to its sign-in-required state.

For the OpenWordCode account provider, disconnect removes the encrypted OAuth credential and leaves the Bridge in its sign-in-required state. The task pane never receives the access or refresh token.

The fallback encrypted file is an isolated, honest fallback rather than a claim that it is equivalent to an OS keychain. A native credential backend is the next hardening step for packaged releases.
