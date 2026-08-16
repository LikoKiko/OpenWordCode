# OpenWordCode Bridge

OpenWordCode includes its own loopback compatibility gateway. It is started by `npm run start` / `npm run dev:core` together with Core, so the Word add-in operates independently without any external local proxy.

## Local surface

The default listener is `http://127.0.0.1:10101`:

- `GET /healthz` — liveness;
- `GET /readyz` — catalog readiness;
- `GET /v1/models` — aggregated models from the configured OpenWordCode runtimes;
- `GET /v1/models/:id` — one model;
- `POST /v1/chat/completions` — OpenAI-compatible streaming and non-streaming chat;
- `POST /v1/responses` — bounded Responses passthrough used by the first-party web-search path.

The bridge accepts text, image data URLs, PDF data URLs, function tools, streamed text, and streamed tool calls. It translates the compatibility request into OpenWordCode's shared provider contract, so the same approval and Word-edit behavior is used whether the task pane calls the Core directly or a local compatibility client calls the Bridge.

## Routing

The bridge prefers the internal OpenWordCode account transport. It then adds other enabled providers to the model catalog when they are configured and discoverable. To pin routing, set:

```powershell
$env:OPENWORDCODE_BRIDGE_PROVIDER_ID = "openwordcode-account"
```

The bridge provider itself is excluded from its upstream list to prevent request loops. The account credential is resolved by the Core's encrypted credential store; the bridge never reads browser cookies, another application's credential files, or passwords.

Set `OPENWORDCODE_BRIDGE_PORT` to change the listener. The task-pane provider is migrated/configured to the same local URL automatically. Set `OPENWORDCODE_BRIDGE_TOKEN` when another local process must authenticate to the compatibility routes; OpenWordCode sends that token on its own bridge-provider requests.

## Open-source boundary

This gateway is OpenWordCode code. It operates completely standalone and does not rely on third-party proxies, browser cookies, or external credential stores. The account OAuth client ID belongs to the OpenWordCode deployment and can be customized through environment variables if needed.

The account transport is experimental and is separate from the documented OpenAI Platform API. The public Word task pane is account-first; deployments with older Core configurations may still have a legacy direct-provider credential behind the bridge, but new API-key setup is intentionally not exposed.
