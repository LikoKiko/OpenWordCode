import { describe, expect, it } from "vitest";
import { MemoryCredentialStore } from "../packages/auth/src/index.js";
import {
  accountProviderFallbackModels,
  antigravityUserAgent,
  createProvider,
  parseAntigravityAvailableModels,
  parseChatGptAccountModels,
  resolveAntigravityEffortWireModel,
} from "../packages/providers/src/index.js";

describe("OpenAI-compatible provider", () => {
  it("discovers models and sends the configured bearer credential only to the provider", async () => {
    const store = new MemoryCredentialStore();
    await store.set("provider:test", "provider-secret");
    const seen: Request[] = [];
    const fetchImpl: typeof fetch = async (input, init) => {
      const request = new Request(input, init);
      seen.push(request);
      if (request.url.endsWith("/models")) return new Response(JSON.stringify({ data: [{ id: "model-a", owned_by: "test" }] }), { headers: { "content-type": "application/json" } });
      return new Response("data: {\"choices\":[{\"delta\":{\"content\":\"hello\"}}]}\n\ndata: [DONE]\n\n", { headers: { "content-type": "text/event-stream" } });
    };
    const provider = createProvider({ id: "test", displayName: "Test", kind: "openai-compatible", baseUrl: "http://127.0.0.1:9999/v1", enabled: true, local: true, auth: { method: "api-key", credentialRef: "provider:test" }, privacyNote: "test" }, { store, fetchImpl });
    expect((await provider.listModels())[0]?.id).toBe("model-a");
    const events = [];
    for await (const event of provider.streamChat({ model: "model-a", effort: "high", messages: [{ role: "user", content: "Review these files" }], attachments: [{ id: "image-1", name: "chart.png", mimeType: "image/png", size: 1, dataUrl: "data:image/png;base64,AA==" }, { id: "pdf-1", name: "brief.pdf", mimeType: "application/pdf", size: 1, dataUrl: "data:application/pdf;base64,AA==" }] })) events.push(event);
    expect(events.some(event => event.type === "text" && event.delta === "hello")).toBe(true);
    expect(seen.every(request => request.headers.get("authorization") === "Bearer provider-secret")).toBe(true);
    const chatBody = JSON.parse(await seen[1]!.text()) as { messages: Array<{ content: unknown }>; reasoning_effort?: string };
    expect(chatBody.reasoning_effort).toBe("high");
    expect(chatBody.messages.at(-1)?.content).toEqual([
      { type: "text", text: "Review these files" },
      { type: "image_url", image_url: { url: "data:image/png;base64,AA==" } },
      { type: "file", file: { filename: "brief.pdf", file_data: "data:application/pdf;base64,AA==" } },
    ]);
  });

  it("does not duplicate a complete function name repeated by a streaming bridge", async () => {
    const store = new MemoryCredentialStore();
    const fetchImpl: typeof fetch = async () => new Response([
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call-1","function":{"name":"web_search","arguments":"{\\"query\\":\\""}}]}}]}',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"name":"web_search","arguments":"Word"}}]}}]}',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"name":"web_search","arguments":"\\"}"}}]}}]}',
      "data: [DONE]",
      "",
    ].join("\n\n"), { headers: { "content-type": "text/event-stream" } });
    const provider = createProvider({ id: "test", displayName: "Test", kind: "openai-compatible", baseUrl: "http://127.0.0.1:9999/v1", enabled: true, local: true, auth: { method: "none" }, privacyNote: "test" }, { store, fetchImpl });
    const events = [];
    for await (const event of provider.streamChat({ model: "model-a", messages: [{ role: "user", content: "Search Word" }] })) events.push(event);
    const toolCall = events.find(event => event.type === "tool_call");
    expect(toolCall?.type === "tool_call" ? toolCall.call.name : undefined).toBe("web_search");
  });

  it("preserves a Google thought signature across the local compatibility bridge", async () => {
    const signature = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdef==";
    const requests: Request[] = [];
    const fetchImpl: typeof fetch = async (input, init) => {
      const request = new Request(input, init);
      requests.push(request);
      if (requests.length === 1) {
        return new Response([
          `data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"google-call","thought_signature":"${signature}","function":{"name":"get_selection","arguments":"{}"}}]}}]}`,
          "data: [DONE]",
          "",
        ].join("\n\n"), { headers: { "content-type": "text/event-stream" } });
      }
      return new Response('data: {"choices":[{"delta":{"content":"done"}}]}\n\ndata: [DONE]\n\n', { headers: { "content-type": "text/event-stream" } });
    };
    const provider = createProvider({ id: "bridge", displayName: "Bridge", kind: "openai-compatible", baseUrl: "http://127.0.0.1:10101/v1", enabled: true, local: true, auth: { method: "none" }, privacyNote: "test" }, { store: new MemoryCredentialStore(), fetchImpl });
    const firstEvents = [];
    for await (const event of provider.streamChat({ model: "gemini", messages: [{ role: "user", content: "Read the selection" }] })) firstEvents.push(event);
    const toolEvent = firstEvents.find(event => event.type === "tool_call");
    expect(toolEvent?.type === "tool_call" ? toolEvent.call.thoughtSignature : undefined).toBe(signature);
    if (!toolEvent || toolEvent.type !== "tool_call") throw new Error("Expected a tool call");
    for await (const _event of provider.streamChat({
      model: "gemini",
      messages: [
        { role: "user", content: "Read the selection" },
        { role: "assistant", content: "", toolCalls: [toolEvent.call] },
        { role: "tool", toolCallId: toolEvent.call.id, content: '{"text":"hello"}' },
      ],
    })) { /* drain */ }
    const secondBody = JSON.parse(await requests[1]!.text()) as { messages: Array<{ tool_calls?: Array<{ thought_signature?: string }> }> };
    expect(secondBody.messages[1]?.tool_calls?.[0]?.thought_signature).toBe(signature);
  });

  it("uses the Responses wire and editor fingerprint for newer GitHub Copilot models", async () => {
    const requests: Request[] = [];
    const fetchImpl: typeof fetch = async (input, init) => {
      const request = new Request(input, init);
      requests.push(request);
      if (request.url.endsWith("/responses")) {
        return new Response('data: {"type":"response.output_text.delta","delta":"new wire"}\n\ndata: {"type":"response.completed","response":{"output":[]}}\n\ndata: [DONE]\n\n', { headers: { "content-type": "text/event-stream" } });
      }
      return new Response('data: {"choices":[{"delta":{"content":"old wire"}}]}\n\ndata: [DONE]\n\n', { headers: { "content-type": "text/event-stream" } });
    };
    const provider = createProvider({ id: "github-copilot", displayName: "GitHub Copilot", kind: "openai-compatible", baseUrl: "https://api.githubcopilot.com", enabled: true, local: false, auth: { method: "oauth", oauthProvider: "github-copilot", oauthCredentialRef: "oauth:github-copilot" }, privacyNote: "test" }, {
      store: new MemoryCredentialStore(),
      fetchImpl,
      oauth: async () => ({ accessToken: "copilot-token", apiBaseUrl: "https://api.individual.githubcopilot.com" }),
    });
    const responseEvents = [];
    for await (const event of provider.streamChat({ model: "gpt-5.6-luna", effort: "high", messages: [{ role: "user", content: "Review" }] })) responseEvents.push(event);
    const chatEvents = [];
    for await (const event of provider.streamChat({ model: "gpt-4o", messages: [{ role: "user", content: "Review" }] })) chatEvents.push(event);
    expect(requests[0]?.url).toBe("https://api.individual.githubcopilot.com/responses");
    expect(requests[1]?.url).toBe("https://api.individual.githubcopilot.com/chat/completions");
    expect(requests.every(request => request.headers.get("authorization") === "Bearer copilot-token")).toBe(true);
    expect(requests.every(request => request.headers.get("copilot-integration-id") === "vscode-chat")).toBe(true);
    expect(requests.every(request => request.headers.get("editor-version") === "openwordcode/0.1.0")).toBe(true);
    const body = JSON.parse(await requests[0]!.text()) as { input?: unknown[]; messages?: unknown[]; reasoning?: { effort?: string } };
    expect(body.input).toBeDefined();
    expect(body.messages).toBeUndefined();
    expect(body.reasoning).toEqual({ effort: "high" });
    expect(responseEvents.some(event => event.type === "text" && event.delta === "new wire")).toBe(true);
    expect(chatEvents.some(event => event.type === "text" && event.delta === "old wire")).toBe(true);
  });

  it("uses the Kimi account catalog and normalizes the local 1M selector on the wire", async () => {
    const requests: Request[] = [];
    const provider = createProvider({ id: "kimi", displayName: "Kimi Code", kind: "openai-compatible", baseUrl: "https://api.kimi.com/coding/v1", enabled: true, local: false, auth: { method: "oauth", oauthProvider: "kimi", oauthCredentialRef: "oauth:kimi" }, privacyNote: "test" }, {
      store: new MemoryCredentialStore(),
      oauth: async () => ({ accessToken: "kimi-token" }),
      fetchImpl: async (input, init) => {
        requests.push(new Request(input, init));
        return new Response('data: {"choices":[{"delta":{"content":"ready"}}]}\n\ndata: [DONE]\n\n', { headers: { "content-type": "text/event-stream" } });
      },
    });
    const models = await provider.listModels();
    expect(models.map(model => model.id)).toEqual(expect.arrayContaining(["k3", "k3[1m]", "kimi-k2.7-code"]));
    expect(requests).toHaveLength(0);
    for await (const _event of provider.streamChat({ model: "k3[1m]", effort: "medium", messages: [{ role: "user", content: "Review" }] })) { /* drain */ }
    expect(requests[0]?.url).toBe("https://api.kimi.com/coding/v1/chat/completions");
    const body = JSON.parse(await requests[0]!.text()) as { model?: string; reasoning_effort?: string; prompt_cache_key?: string };
    expect(body.model).toBe("k3");
    expect(body.reasoning_effort).toBe("high");
    expect(body.prompt_cache_key).toMatch(/^owc-[a-f0-9]{32}$/);
  });

  it("falls back to the current xAI catalog and uses Responses for OAuth Grok 4.6", async () => {
    const requests: Request[] = [];
    const provider = createProvider({ id: "xai", displayName: "xAI", kind: "openai-compatible", baseUrl: "https://api.x.ai/v1", enabled: true, local: false, auth: { method: "oauth", oauthProvider: "xai", oauthCredentialRef: "oauth:xai" }, privacyNote: "test" }, {
      store: new MemoryCredentialStore(),
      oauth: async () => ({ accessToken: "xai-token" }),
      fetchImpl: async (input, init) => {
        const request = new Request(input, init);
        requests.push(request);
        if (request.url.endsWith("/models")) return new Response("unavailable", { status: 503 });
        return new Response('data: {"type":"response.output_text.delta","delta":"ready"}\n\ndata: {"type":"response.completed","response":{"output":[]}}\n\ndata: [DONE]\n\n', { headers: { "content-type": "text/event-stream" } });
      },
    });
    expect((await provider.listModels()).map(model => model.id)).toContain("grok-4.6");
    for await (const _event of provider.streamChat({ model: "grok-4.6", effort: "xhigh", messages: [{ role: "user", content: "Review" }] })) { /* drain */ }
    expect(requests[1]?.url).toBe("https://api.x.ai/v1/responses");
    const body = JSON.parse(await requests[1]!.text()) as { reasoning?: { effort?: string } };
    expect(body.reasoning).toEqual({ effort: "xhigh" });
  });

  it("keeps live discovery failures strict for providers without an account fallback", async () => {
    const provider = createProvider({ id: "custom", displayName: "Custom", kind: "openai-compatible", baseUrl: "https://provider.test/v1", enabled: true, local: false, auth: { method: "none" }, privacyNote: "test" }, {
      store: new MemoryCredentialStore(),
      fetchImpl: async () => new Response("unavailable", { status: 503 }),
    });
    await expect(provider.listModels()).rejects.toMatchObject({ code: "provider_http_error", status: 503 });
  });

  it("keeps complete fallback catalogs for every supported account provider", () => {
    const config = (oauthProvider: "xai" | "kimi" | "nous" | "github-copilot") => ({
      id: oauthProvider,
      displayName: oauthProvider,
      kind: "openai-compatible" as const,
      baseUrl: "https://provider.test/v1",
      enabled: true,
      local: false,
      auth: { method: "oauth" as const, oauthProvider, oauthCredentialRef: `oauth:${oauthProvider}` },
      privacyNote: "test",
    });
    expect(accountProviderFallbackModels(config("xai")).map(model => model.id)).toContain("grok-4.6");
    expect(accountProviderFallbackModels(config("kimi")).map(model => model.id)).toContain("k3[1m]");
    expect(accountProviderFallbackModels(config("nous")).map(model => model.id)).toEqual(expect.arrayContaining(["tencent/hy3:free", "stepfun/step-3.7-flash:free"]));
    expect(accountProviderFallbackModels(config("github-copilot")).map(model => model.id)).toEqual(expect.arrayContaining(["gpt-4o", "gpt-5.6-luna", "gemini-2.5-pro"]));
  });
});

describe("native tool providers", () => {
  it("preserves Anthropic tool-use and tool-result turns", async () => {
    const store = new MemoryCredentialStore();
    await store.set("provider:anthropic-test", "anthropic-secret");
    let requestBody: { messages?: Array<{ role: string; content: unknown }> } | undefined;
    const fetchImpl: typeof fetch = async (_input, init) => {
      requestBody = JSON.parse(String(init?.body)) as typeof requestBody;
      return new Response([
        'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"call-1","name":"run_console"}}',
        'data: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"command\\":\\"pwd\\"}"}}',
        'data: {"type":"message_stop"}',
        "",
      ].join("\n\n"), { headers: { "content-type": "text/event-stream" } });
    };
    const provider = createProvider({ id: "anthropic-test", displayName: "Anthropic Test", kind: "anthropic", baseUrl: "http://127.0.0.1:9999/v1", enabled: true, local: true, auth: { method: "api-key", credentialRef: "provider:anthropic-test" }, privacyNote: "test" }, { store, fetchImpl });
    const events = [];
    for await (const event of provider.streamChat({
      model: "claude-test",
      messages: [
        { role: "user", content: "Inspect the workspace" },
        { role: "assistant", content: "", toolCalls: [{ id: "call-1", name: "run_console", arguments: '{"command":"pwd"}' }] },
        { role: "tool", toolCallId: "call-1", content: '{"output":"C:\\\\workspace"}' },
      ],
      tools: [{ type: "function", function: { name: "run_console", description: "Inspect safely", parameters: { type: "object" } } }],
    })) events.push(event);
    expect(requestBody?.messages?.[1]?.content).toEqual([{ type: "tool_use", id: "call-1", name: "run_console", input: { command: "pwd" } }]);
    expect(requestBody?.messages?.[2]?.content).toEqual([{ type: "tool_result", tool_use_id: "call-1", content: '{"output":"C:\\\\workspace"}' }]);
    expect(events.some(event => event.type === "tool_call" && event.call.name === "run_console")).toBe(true);
  });

  it("uses the static Claude account catalog without calling an unsupported model endpoint", async () => {
    let fetchCount = 0;
    const provider = createProvider({ id: "anthropic", displayName: "Anthropic", kind: "anthropic", baseUrl: "https://api.anthropic.com/v1", enabled: true, local: false, auth: { method: "oauth", oauthProvider: "anthropic", oauthCredentialRef: "oauth:anthropic" }, privacyNote: "test" }, {
      store: new MemoryCredentialStore(),
      oauth: async () => ({ accessToken: "claude-token" }),
      fetchImpl: async () => {
        fetchCount += 1;
        return new Response("unexpected", { status: 500 });
      },
    });
    expect((await provider.listModels()).map(model => model.id)).toEqual(expect.arrayContaining(["claude-sonnet-5", "claude-opus-5", "claude-haiku-4-5"]));
    expect(fetchCount).toBe(0);
  });

  it("sends Gemini function declarations and returns streamed function calls", async () => {
    const store = new MemoryCredentialStore();
    await store.set("provider:gemini-test", "gemini-secret");
    const signature = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdef==";
    let requestBody: { contents?: unknown[]; tools?: unknown[] } | undefined;
    const fetchImpl: typeof fetch = async (_input, init) => {
      requestBody = JSON.parse(String(init?.body)) as typeof requestBody;
      return new Response(`data: {"candidates":[{"content":{"parts":[{"functionCall":{"id":"provider-call","name":"run_console","args":{"command":"pwd"}},"thoughtSignature":"${signature}"}]}}]}\n\ndata: [DONE]\n\n`, { headers: { "content-type": "text/event-stream" } });
    };
    const provider = createProvider({ id: "gemini-test", displayName: "Gemini Test", kind: "gemini", baseUrl: "http://127.0.0.1:9999/v1beta", enabled: true, local: true, auth: { method: "api-key", credentialRef: "provider:gemini-test" }, privacyNote: "test" }, { store, fetchImpl });
    const events = [];
    for await (const event of provider.streamChat({
      model: "gemini-test",
      messages: [
        { role: "user", content: "Inspect the workspace" },
        { role: "assistant", content: "", toolCalls: [{ id: "call-1", name: "run_console", arguments: '{"command":"pwd"}', thoughtSignature: signature }] },
        { role: "tool", toolCallId: "call-1", content: '{"output":"C:\\\\workspace"}' },
      ],
      tools: [{ type: "function", function: { name: "run_console", description: "Inspect safely", parameters: { type: "object" } } }],
    })) events.push(event);
    expect(requestBody?.tools).toEqual([{ functionDeclarations: [{ name: "run_console", description: "Inspect safely", parameters: { type: "object" } }] }]);
    expect((requestBody?.contents?.[1] as { parts?: unknown[] })?.parts?.[0]).toEqual({ functionCall: { id: "call-1", name: "run_console", args: { command: "pwd" } }, thoughtSignature: signature });
    expect((requestBody?.contents?.[2] as { parts?: unknown[] })?.parts?.[0]).toEqual({ functionResponse: { id: "call-1", name: "run_console", response: { output: "C:\\workspace" } } });
    expect(events.some(event => event.type === "tool_call" && event.call.id === "provider-call" && event.call.name === "run_console" && event.call.thoughtSignature === signature)).toBe(true);
  });
});

describe("Google model compatibility", () => {
  it("maps current and retired Antigravity model names to the accepted wire contract", () => {
    expect(resolveAntigravityEffortWireModel("gemini-3.7-flash")).toEqual({ wireModelId: "gemini-3.7-flash-tiered", thinkingLevel: "medium" });
    expect(resolveAntigravityEffortWireModel("gemini-3.7-flash", "high")).toEqual({ wireModelId: "gemini-3.7-flash-tiered", thinkingLevel: "high" });
    expect(resolveAntigravityEffortWireModel("gemini-3.1-pro", "low")).toEqual({ wireModelId: "gemini-3.1-pro-low", thinkingLevel: "low" });
    expect(resolveAntigravityEffortWireModel("gemini-3.1-pro", "high")).toEqual({ wireModelId: "gemini-pro-agent", thinkingLevel: "high" });
    expect(resolveAntigravityEffortWireModel("gemini-3.1-pro")).toEqual({ wireModelId: "gemini-pro-agent" });
    expect(resolveAntigravityEffortWireModel("gemini-3.6-flash-low")).toEqual({ wireModelId: "gemini-3.7-flash-tiered", thinkingLevel: "low" });
    expect(resolveAntigravityEffortWireModel("gemini-3.5-flash-high", "ultra")).toEqual({ wireModelId: "gemini-3.7-flash-tiered", thinkingLevel: "high" });
  });

  it("retries transient Gemini failures and sends the current tiered model with thinking level", async () => {
    const store = new MemoryCredentialStore();
    await store.set("provider:gemini-current", "gemini-secret");
    const requests: Request[] = [];
    const fetchImpl: typeof fetch = async (input, init) => {
      const request = new Request(input, init);
      requests.push(request);
      if (requests.length === 1) return new Response("temporary", { status: 503, headers: { "retry-after": "0" } });
      return new Response('data: {"candidates":[{"content":{"parts":[{"text":"ready"}]}}]}\n\ndata: [DONE]\n\n', { headers: { "content-type": "text/event-stream" } });
    };
    const provider = createProvider({ id: "gemini-current", displayName: "Gemini", kind: "gemini", baseUrl: "https://generativelanguage.googleapis.com/v1beta", enabled: true, local: false, auth: { method: "api-key", credentialRef: "provider:gemini-current" }, privacyNote: "test" }, { store, fetchImpl });
    const events = [];
    for await (const event of provider.streamChat({
      model: "gemini-3.7-flash",
      effort: "high",
      messages: [{ role: "user", content: "Review" }],
      tools: [{ type: "function", function: { name: "get_selection", description: "Read", parameters: { type: "object" } } }],
    })) events.push(event);
    expect(requests).toHaveLength(2);
    expect(requests[1]?.url).toContain("/models/gemini-3.7-flash-tiered:streamGenerateContent");
    const body = JSON.parse(await requests[1]!.text()) as { generationConfig?: { thinkingConfig?: { thinkingLevel?: string } }; toolConfig?: { functionCallingConfig?: { mode?: string } } };
    expect(body.generationConfig?.thinkingConfig?.thinkingLevel).toBe("high");
    expect(body.toolConfig?.functionCallingConfig?.mode).toBe("AUTO");
    expect(events.some(event => event.type === "text" && event.delta === "ready")).toBe(true);
  });

  it("discovers Antigravity models and uses the exact IDE fingerprint and per-family tool mode", async () => {
    const discoveryPayload = {
      models: {
        "gemini-3.7-flash-tiered": { displayName: "Gemini 3.7 Flash", maxTokens: 1_048_576, supportsThinking: true, supportsImages: true },
        "gemini-3.1-pro-low": { maxTokens: 1_048_576, supportsThinking: true },
        "gemini-pro-agent": { maxTokens: 1_048_576, supportsThinking: true },
        "gemini-3.1-flash-image": { maxTokens: 1_048_576, supportsImages: true },
        "claude-sonnet-4-6": { maxTokens: 250_000 },
      },
      agentModelSorts: [{ groups: [{ modelIds: ["gemini-3.7-flash-tiered", "gemini-3.1-pro-low", "gemini-pro-agent", "claude-sonnet-4-6"] }] }],
      imageGenerationModelIds: ["gemini-3.1-flash-image"],
      tieredModelIds: { flash: ["gemini-3.7-flash-tiered"] },
    };
    expect(parseAntigravityAvailableModels(discoveryPayload, "google-antigravity").map(model => model.id)).toEqual(expect.arrayContaining([
      "gemini-3.7-flash",
      "gemini-3.1-pro",
      "gemini-3.1-flash-image",
      "claude-sonnet-4-6",
    ]));

    const requests: Request[] = [];
    const fetchImpl: typeof fetch = async (input, init) => {
      const request = new Request(input, init);
      requests.push(request);
      if (request.url.endsWith("/v1internal:fetchAvailableModels")) return Response.json(discoveryPayload);
      return new Response('data: {"response":{"candidates":[{"content":{"parts":[{"text":"ready"}]}}]}}\n\ndata: [DONE]\n\n', { headers: { "content-type": "text/event-stream" } });
    };
    const provider = createProvider({ id: "google-antigravity", displayName: "Google Antigravity", kind: "google-antigravity", baseUrl: "https://daily-cloudcode-pa.googleapis.com", enabled: true, local: false, auth: { method: "oauth", oauthProvider: "google-antigravity", oauthCredentialRef: "oauth:google-antigravity" }, privacyNote: "test" }, {
      store: new MemoryCredentialStore(),
      fetchImpl,
      oauth: async () => ({ accessToken: "google-token", projectId: "project-123" }),
    });
    const models = await provider.listModels();
    expect(models.some(model => model.id === "gemini-3.7-flash")).toBe(true);
    for await (const _event of provider.streamChat({
      model: "gemini-3.7-flash",
      effort: "high",
      messages: [{ role: "user", content: "Review" }],
      tools: [{ type: "function", function: { name: "get_selection", description: "Read", parameters: { type: "object" } } }],
    })) { /* drain */ }
    for await (const _event of provider.streamChat({
      model: "claude-sonnet-4-6",
      messages: [{ role: "user", content: "Review" }, { role: "assistant", content: "Working" }],
      tools: [{ type: "function", function: { name: "get_selection", description: "Read", parameters: { type: "object" } } }],
    })) { /* drain */ }
    expect(requests.every(request => request.headers.get("user-agent") === antigravityUserAgent({}))).toBe(true);
    expect(requests.every(request => request.headers.get("x-goog-api-client") === null)).toBe(true);
    const geminiBody = JSON.parse(await requests[1]!.text()) as { model?: string; request?: { generationConfig?: { thinkingConfig?: { thinkingLevel?: string } }; toolConfig?: { functionCallingConfig?: { mode?: string } } } };
    expect(geminiBody.model).toBe("gemini-3.7-flash-tiered");
    expect(geminiBody.request?.generationConfig?.thinkingConfig?.thinkingLevel).toBe("high");
    expect(geminiBody.request?.toolConfig?.functionCallingConfig?.mode).toBe("AUTO");
    const claudeBody = JSON.parse(await requests[2]!.text()) as { request?: { contents?: Array<{ role?: string; parts?: Array<{ text?: string }> }>; toolConfig?: { functionCallingConfig?: { mode?: string } } } };
    expect(claudeBody.request?.toolConfig?.functionCallingConfig?.mode).toBe("VALIDATED");
    expect(claudeBody.request?.contents?.at(-1)).toEqual({ role: "user", parts: [{ text: "(continue)" }] });
  });

  it("refreshes Antigravity discovery and retries once when Google replaces a wire model", async () => {
    const discovery = (wireId: string) => ({
      models: { [wireId]: { displayName: "Gemini 3.7 Flash", maxTokens: 1_048_576, supportsThinking: true, supportsImages: true } },
      agentModelSorts: [{ groups: [{ modelIds: [wireId] }] }],
    });
    const requests: Request[] = [];
    let discoveryCount = 0;
    const provider = createProvider({ id: "antigravity-rollout", displayName: "Google Antigravity", kind: "google-antigravity", baseUrl: "https://antigravity-rollout.test", enabled: true, local: false, auth: { method: "oauth", oauthProvider: "google-antigravity", oauthCredentialRef: "oauth:google-antigravity" }, privacyNote: "test" }, {
      store: new MemoryCredentialStore(),
      oauth: async () => ({ accessToken: "google-token", projectId: "project-123" }),
      fetchImpl: async (input, init) => {
        const request = new Request(input, init);
        requests.push(request);
        if (request.url.endsWith("/v1internal:fetchAvailableModels")) {
          discoveryCount += 1;
          return Response.json(discovery(discoveryCount === 1 ? "gemini-flash-rollout-a" : "gemini-flash-rollout-b"));
        }
        const body = JSON.parse(await request.clone().text()) as { model?: string };
        if (body.model === "gemini-flash-rollout-a") return new Response("model retired", { status: 404 });
        return new Response('data: {"response":{"candidates":[{"content":{"parts":[{"text":"ready"}]}}]}}\n\ndata: [DONE]\n\n', { headers: { "content-type": "text/event-stream" } });
      },
    });
    expect((await provider.listModels()).map(model => model.id)).toContain("gemini-3.7-flash");
    const events = [];
    for await (const event of provider.streamChat({ model: "gemini-3.7-flash", effort: "high", messages: [{ role: "user", content: "Review" }] })) events.push(event);
    const streamedBodies = await Promise.all(requests
      .filter(request => request.url.includes("streamGenerateContent"))
      .map(async request => JSON.parse(await request.clone().text()) as { model?: string }));
    expect(streamedBodies.map(body => body.model)).toEqual(["gemini-flash-rollout-a", "gemini-flash-rollout-b"]);
    expect(events.some(event => event.type === "text" && event.delta === "ready")).toBe(true);
  });
});

describe("ChatGPT/Codex account provider", () => {
  it("uses the signed-in account catalog and filters hidden or unsupported models", async () => {
    const payload = {
      models: [
        { slug: "gpt-5.6-luna", display_name: "GPT-5.6 Luna", supported_in_api: true, visibility: "show", context_window: 373_000 },
        { slug: "gpt-hidden", supported_in_api: true, visibility: "hide" },
        { slug: "gpt-web-only", supported_in_api: false, visibility: "show" },
      ],
    };
    expect(parseChatGptAccountModels(payload, "openwordcode-account").map(model => model.id)).toEqual(["gpt-5.6-luna"]);
    const seen: Request[] = [];
    const provider = createProvider({ id: "openwordcode-account", displayName: "OpenWordCode account", kind: "openai-codex", baseUrl: "https://chatgpt.com/backend-api/codex", enabled: true, local: false, auth: { method: "oauth", credentialRef: "provider:openwordcode-account" }, privacyNote: "test" }, {
      store: new MemoryCredentialStore(),
      fetchImpl: async (input, init) => {
        seen.push(new Request(input, init));
        return Response.json(payload);
      },
      oauth: async () => ({ accessToken: "oauth-access", accountId: "acct-test", apiBaseUrl: "https://chatgpt.example.test/backend-api/codex" }),
    });
    expect((await provider.listModels()).map(model => model.id)).toEqual(["gpt-5.6-luna"]);
    expect(seen[0]?.url).toBe("https://chatgpt.example.test/backend-api/codex/models?client_version=0.0.0");
    expect(seen[0]?.headers.get("chatgpt-account-id")).toBe("acct-test");
  });

  it("uses the OAuth credential resolver and Responses-style streaming messages", async () => {
    const seen: Request[] = [];
    const fetchImpl: typeof fetch = async (input, init) => {
      const request = new Request(input, init);
      seen.push(request);
      return new Response([
        'data: {"type":"response.output_text.delta","delta":"hello"}',
        'data: {"type":"response.output_item.added","item":{"type":"function_call","call_id":"call-1","name":"get_selection","arguments":""}}',
        'data: {"type":"response.function_call_arguments.delta","item_id":"call-1","delta":"{}"}',
        'data: {"type":"response.output_item.done","item":{"type":"function_call","call_id":"call-1","name":"get_selection","arguments":"{}"}}',
        'data: {"type":"response.completed","response":{"output":[]}}',
        "data: [DONE]",
        "",
      ].join("\n\n"), { headers: { "content-type": "text/event-stream" } });
    };
    const provider = createProvider({ id: "openwordcode-account", displayName: "OpenWordCode account", kind: "openai-codex", baseUrl: "https://chatgpt.com/backend-api/codex", enabled: true, local: false, auth: { method: "oauth", credentialRef: "provider:openwordcode-account" }, privacyNote: "test" }, {
      store: new MemoryCredentialStore(),
      fetchImpl,
      oauth: async () => ({ accessToken: "oauth-access", accountId: "acct-test", apiBaseUrl: "https://chatgpt.example.test/backend-api/codex" }),
    });
    const events = [];
    for await (const event of provider.streamChat({ model: "gpt-5.6-luna", effort: "high", messages: [{ role: "system", content: "Be concise" }, { role: "user", content: "Review this" }, { role: "tool", toolCallId: "call-0", content: "done" }], tools: [{ type: "function", function: { name: "get_selection", description: "Read selection", parameters: { type: "object" } } }] })) events.push(event);
    expect(seen[0]?.url).toBe("https://chatgpt.example.test/backend-api/codex/responses");
    expect(seen[0]?.headers.get("authorization")).toBe("Bearer oauth-access");
    expect(seen[0]?.headers.get("chatgpt-account-id")).toBe("acct-test");
    const body = JSON.parse(await seen[0]!.text()) as { instructions?: string; input: unknown[]; tools: unknown[]; tool_choice?: string; reasoning?: { effort?: string } };
    expect(body.instructions).toBe("Be concise");
    expect(body.input).toEqual([
      { type: "message", role: "user", content: [{ type: "input_text", text: "Review this" }] },
      { type: "function_call_output", call_id: "call-0", output: "done" },
    ]);
    expect(body.tools).toEqual([{ type: "function", name: "get_selection", description: "Read selection", parameters: { type: "object" } }]);
    expect(body.reasoning).toEqual({ effort: "high" });
    expect(events.some(event => event.type === "text" && event.delta === "hello")).toBe(true);
    expect(events.some(event => event.type === "tool_call" && event.call.id === "call-1" && event.call.name === "get_selection")).toBe(true);
  });
});
