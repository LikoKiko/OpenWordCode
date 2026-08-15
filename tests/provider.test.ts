import { describe, expect, it } from "vitest";
import { MemoryCredentialStore } from "../packages/auth/src/index.js";
import { createProvider } from "../packages/providers/src/index.js";

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
    for await (const event of provider.streamChat({ model: "model-a", messages: [{ role: "user", content: "Review these files" }], attachments: [{ id: "image-1", name: "chart.png", mimeType: "image/png", size: 1, dataUrl: "data:image/png;base64,AA==" }, { id: "pdf-1", name: "brief.pdf", mimeType: "application/pdf", size: 1, dataUrl: "data:application/pdf;base64,AA==" }] })) events.push(event);
    expect(events.some(event => event.type === "text" && event.delta === "hello")).toBe(true);
    expect(seen.every(request => request.headers.get("authorization") === "Bearer provider-secret")).toBe(true);
    const chatBody = JSON.parse(await seen[1]!.text()) as { messages: Array<{ content: unknown }> };
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

  it("sends Gemini function declarations and returns streamed function calls", async () => {
    const store = new MemoryCredentialStore();
    await store.set("provider:gemini-test", "gemini-secret");
    let requestBody: { contents?: unknown[]; tools?: unknown[] } | undefined;
    const fetchImpl: typeof fetch = async (_input, init) => {
      requestBody = JSON.parse(String(init?.body)) as typeof requestBody;
      return new Response('data: {"candidates":[{"content":{"parts":[{"functionCall":{"name":"run_console","args":{"command":"pwd"}}}]}}]}\n\ndata: [DONE]\n\n', { headers: { "content-type": "text/event-stream" } });
    };
    const provider = createProvider({ id: "gemini-test", displayName: "Gemini Test", kind: "gemini", baseUrl: "http://127.0.0.1:9999/v1beta", enabled: true, local: true, auth: { method: "api-key", credentialRef: "provider:gemini-test" }, privacyNote: "test" }, { store, fetchImpl });
    const events = [];
    for await (const event of provider.streamChat({
      model: "gemini-test",
      messages: [
        { role: "user", content: "Inspect the workspace" },
        { role: "assistant", content: "", toolCalls: [{ id: "call-1", name: "run_console", arguments: '{"command":"pwd"}' }] },
        { role: "tool", toolCallId: "call-1", content: '{"output":"C:\\\\workspace"}' },
      ],
      tools: [{ type: "function", function: { name: "run_console", description: "Inspect safely", parameters: { type: "object" } } }],
    })) events.push(event);
    expect(requestBody?.tools).toEqual([{ functionDeclarations: [{ name: "run_console", description: "Inspect safely", parameters: { type: "object" } }] }]);
    expect((requestBody?.contents?.[1] as { parts?: unknown[] })?.parts?.[0]).toEqual({ functionCall: { name: "run_console", args: { command: "pwd" } } });
    expect(events.some(event => event.type === "tool_call" && event.call.name === "run_console")).toBe(true);
  });
});

describe("ChatGPT/Codex account provider", () => {
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
      oauth: async () => ({ accessToken: "oauth-access", accountId: "acct-test" }),
    });
    const events = [];
    for await (const event of provider.streamChat({ model: "gpt-5.6-luna", messages: [{ role: "system", content: "Be concise" }, { role: "user", content: "Review this" }, { role: "tool", toolCallId: "call-0", content: "done" }], tools: [{ type: "function", function: { name: "get_selection", description: "Read selection", parameters: { type: "object" } } }] })) events.push(event);
    expect(seen[0]?.url).toBe("https://chatgpt.com/backend-api/codex/responses");
    expect(seen[0]?.headers.get("authorization")).toBe("Bearer oauth-access");
    expect(seen[0]?.headers.get("chatgpt-account-id")).toBe("acct-test");
    const body = JSON.parse(await seen[0]!.text()) as { instructions?: string; input: unknown[]; tools: unknown[] };
    expect(body.instructions).toBe("Be concise");
    expect(body.input).toEqual([
      { type: "message", role: "user", content: [{ type: "input_text", text: "Review this" }] },
      { type: "function_call_output", call_id: "call-0", output: "done" },
    ]);
    expect(body.tools).toEqual([{ type: "function", name: "get_selection", description: "Read selection", parameters: { type: "object" } }]);
    expect(events.some(event => event.type === "text" && event.delta === "hello")).toBe(true);
    expect(events.some(event => event.type === "tool_call" && event.call.id === "call-1" && event.call.name === "get_selection")).toBe(true);
  });
});
