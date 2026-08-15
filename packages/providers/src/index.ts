import { randomUUID } from "node:crypto";
import {
  type ChatMessage,
  type ChatAttachment,
  type ChatToolDefinition,
  type ModelInfo,
  type ProviderConfig,
  type ProviderId,
  type ToolCall,
} from "../../shared/src/index.js";
import { type CredentialStore } from "../../auth/src/index.js";
import { normalizeBaseUrl, safeErrorMessage } from "../../security/src/index.js";

export interface ProviderChatRequest {
  model: string;
  messages: ChatMessage[];
  tools?: ChatToolDefinition[];
  attachments?: ChatAttachment[];
  signal?: AbortSignal;
}

export type ProviderStreamEvent =
  | { type: "text"; delta: string }
  | { type: "tool_call"; call: ToolCall }
  | { type: "done" };

export interface ProviderRuntime {
  readonly config: ProviderConfig;
  listModels(signal?: AbortSignal): Promise<ModelInfo[]>;
  streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent>;
}

export class ProviderError extends Error {
  constructor(readonly code: string, message: string, readonly status?: number) {
    super(message);
    this.name = "ProviderError";
  }
}

type JsonObject = Record<string, unknown>;

function isObject(value: unknown): value is JsonObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function arrayValue(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

function providerResponseError(config: ProviderConfig, response: Response): ProviderError {
  const status = response.status;
  if (status === 401 || status === 403) return new ProviderError("auth_invalid", `${config.displayName} rejected the configured authentication`, status);
  if (status === 429) return new ProviderError("rate_limited", `${config.displayName} rate-limited the request`, status);
  return new ProviderError("provider_http_error", `${config.displayName} returned HTTP ${status}`, status);
}

async function requireResponse(config: ProviderConfig, response: Response): Promise<Response> {
  if (response.ok) return response;
  throw providerResponseError(config, response);
}

async function parseJson(response: Response): Promise<unknown> {
  try { return await response.json() as unknown; } catch { throw new ProviderError("invalid_provider_response", "The provider returned malformed JSON"); }
}

async function* sseJson(response: Response): AsyncGenerator<unknown> {
  if (!response.body) throw new ProviderError("empty_provider_response", "The provider closed the stream without a body");
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  try {
    while (true) {
      const chunk = await reader.read();
      if (chunk.done) break;
      buffer += decoder.decode(chunk.value, { stream: true });
      const frames = buffer.split(/\r?\n\r?\n/);
      buffer = frames.pop() ?? "";
      for (const frame of frames) {
        const data = frame.split(/\r?\n/)
          .filter(line => line.startsWith("data:"))
          .map(line => line.slice(5).trim())
          .join("\n");
        if (!data || data === "[DONE]") continue;
        try { yield JSON.parse(data) as unknown; } catch { /* malformed keep-alive frame; fail closed at the provider boundary */ }
      }
    }
    buffer += decoder.decode();
    const data = buffer.split(/\r?\n/)
      .filter(line => line.startsWith("data:"))
      .map(line => line.slice(5).trim())
      .join("\n");
    if (data && data !== "[DONE]") {
      try { yield JSON.parse(data) as unknown; } catch { /* incomplete terminal frame */ }
    }
  } finally {
    reader.releaseLock();
  }
}

interface ResolvedCredential {
  token: string;
  oauth: boolean;
  apiBaseUrl?: string;
  projectId?: string;
}

async function readCredential(config: ProviderConfig, store: CredentialStore, env: NodeJS.ProcessEnv, oauth?: OAuthCredentialResolver): Promise<ResolvedCredential | null> {
  if (config.auth.method === "none") return null;
  if (config.auth.method === "oauth") {
    if (!oauth) throw new ProviderError("oauth_unavailable", `${config.displayName} OAuth is not configured in Core`);
    const value = await oauth(config);
    return value ? { token: value.accessToken, oauth: true, ...(value.apiBaseUrl ? { apiBaseUrl: value.apiBaseUrl } : {}), ...(value.projectId ? { projectId: value.projectId } : {}) } : null;
  }
  if (config.auth.method === "environment") {
    const variable = config.auth.envVar?.trim();
    const token = variable ? env[variable]?.trim() || null : null;
    return token ? { token, oauth: false } : null;
  }
  if (config.auth.method === "api-key") {
    const token = config.auth.credentialRef ? await store.get(config.auth.credentialRef) : null;
    return token ? { token, oauth: false } : null;
  }
  return null;
}

function authHeaders(config: ProviderConfig, credential: ResolvedCredential | null, env: NodeJS.ProcessEnv = process.env): Record<string, string> {
  if (config.auth.method === "api-key" || config.auth.method === "environment" || config.auth.method === "oauth") {
    if (!credential) throw new ProviderError("auth_required", `${config.displayName} needs a configured credential or account sign-in`);
    return { Authorization: `Bearer ${credential.token}` };
  }
  if (config.kind === "openwordcode-bridge") {
    const bridgeToken = env.OPENWORDCODE_BRIDGE_TOKEN?.trim();
    return bridgeToken ? { "x-openwordcode-bridge-token": bridgeToken } : {};
  }
  return {};
}

function modelFromRow(providerId: ProviderId, row: unknown): ModelInfo | null {
  if (!isObject(row)) return null;
  const id = stringValue(row.id) ?? stringValue(row.name);
  if (!id) return null;
  const capabilities = isObject(row.capabilities) ? row.capabilities : {};
  return {
    id,
    providerId,
    name: stringValue(row.name) ?? id,
    contextWindow: typeof row.context_length === "number" ? row.context_length : undefined,
    capabilities: {
      streaming: row.supports_streaming !== false,
      tools: capabilities.tools !== false && row.supports_tools !== false,
      reasoning: row.supports_reasoning_effort === true || row.reasoning_effort !== undefined,
      vision: row.supports_vision === true,
    },
  };
}

function extractText(value: unknown): string {
  if (typeof value === "string") return value;
  if (!Array.isArray(value)) return "";
  return value.map(part => isObject(part) ? stringValue(part.text) ?? "" : "").join("");
}

function attachmentBase64(dataUrl: string): string {
  const separator = dataUrl.indexOf(",");
  return separator >= 0 ? dataUrl.slice(separator + 1) : dataUrl;
}

const ANTHROPIC_SYSTEM_IDENTITY = "You are a Claude agent, built on Anthropic's Claude Agent SDK.";

function lastUserMessage(messages: ChatMessage[]): ChatMessage | undefined {
  return [...messages].reverse().find(message => message.role === "user");
}

function openAiMessageContent(message: ChatMessage, attachments: ChatAttachment[], includeAttachments: boolean): unknown {
  if (!includeAttachments || !attachments.length) return message.content;
  return [
    { type: "text", text: message.content },
    ...attachments.map(file => file.mimeType.startsWith("image/")
      ? { type: "image_url", image_url: { url: file.dataUrl } }
      : { type: "file", file: { filename: file.name, file_data: file.dataUrl } }),
  ];
}

function anthropicMessageContent(message: ChatMessage, attachments: ChatAttachment[], includeAttachments: boolean): unknown {
  if (!includeAttachments || !attachments.length) return message.content;
  return [
    { type: "text", text: message.content },
    ...attachments.map(file => file.mimeType.startsWith("image/")
      ? { type: "image", source: { type: "base64", media_type: file.mimeType, data: attachmentBase64(file.dataUrl) } }
      : { type: "document", source: { type: "base64", media_type: file.mimeType, data: attachmentBase64(file.dataUrl) }, title: file.name }),
  ];
}

function geminiMessageParts(message: ChatMessage, attachments: ChatAttachment[], includeAttachments: boolean): unknown[] {
  if (!includeAttachments || !attachments.length) return [{ text: message.content }];
  return [
    { text: message.content },
    ...attachments.map(file => ({ inline_data: { mime_type: file.mimeType, data: attachmentBase64(file.dataUrl) } })),
  ];
}

function anthropicMessages(messages: ChatMessage[], attachments: ChatAttachment[], oauth = false): unknown[] {
  const lastUser = lastUserMessage(messages);
  return messages
    .filter(message => message.role !== "system")
    .map(message => {
      if (message.role === "assistant") {
        const content: unknown[] = [];
        if (message.content) content.push({ type: "text", text: message.content });
        for (const call of message.toolCalls ?? []) {
          let input: unknown = {};
          try { input = JSON.parse(call.arguments) as unknown; } catch { /* keep an empty object for malformed provider arguments */ }
          content.push({ type: "tool_use", id: call.id, name: oauth ? `custom_${call.name}` : call.name, input });
        }
        return { role: "assistant", content: content.length ? content : [{ type: "text", text: "" }] };
      }
      if (message.role === "tool") {
        return { role: "user", content: [{ type: "tool_result", tool_use_id: message.toolCallId ?? "unknown-tool", content: message.content }] };
      }
      return { role: "user", content: anthropicMessageContent(message, attachments, message === lastUser) };
    });
}

function jsonValue(value: string): unknown {
  try { return JSON.parse(value) as unknown; } catch { return { output: value }; }
}

function geminiContents(messages: ChatMessage[], attachments: ChatAttachment[]): unknown[] {
  const lastUser = lastUserMessage(messages);
  const toolNames = new Map<string, string>();
  for (const message of messages) for (const call of message.toolCalls ?? []) toolNames.set(call.id, call.name);
  return messages
    .filter(message => message.role !== "system")
    .map(message => {
      if (message.role === "assistant") {
        const parts: unknown[] = [];
        if (message.content) parts.push({ text: message.content });
        for (const call of message.toolCalls ?? []) parts.push({ functionCall: { name: call.name, args: jsonValue(call.arguments) } });
        return { role: "model", parts: parts.length ? parts : [{ text: "" }] };
      }
      if (message.role === "tool") {
        const name = toolNames.get(message.toolCallId ?? "") ?? message.toolCallId ?? "tool";
        return { role: "user", parts: [{ functionResponse: { name, response: jsonValue(message.content) } }] };
      }
      return { role: "user", parts: geminiMessageParts(message, attachments, message === lastUser) };
    });
}

function responsesAttachmentContent(file: ChatAttachment): JsonObject {
  if (file.mimeType.startsWith("image/")) return { type: "input_image", image_url: file.dataUrl };
  return { type: "input_file", filename: file.name, file_data: file.dataUrl };
}

function responsesInput(messages: ChatMessage[], attachments: ChatAttachment[]): unknown[] {
  const lastUser = lastUserMessage(messages);
  return messages
    .filter(message => message.role !== "system")
    .flatMap(message => {
      if (message.role === "tool") return [{ type: "function_call_output", call_id: message.toolCallId ?? "unknown-tool", output: message.content }];
      if (message.role === "assistant") {
        const items: unknown[] = [];
        if (message.content) items.push({ type: "message", role: "assistant", content: [{ type: "output_text", text: message.content }] });
        for (const call of message.toolCalls ?? []) items.push({ type: "function_call", call_id: call.id, name: call.name, arguments: call.arguments });
        return items;
      }
      return [{
        type: "message",
        role: "user",
        content: [
          { type: "input_text", text: message.content },
          ...(message === lastUser ? attachments.map(responsesAttachmentContent) : []),
        ],
      }];
    });
}

function responseToolDefinitions(tools: ChatToolDefinition[]): unknown[] {
  return tools.map(tool => ({ type: "function", name: tool.function.name, description: tool.function.description, parameters: tool.function.parameters }));
}

function responseFunctionCall(value: unknown): ToolCall | null {
  if (!isObject(value) || value.type !== "function_call") return null;
  const name = stringValue(value.name);
  if (!name) return null;
  return {
    id: stringValue(value.call_id) ?? stringValue(value.id) ?? `tool-${Math.random().toString(16).slice(2)}`,
    name,
    arguments: stringValue(value.arguments) ?? "",
  };
}

function responseOutputText(value: unknown): string {
  if (!isObject(value)) return "";
  const outputText = stringValue(value.output_text);
  if (outputText) return outputText;
  const output = arrayValue(value.output);
  return output.flatMap(item => {
    if (!isObject(item) || item.type !== "message") return [];
    return arrayValue(item.content).flatMap(part => isObject(part) && part.type === "output_text" && typeof part.text === "string" ? [part.text] : []);
  }).join("");
}

function responseError(value: unknown): string | null {
  if (!isObject(value)) return null;
  const error = isObject(value.error) ? value.error : value;
  return stringValue(error.message) ?? stringValue(error.code) ?? null;
}

export class OpenAICompatibleProvider implements ProviderRuntime {
  constructor(
    readonly config: ProviderConfig,
    private readonly store: CredentialStore,
    private readonly env: NodeJS.ProcessEnv = process.env,
    private readonly fetchImpl: typeof fetch = fetch,
    private readonly oauth?: OAuthCredentialResolver,
  ) {}

  async listModels(signal?: AbortSignal): Promise<ModelInfo[]> {
    const credential = await readCredential(this.config, this.store, this.env, this.oauth);
    const baseUrl = normalizeBaseUrl(credential?.apiBaseUrl ?? this.config.baseUrl);
    const response = await this.fetchImpl(`${baseUrl}/models`, {
      headers: { Accept: "application/json", ...authHeaders(this.config, credential, this.env) },
      signal,
    });
    await requireResponse(this.config, response);
    const payload = await parseJson(response);
    const rows = isObject(payload) ? arrayValue(payload.data) : [];
    return rows.map(row => modelFromRow(this.config.id, row)).filter((model): model is ModelInfo => model !== null);
  }

  async *streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent> {
    const credential = await readCredential(this.config, this.store, this.env, this.oauth);
    const baseUrl = normalizeBaseUrl(credential?.apiBaseUrl ?? this.config.baseUrl);
    const lastUser = lastUserMessage(request.messages);
    const body: JsonObject = {
      model: request.model,
      messages: request.messages.map(message => ({
        role: message.role,
        content: openAiMessageContent(message, request.attachments ?? [], message === lastUser),
        ...(message.toolCallId ? { tool_call_id: message.toolCallId } : {}),
        ...(message.toolCalls ? { tool_calls: message.toolCalls.map(call => ({ id: call.id, type: "function", function: { name: call.name, arguments: call.arguments } })) } : {}),
      })),
      stream: true,
    };
    if (request.tools?.length) {
      body.tools = request.tools;
      body.tool_choice = "auto";
    }
    const response = await this.fetchImpl(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "text/event-stream", ...authHeaders(this.config, credential, this.env) },
      body: JSON.stringify(body),
      signal: request.signal,
    });
    await requireResponse(this.config, response);
    const contentType = response.headers.get("content-type") ?? "";
    if (!contentType.includes("text/event-stream")) {
      const payload = await parseJson(response);
      const choice = isObject(payload) && Array.isArray(payload.choices) ? payload.choices[0] : undefined;
      const message = isObject(choice) ? choice.message : undefined;
      if (isObject(message)) {
        const text = extractText(message.content);
        if (text) yield { type: "text", delta: text };
        for (const call of parseToolCalls(message.tool_calls)) yield { type: "tool_call", call };
      }
      yield { type: "done" };
      return;
    }
    const toolBuffers = new Map<number, ToolCall>();
    for await (const raw of sseJson(response)) {
      if (!isObject(raw)) continue;
      const choices = arrayValue(raw.choices);
      const choice = choices[0];
      if (!isObject(choice)) continue;
      const delta = isObject(choice.delta) ? choice.delta : {};
      const text = extractText(delta.content);
      if (text) yield { type: "text", delta: text };
      for (const rawCall of arrayValue(delta.tool_calls)) {
        if (!isObject(rawCall)) continue;
        const index = typeof rawCall.index === "number" ? rawCall.index : 0;
        const functionValue = isObject(rawCall.function) ? rawCall.function : {};
        const existing = toolBuffers.get(index) ?? {
          id: stringValue(rawCall.id) ?? `tool-${index}`,
          name: stringValue(functionValue.name) ?? "",
          arguments: "",
        };
        const nameDelta = stringValue(functionValue.name) ?? "";
        // Some OpenAI-compatible bridges repeat the complete function name on every
        // streamed tool delta. Treat repeated/cumulative names as one name; still allow
        // genuinely fragmented names from stricter providers.
        if (nameDelta) {
          if (!existing.name) existing.name = nameDelta;
          else if (nameDelta.startsWith(existing.name)) existing.name = nameDelta;
          else if (!existing.name.endsWith(nameDelta)) existing.name += nameDelta;
        }
        existing.arguments += stringValue(functionValue.arguments) ?? "";
        if (stringValue(rawCall.id)) existing.id = stringValue(rawCall.id) ?? existing.id;
        toolBuffers.set(index, existing);
      }
    }
    for (const call of [...toolBuffers.values()]) {
      if (call.name) yield { type: "tool_call", call };
    }
    yield { type: "done" };
  }
}

function parseToolCalls(value: unknown): ToolCall[] {
  return arrayValue(value).flatMap(item => {
    if (!isObject(item)) return [];
    const fn = isObject(item.function) ? item.function : {};
    const name = stringValue(fn.name);
    if (!name) return [];
    return [{ id: stringValue(item.id) ?? `tool-${Math.random().toString(16).slice(2)}`, name, arguments: stringValue(fn.arguments) ?? "{}" }];
  });
}

export class AnthropicProvider implements ProviderRuntime {
  constructor(
    readonly config: ProviderConfig,
    private readonly store: CredentialStore,
    private readonly env: NodeJS.ProcessEnv = process.env,
    private readonly fetchImpl: typeof fetch = fetch,
    private readonly oauth?: OAuthCredentialResolver,
  ) {}

  private async credential(): Promise<ResolvedCredential> {
    const value = await readCredential(this.config, this.store, this.env, this.oauth);
    if (!value) throw new ProviderError("auth_required", "Anthropic needs an API key or account sign-in");
    return value;
  }

  private oauthMode(): boolean { return this.config.auth.method === "oauth"; }

  private headers(credential: ResolvedCredential, stream: boolean): Record<string, string> {
    const headers: Record<string, string> = {
      Accept: stream ? "text/event-stream" : "application/json",
      "anthropic-version": "2023-06-01",
      "User-Agent": "@anthropic-ai/sdk/0.74.0",
    };
    if (credential.oauth) {
      Object.assign(headers, {
        Authorization: `Bearer ${credential.token}`,
        "anthropic-beta": "claude-code-20250219,oauth-2025-04-20",
        "X-App": "cli",
        "X-Stainless-Retry-Count": "0",
        "X-Stainless-Runtime": "node",
        "X-Stainless-Lang": "js",
        "X-Stainless-Timeout": "600",
        "X-Stainless-Package-Version": "0.74.0",
        "X-Claude-Code-Session-Id": `openwordcode-${credential.token.slice(0, 24)}`,
      });
    } else {
      headers["x-api-key"] = credential.token;
    }
    return headers;
  }

  async listModels(signal?: AbortSignal): Promise<ModelInfo[]> {
    const credential = await this.credential();
    const response = await this.fetchImpl(`${normalizeBaseUrl(this.config.baseUrl)}/models?limit=100`, {
      headers: this.headers(credential, false),
      signal,
    });
    await requireResponse(this.config, response);
    const payload = await parseJson(response);
    const rows = isObject(payload) ? arrayValue(payload.data) : [];
    return rows.map(row => modelFromRow(this.config.id, row)).filter((model): model is ModelInfo => model !== null)
      .map(model => ({ ...model, capabilities: { ...model.capabilities, tools: true, streaming: true } }));
  }

  async *streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent> {
    const credential = await this.credential();
    const oauth = this.oauthMode();
    const messages = anthropicMessages(request.messages, request.attachments ?? [], oauth);
    const system = request.messages.filter(message => message.role === "system").map(message => message.content).join("\n\n");
    const tools = request.tools?.map(tool => ({ name: oauth ? `custom_${tool.function.name}` : tool.function.name, description: tool.function.description, input_schema: tool.function.parameters })) ?? [];
    const body: JsonObject = { model: request.model, max_tokens: 4096, stream: true, messages };
    if (oauth) body.system = [ANTHROPIC_SYSTEM_IDENTITY, ...(system ? [system] : [])].join("\n\n");
    else if (system) body.system = system;
    if (tools.length) body.tools = tools;
    const response = await this.fetchImpl(`${normalizeBaseUrl(this.config.baseUrl)}/messages`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...this.headers(credential, true) },
      body: JSON.stringify(body),
      signal: request.signal,
    });
    await requireResponse(this.config, response);
    const toolCalls = new Map<number, ToolCall>();
    for await (const raw of sseJson(response)) {
      if (!isObject(raw)) continue;
      const eventType = stringValue(raw.type);
      if (eventType === "content_block_start" && isObject(raw.content_block) && raw.content_block.type === "tool_use") {
        const index = typeof raw.index === "number" ? raw.index : toolCalls.size;
        const wireName = stringValue(raw.content_block.name) ?? "";
        toolCalls.set(index, { id: stringValue(raw.content_block.id) ?? `tool-${index}`, name: oauth && wireName.startsWith("custom_") ? wireName.slice("custom_".length) : wireName, arguments: "" });
      }
      if (eventType === "content_block_delta" && isObject(raw.delta)) {
        if (raw.delta.type === "text_delta") {
          const delta = stringValue(raw.delta.text);
          if (delta) yield { type: "text", delta };
        }
        if (raw.delta.type === "input_json_delta") {
          const index = typeof raw.index === "number" ? raw.index : 0;
          const call = toolCalls.get(index);
          if (call) call.arguments += stringValue(raw.delta.partial_json) ?? "";
        }
      }
    }
    for (const call of toolCalls.values()) if (call.name) yield { type: "tool_call", call };
    yield { type: "done" };
  }
}

export class GeminiProvider implements ProviderRuntime {
  constructor(
    readonly config: ProviderConfig,
    private readonly store: CredentialStore,
    private readonly env: NodeJS.ProcessEnv = process.env,
    private readonly fetchImpl: typeof fetch = fetch,
  ) {}

  private async credential(): Promise<string> {
    const value = await readCredential(this.config, this.store, this.env);
    if (!value) throw new ProviderError("auth_required", "Google Gemini needs a GEMINI_API_KEY or configured API key");
    return value.token;
  }

  async listModels(signal?: AbortSignal): Promise<ModelInfo[]> {
    const key = await this.credential();
    const response = await this.fetchImpl(`${normalizeBaseUrl(this.config.baseUrl)}/models?key=${encodeURIComponent(key)}`, { signal, headers: { Accept: "application/json" } });
    await requireResponse(this.config, response);
    const payload = await parseJson(response);
    const rows = isObject(payload) ? arrayValue(payload.models) : [];
    return rows.flatMap(row => {
      if (!isObject(row)) return [];
      const rawName = stringValue(row.name);
      if (!rawName) return [];
      const id = rawName.replace(/^models\//, "");
      return [{ id, providerId: this.config.id, name: stringValue(row.displayName) ?? id, capabilities: { streaming: true, tools: true, vision: true } }];
    });
  }

  async *streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent> {
    const key = await this.credential();
    const system = request.messages.filter(message => message.role === "system").map(message => message.content).join("\n\n");
    const contents = geminiContents(request.messages, request.attachments ?? []);
    const body: JsonObject = { contents, generationConfig: { temperature: 0.2 } };
    if (system) body.systemInstruction = { parts: [{ text: system }] };
    if (request.tools?.length) {
      body.tools = [{ functionDeclarations: request.tools.map(tool => ({ name: tool.function.name, description: tool.function.description, parameters: tool.function.parameters })) }];
      body.toolConfig = { functionCallingConfig: { mode: "AUTO" } };
    }
    const response = await this.fetchImpl(`${normalizeBaseUrl(this.config.baseUrl)}/models/${encodeURIComponent(request.model)}:streamGenerateContent?alt=sse&key=${encodeURIComponent(key)}`, {
      method: "POST", headers: { "Content-Type": "application/json", Accept: "text/event-stream" }, body: JSON.stringify(body), signal: request.signal,
    });
    await requireResponse(this.config, response);
    const toolCalls = new Map<string, ToolCall>();
    for await (const raw of sseJson(response)) {
      if (!isObject(raw)) continue;
      const candidates = arrayValue(raw.candidates);
      const candidate = candidates[0];
      if (!isObject(candidate) || !isObject(candidate.content)) continue;
      for (const part of arrayValue(candidate.content.parts)) {
        if (!isObject(part)) continue;
        if (typeof part.text === "string") yield { type: "text", delta: part.text };
        if (isObject(part.functionCall)) {
          const name = stringValue(part.functionCall.name);
          if (!name) continue;
          const existing = toolCalls.get(name) ?? { id: `gemini-${toolCalls.size}`, name, arguments: "{}" };
          const previous = jsonValue(existing.arguments);
          const previousObject = isObject(previous) ? previous : {};
          const nextArguments = isObject(part.functionCall.args) ? { ...previousObject, ...part.functionCall.args } : previousObject;
          existing.arguments = JSON.stringify(nextArguments);
          toolCalls.set(name, existing);
        }
      }
    }
    for (const call of toolCalls.values()) yield { type: "tool_call", call };
    yield { type: "done" };
  }
}

const ANTIGRAVITY_MODELS = [
  "gemini-3.7-flash",
  "gemini-3.1-pro",
  "claude-sonnet-4-6",
  "claude-opus-4-6-thinking",
  "gpt-oss-120b-medium",
];

/** Minimal Cloud Code Assist runtime. Its OAuth token is not a Gemini API key. */
export class GoogleAntigravityProvider implements ProviderRuntime {
  constructor(
    readonly config: ProviderConfig,
    private readonly store: CredentialStore,
    private readonly env: NodeJS.ProcessEnv = process.env,
    private readonly fetchImpl: typeof fetch = fetch,
    private readonly oauth?: OAuthCredentialResolver,
  ) {}

  private async credential(): Promise<ResolvedCredential> {
    const value = await readCredential(this.config, this.store, this.env, this.oauth);
    if (!value) throw new ProviderError("auth_required", "Google Antigravity needs an account sign-in");
    if (!value.projectId) throw new ProviderError("oauth_setup_incomplete", "Google Antigravity has no Cloud Code Assist project yet; disconnect and sign in again");
    return value;
  }

  async listModels(): Promise<ModelInfo[]> {
    return ANTIGRAVITY_MODELS.map(id => ({
      id,
      providerId: this.config.id,
      name: id,
      capabilities: { streaming: true, tools: true, reasoning: /gemini|opus|sonnet/iu.test(id), vision: /image|gemini/iu.test(id) },
    }));
  }

  async *streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent> {
    const credential = await this.credential();
    const system = request.messages.filter(message => message.role === "system").map(message => message.content).join("\n\n");
    const contents = geminiContents(request.messages, request.attachments ?? []);
    const requestBody: JsonObject = { contents, sessionId: randomUUID() };
    if (system) requestBody.systemInstruction = { parts: [{ text: system }] };
    if (request.tools?.length) {
      requestBody.tools = [{ functionDeclarations: request.tools.map(tool => ({ name: tool.function.name, description: tool.function.description, parameters: tool.function.parameters })) }];
      requestBody.toolConfig = { functionCallingConfig: { mode: "VALIDATED" } };
    }
    const envelope = {
      model: request.model,
      userAgent: "antigravity",
      requestType: "agent",
      project: credential.projectId,
      requestId: `agent-${randomUUID()}`,
      request: requestBody,
    };
    const base = normalizeBaseUrl(this.config.baseUrl);
    const response = await this.fetchImpl(`${base}/v1internal:streamGenerateContent?alt=sse`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "text/event-stream",
        Authorization: `Bearer ${credential.token}`,
        "User-Agent": "antigravity/ide/2.5.5 (aidev_client; os_type=windows; arch=amd64)",
      },
      body: JSON.stringify(envelope),
      signal: request.signal,
    });
    await requireResponse(this.config, response);
    const toolCalls = new Map<string, ToolCall>();
    for await (const raw of sseJson(response)) {
      if (!isObject(raw)) continue;
      const root = isObject(raw.response) ? raw.response : raw;
      for (const candidate of arrayValue(root.candidates)) {
        if (!isObject(candidate) || !isObject(candidate.content)) continue;
        for (const part of arrayValue(candidate.content.parts)) {
          if (!isObject(part)) continue;
          if (typeof part.text === "string" && part.thought !== true) yield { type: "text", delta: part.text };
          if (isObject(part.functionCall)) {
            const name = stringValue(part.functionCall.name);
            if (!name) continue;
            const id = `antigravity-${toolCalls.size}`;
            toolCalls.set(id, { id, name, arguments: JSON.stringify(isObject(part.functionCall.args) ? part.functionCall.args : {}) });
          }
        }
      }
    }
    for (const call of toolCalls.values()) yield { type: "tool_call", call };
    yield { type: "done" };
  }
}

export interface ProviderOAuthCredential {
  accessToken: string;
  accountId?: string;
  apiBaseUrl?: string;
  projectId?: string;
}

export type OAuthCredentialResolver = (config: ProviderConfig) => Promise<ProviderOAuthCredential | null>;

/**
 * Experimental ChatGPT/Codex account runtime. The endpoint and OAuth token
 * contract are intentionally isolated from the normal OpenAI API-key runtime:
 * ChatGPT account sessions are not OpenAI Platform API keys.
 */
export class ChatGPTCodexProvider implements ProviderRuntime {
  constructor(
    readonly config: ProviderConfig,
    private readonly resolveCredential: OAuthCredentialResolver,
    private readonly fetchImpl: typeof fetch = fetch,
  ) {}

  async listModels(): Promise<ModelInfo[]> {
    return [
      { id: "gpt-5.6-sol", providerId: this.config.id, name: "GPT-5.6 Sol", capabilities: { streaming: true, tools: true, reasoning: true, vision: true } },
      { id: "gpt-5.6-terra", providerId: this.config.id, name: "GPT-5.6 Terra", capabilities: { streaming: true, tools: true, reasoning: true, vision: true } },
      { id: "gpt-5.6-luna", providerId: this.config.id, name: "GPT-5.6 Luna", capabilities: { streaming: true, tools: true, reasoning: true, vision: true } },
    ];
  }

  async *streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent> {
    const credential = await this.resolveCredential(this.config);
    if (!credential) throw new ProviderError("auth_required", "Sign in to the OpenWordCode account before sending a message");
    const instructions = request.messages.filter(message => message.role === "system").map(message => message.content).join("\n\n");
    const body: JsonObject = {
      model: request.model,
      input: responsesInput(request.messages, request.attachments ?? []),
      stream: true,
    };
    if (instructions) body.instructions = instructions;
    if (request.tools?.length) body.tools = responseToolDefinitions(request.tools);
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      Accept: "text/event-stream",
      Authorization: `Bearer ${credential.accessToken}`,
    };
    if (credential.accountId) headers["ChatGPT-Account-Id"] = credential.accountId;
    const response = await this.fetchImpl(`${normalizeBaseUrl(this.config.baseUrl)}/responses`, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
      signal: request.signal,
    });
    await requireResponse(this.config, response);
    const contentType = response.headers.get("content-type") ?? "";
    if (!contentType.includes("text/event-stream")) {
      const payload = await parseJson(response);
      const text = responseOutputText(payload);
      if (text) yield { type: "text", delta: text };
      const output = isObject(payload) ? arrayValue(payload.output) : [];
      for (const item of output) {
        const call = responseFunctionCall(item);
        if (call) yield { type: "tool_call", call: { ...call, arguments: call.arguments || "{}" } };
      }
      const error = responseError(payload);
      if (error) throw new ProviderError("provider_response_error", error);
      yield { type: "done" };
      return;
    }

    const toolCalls = new Map<string, ToolCall>();
    const itemKeys = new Map<string, string>();
    const emittedCalls = new Set<string>();
    for await (const raw of sseJson(response)) {
      if (!isObject(raw)) continue;
      const eventType = stringValue(raw.type);
      const error = eventType === "error" || eventType === "response.failed" ? responseError(raw) : null;
      if (error) throw new ProviderError("provider_response_error", error);
      if (eventType === "response.output_text.delta") {
        const delta = stringValue(raw.delta);
        if (delta) yield { type: "text", delta };
      }
      if (eventType === "response.output_item.added" || eventType === "response.output_item.done") {
        const item = isObject(raw.item) ? raw.item : undefined;
        const call = responseFunctionCall(item);
        if (call) {
          const itemId = stringValue(item?.id);
          const callId = stringValue(item?.call_id);
          const key = (itemId ? itemKeys.get(itemId) : undefined) ?? callId ?? itemId ?? call.id;
          if (itemId) itemKeys.set(itemId, key);
          const existing = toolCalls.get(key) ?? (callId ? toolCalls.get(callId) : undefined) ?? call;
          existing.id = call.id;
          existing.name = call.name;
          if (call.arguments !== "{}") existing.arguments = call.arguments;
          toolCalls.set(key, existing);
          if (eventType === "response.output_item.done" && !emittedCalls.has(key)) {
            emittedCalls.add(key);
            yield { type: "tool_call", call: { ...existing, arguments: existing.arguments || "{}" } };
          }
        }
      }
      if (eventType === "response.function_call_arguments.delta") {
        const itemId = stringValue(raw.item_id);
        const key = (itemId ? itemKeys.get(itemId) : undefined) ?? stringValue(raw.call_id) ?? itemId ?? `tool-${toolCalls.size}`;
        if (itemId) itemKeys.set(itemId, key);
        const existing = toolCalls.get(key) ?? { id: stringValue(raw.call_id) ?? key, name: stringValue(raw.name) ?? "", arguments: "" };
        if (!existing.name && stringValue(raw.name)) existing.name = stringValue(raw.name) ?? "";
        existing.arguments += stringValue(raw.delta) ?? "";
        toolCalls.set(key, existing);
      }
      if (eventType === "response.function_call_arguments.done") {
        const itemId = stringValue(raw.item_id);
        const key = (itemId ? itemKeys.get(itemId) : undefined) ?? stringValue(raw.call_id) ?? itemId ?? "";
        const existing = key ? toolCalls.get(key) : undefined;
        if (existing && stringValue(raw.arguments)) existing.arguments = stringValue(raw.arguments) ?? existing.arguments;
      }
      if (eventType === "response.completed") {
        const responseValue = isObject(raw.response) ? raw.response : undefined;
        for (const item of responseValue ? arrayValue(responseValue.output) : []) {
          const call = responseFunctionCall(item);
          if (!call) continue;
          const key = stringValue(isObject(item) ? item.call_id : undefined) ?? call.id;
          if (!emittedCalls.has(key)) {
            emittedCalls.add(key);
            yield { type: "tool_call", call: { ...call, arguments: call.arguments || "{}" } };
          }
        }
      }
    }
    for (const [key, call] of toolCalls) if (call.name && !emittedCalls.has(key)) yield { type: "tool_call", call: { ...call, arguments: call.arguments || "{}" } };
    yield { type: "done" };
  }
}

export class DemoProvider implements ProviderRuntime {
  constructor(readonly config: ProviderConfig) {}

  async listModels(): Promise<ModelInfo[]> {
    return [{ id: "demo-rewrite", providerId: this.config.id, name: "Demo Rewrite (offline)", capabilities: { streaming: true, tools: true } }];
  }

  async *streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent> {
    const user = [...request.messages].reverse().find(message => message.role === "user")?.content ?? "";
    const selectionMatch = /\[SELECTION\]\n([\s\S]*?)\n\[\/SELECTION\]/.exec(user);
    const selection = selectionMatch?.[1]?.trim() || "Select some text in Word to preview a proposed change.";
    const instruction = user.toLowerCase();
    let after = selection;
    if (instruction.includes("grammar")) {
      after = after.replace(/\bcustomers needs\b/gi, "customer needs").replace(/\bthey was\b/gi, "they were");
    } else if (instruction.includes("shorten")) {
      after = after.split(/(?<=[.!?])\s+/).slice(0, 1).join(" ") || after;
    } else if (instruction.includes("professional")) {
      after = after.replace(/\bcan't\b/gi, "cannot").replace(/\bwon't\b/gi, "will not").replace(/\bget\b/gi, "obtain");
    }
    const payload = JSON.stringify({
      answer: after === selection ? "I reviewed the selected text and prepared a focused revision." : "I prepared a focused revision of the selected text.",
      proposedChanges: after === selection ? [] : [{ type: "replace_text", targetId: "selection", before: selection, after, description: "Apply the requested wording improvement to the current selection." }],
    });
    for (let index = 0; index < payload.length; index += 32) yield { type: "text", delta: payload.slice(index, index + 32) };
    yield { type: "done" };
  }
}

export interface ProviderFactoryOptions {
  store: CredentialStore;
  env?: NodeJS.ProcessEnv;
  fetchImpl?: typeof fetch;
  oauth?: OAuthCredentialResolver;
}

export function createProvider(config: ProviderConfig, options: ProviderFactoryOptions): ProviderRuntime {
  if (config.kind === "demo") return new DemoProvider(config);
  if (config.kind === "openai-codex") {
    if (!options.oauth) throw new ProviderError("oauth_unavailable", "OpenWordCode account OAuth is not configured in Core");
    return new ChatGPTCodexProvider(config, options.oauth, options.fetchImpl);
  }
  if (config.kind === "google-antigravity") return new GoogleAntigravityProvider(config, options.store, options.env, options.fetchImpl, options.oauth);
  if (config.kind === "anthropic") return new AnthropicProvider(config, options.store, options.env, options.fetchImpl, options.oauth);
  if (config.kind === "gemini") return new GeminiProvider(config, options.store, options.env, options.fetchImpl);
  return new OpenAICompatibleProvider(config, options.store, options.env, options.fetchImpl, options.oauth);
}

export function defaultProviderConfigs(): ProviderConfig[] {
  return [
    {
      id: "openwordcode-bridge", displayName: "OpenWordCode Bridge", kind: "openwordcode-bridge", baseUrl: "http://127.0.0.1:10101/v1", enabled: true, local: true,
      auth: { method: "oauth", credentialRef: "provider:openwordcode-account" }, defaultModel: "gpt-5.6-luna", privacyNote: "Uses the local OpenWordCode Bridge. Its configured provider/account determines where document content is sent.",
    },
    {
      id: "openwordcode-account", displayName: "OpenWordCode account", kind: "openai-codex", baseUrl: "https://chatgpt.com/backend-api/codex", enabled: true, local: false, internal: true,
      auth: { method: "oauth", credentialRef: "provider:openwordcode-account" }, defaultModel: "gpt-5.6-luna", privacyNote: "Internal account transport used by the OpenWordCode Bridge.",
    },
    {
      id: "ollama", displayName: "Ollama", kind: "ollama", baseUrl: "http://127.0.0.1:11434/v1", enabled: true, local: true,
      auth: { method: "none" }, defaultModel: "", privacyNote: "Local model; document content remains on this machine unless the Ollama endpoint is remote.",
    },
    {
      id: "lm-studio", displayName: "LM Studio", kind: "lm-studio", baseUrl: "http://127.0.0.1:1234/v1", enabled: true, local: true,
      auth: { method: "none" }, defaultModel: "", privacyNote: "Local model; document content remains on this machine unless the endpoint is remote.",
    },
    {
      id: "openai", displayName: "OpenAI", kind: "openai-compatible", baseUrl: "https://api.openai.com/v1", enabled: true, local: false,
      auth: { method: "api-key", credentialRef: "provider:openai" }, defaultModel: "gpt-4.1-mini", privacyNote: "Selected document context is sent to OpenAI's API for the chosen model.",
    },
    {
      id: "anthropic", displayName: "Anthropic", kind: "anthropic", baseUrl: "https://api.anthropic.com/v1", enabled: true, local: false,
      auth: { method: "environment", envVar: "ANTHROPIC_API_KEY", oauthProvider: "anthropic", oauthCredentialRef: "oauth:anthropic" }, defaultModel: "claude-sonnet-4-6", privacyNote: "Selected document context is sent to Anthropic's API or Claude account transport for the chosen model.",
    },
    {
      id: "openrouter", displayName: "OpenRouter", kind: "openai-compatible", baseUrl: "https://openrouter.ai/api/v1", enabled: true, local: false,
      auth: { method: "environment", envVar: "OPENROUTER_API_KEY" }, defaultModel: "", privacyNote: "Selected document context is sent to OpenRouter and then to its routed model/provider.",
    },
    {
      id: "google", displayName: "Google Gemini", kind: "gemini", baseUrl: "https://generativelanguage.googleapis.com/v1beta", enabled: true, local: false,
      auth: { method: "environment", envVar: "GEMINI_API_KEY" }, defaultModel: "gemini-2.0-flash", privacyNote: "Selected document context is sent to Google's Gemini API for the chosen model.",
    },
    {
      id: "google-antigravity", displayName: "Google Antigravity", kind: "google-antigravity", baseUrl: "https://daily-cloudcode-pa.googleapis.com", enabled: true, local: false,
      auth: { method: "oauth", oauthProvider: "google-antigravity", oauthCredentialRef: "oauth:google-antigravity" }, defaultModel: "gemini-3.7-flash", privacyNote: "Selected document context is sent to Google Cloud Code Assist through the signed-in Antigravity account.",
    },
    {
      id: "xai", displayName: "xAI", kind: "openai-compatible", baseUrl: "https://api.x.ai/v1", enabled: true, local: false,
      auth: { method: "environment", envVar: "XAI_API_KEY", oauthProvider: "xai", oauthCredentialRef: "oauth:xai" }, defaultModel: "grok-4-1-fast-reasoning", privacyNote: "Selected document context is sent to xAI's API for the chosen model.",
    },
    {
      id: "kimi", displayName: "Kimi Code", kind: "openai-compatible", baseUrl: "https://api.kimi.com/coding/v1", enabled: true, local: false,
      auth: { method: "environment", envVar: "MOONSHOT_API_KEY", oauthProvider: "kimi", oauthCredentialRef: "oauth:kimi" }, defaultModel: "kimi-k2.7-code", privacyNote: "Selected document context is sent to Kimi Code or the configured Moonshot-compatible API for the chosen model.",
    },
    {
      id: "nous", displayName: "Nous Portal", kind: "openai-compatible", baseUrl: "https://inference-api.nousresearch.com/v1", enabled: true, local: false,
      auth: { method: "oauth", oauthProvider: "nous", oauthCredentialRef: "oauth:nous" }, defaultModel: "tencent/hy3:free", privacyNote: "Selected document context is sent to the Nous Portal inference service for the signed-in account.",
    },
    {
      id: "github-copilot", displayName: "GitHub Copilot", kind: "openai-compatible", baseUrl: "https://api.githubcopilot.com", enabled: true, local: false,
      auth: { method: "oauth", oauthProvider: "github-copilot", oauthCredentialRef: "oauth:github-copilot" }, defaultModel: "gpt-4o", privacyNote: "Selected document context is sent to GitHub Copilot using the signed-in GitHub account.",
    },
    {
      id: "demo", displayName: "Demo (offline)", kind: "demo", baseUrl: "demo://offline", enabled: true, local: true,
      auth: { method: "none" }, defaultModel: "demo-rewrite", privacyNote: "Offline deterministic demo; no document content leaves this process.",
    },
  ];
}

export function providerAuthUsesSecret(config: ProviderConfig): boolean {
  return config.auth.method === "api-key" || config.auth.method === "environment";
}

export function providerErrorDetail(error: unknown): string {
  return error instanceof ProviderError ? error.message : safeErrorMessage(error);
}
