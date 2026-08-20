import { createHash, randomUUID } from "node:crypto";
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
  effort?: string;
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

const GOOGLE_RETRYABLE_STATUSES = new Set([408, 429, 500, 502, 503, 504]);

function abortReason(signal: AbortSignal): Error {
  return signal.reason instanceof Error ? signal.reason : new DOMException("The request was aborted", "AbortError");
}

function sleepWithSignal(delayMs: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(abortReason(signal));
      return;
    }
    const timer = setTimeout(resolve, delayMs);
    signal?.addEventListener("abort", () => {
      clearTimeout(timer);
      reject(abortReason(signal));
    }, { once: true });
  });
}

function googleRetryDelay(response: Response | undefined, attempt: number): number {
  const retryAfter = response?.headers.get("retry-after")?.trim();
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds) && seconds >= 0) return Math.min(2_000, Math.floor(seconds * 1_000));
    const date = Date.parse(retryAfter);
    if (Number.isFinite(date)) return Math.min(2_000, Math.max(0, date - Date.now()));
  }
  return Math.min(2_000, 250 * (2 ** attempt));
}

/** Bounded retries for the transient status and network failures common to Google model endpoints. */
export async function fetchGoogleWithRetry(
  fetchImpl: typeof fetch,
  input: string | URL | Request,
  init: RequestInit,
): Promise<Response> {
  let lastError: unknown;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    if (init.signal?.aborted) throw abortReason(init.signal);
    let response: Response | undefined;
    try {
      response = await fetchImpl(input, init);
      if (!GOOGLE_RETRYABLE_STATUSES.has(response.status) || attempt === 2) return response;
      try { await response.body?.cancel(); } catch { /* best-effort connection cleanup */ }
    } catch (error) {
      if (init.signal?.aborted) throw error;
      lastError = error;
      if (attempt === 2) throw error;
    }
    await sleepWithSignal(googleRetryDelay(response, attempt), init.signal ?? undefined);
  }
  throw lastError instanceof Error ? lastError : new ProviderError("provider_network_error", "Google request failed");
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
  accountId?: string;
  apiBaseUrl?: string;
  projectId?: string;
}

async function readCredential(config: ProviderConfig, store: CredentialStore, env: NodeJS.ProcessEnv, oauth?: OAuthCredentialResolver): Promise<ResolvedCredential | null> {
  if (config.auth.method === "none") return null;
  if (config.auth.method === "oauth") {
    if (!oauth) throw new ProviderError("oauth_unavailable", `${config.displayName} OAuth is not configured in Core`);
    const value = await oauth(config);
    return value ? {
      token: value.accessToken,
      oauth: true,
      ...(value.accountId ? { accountId: value.accountId } : {}),
      ...(value.apiBaseUrl ? { apiBaseUrl: value.apiBaseUrl } : {}),
      ...(value.projectId ? { projectId: value.projectId } : {}),
    } : null;
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
    if (config.auth.oauthProvider === "kimi" && config.auth.method === "oauth") {
      return {
        Authorization: `Bearer ${credential.token}`,
        "User-Agent": "KimiCLI/0.14.0",
        "X-Msh-Platform": "kimi_cli",
        "X-Msh-Version": "0.14.0",
        "X-Msh-Device-Id": "openwordcode-local-cli",
      };
    }
    if (config.auth.oauthProvider === "github-copilot" && config.auth.method === "oauth") {
      return {
        Authorization: `Bearer ${credential.token}`,
        "Editor-Version": "openwordcode/0.1.0",
        "Editor-Plugin-Version": "openwordcode/0.1.0",
        "Copilot-Integration-Id": "vscode-chat",
        "User-Agent": "openwordcode",
      };
    }
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

/** Gemini requires function arguments and function responses to be JSON objects. */
export function jsonObjectValue(value: string): JsonObject {
  const parsed = jsonValue(value);
  return isObject(parsed) ? parsed : { output: parsed };
}

function googleToolCallId(value: string): string {
  const source = value.trim();
  const normalized = source.replace(/[^A-Za-z0-9_-]/g, "_");
  if (!normalized) return `call_${randomUUID().replace(/-/g, "").slice(0, 24)}`;
  if (normalized === source && normalized.length <= 96) return normalized;
  const digest = createHash("sha256").update(source, "utf8").digest("hex").slice(0, 12);
  return `${normalized.slice(0, 80)}_${digest}`;
}

/** Reject synthetic Responses/Anthropic IDs that Google cannot decode as thought signatures. */
export function isLikelyRealThoughtSignature(value: string | undefined): value is string {
  if (typeof value !== "string" || value.length < 16) return false;
  if (/^(fc|ctc|tsc|call|msg|rs|resp|reasoning|item|ws|toolu|tool|func|function)[-_]/iu.test(value)) return false;
  return /^[A-Za-z0-9+/_=-]+$/u.test(value);
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
        for (const call of message.toolCalls ?? []) {
          parts.push({
            functionCall: { id: googleToolCallId(call.id), name: call.name, args: jsonObjectValue(call.arguments) },
            ...(isLikelyRealThoughtSignature(call.thoughtSignature) ? { thoughtSignature: call.thoughtSignature } : {}),
          });
        }
        return { role: "model", parts: parts.length ? parts : [{ text: "" }] };
      }
      if (message.role === "tool") {
        const name = toolNames.get(message.toolCallId ?? "") ?? message.toolCallId ?? "tool";
        return { role: "user", parts: [{ functionResponse: { id: googleToolCallId(message.toolCallId ?? name), name, response: jsonObjectValue(message.content) } }] };
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

export const GITHUB_COPILOT_RESPONSES_MODELS: ReadonlySet<string> = new Set([
  "gpt-5.3-codex",
  "gpt-5.4",
  "gpt-5.4-mini",
  "gpt-5.5",
  "gpt-5.6-luna",
  "gpt-5.6-sol",
  "gpt-5.6-terra",
]);

export const XAI_RESPONSES_MODELS: ReadonlySet<string> = new Set(["grok-4.6", "grok-4.5"]);

const ANTHROPIC_ACCOUNT_MODELS: ReadonlyArray<readonly [string, number]> = [
  ["claude-fable-5", 1_000_000],
  ["claude-sonnet-5", 1_000_000],
  ["claude-opus-5", 1_000_000],
  ["claude-opus-4-8", 1_000_000],
  ["claude-opus-4-7", 1_000_000],
  ["claude-opus-4-6", 1_000_000],
  ["claude-sonnet-4-6", 1_000_000],
  ["claude-haiku-4-5", 200_000],
];

const XAI_ACCOUNT_MODELS: ReadonlyArray<readonly [string, number | undefined]> = [
  ["grok-4.6", 500_000],
  ["grok-4.5", 500_000],
  ["grok-4.3", 1_000_000],
  ["grok-4.20-0309-reasoning", 1_000_000],
  ["grok-4.20-0309-non-reasoning", 1_000_000],
  ["grok-build-0.1", 256_000],
  ["grok-composer-2.5-fast", undefined],
];

const KIMI_ACCOUNT_MODELS: ReadonlyArray<readonly [string, number]> = [
  ["k3", 262_144],
  ["k3[1m]", 1_048_576],
  ["kimi-k2.7-code", 262_144],
  ["kimi-k2.7-code-highspeed", 262_144],
  ["kimi-k2.6", 262_144],
  ["kimi-k2.5", 262_144],
  ["kimi-for-coding", 262_144],
];

const NOUS_ACCOUNT_MODELS = [
  "tencent/hy3:free",
  "poolside/laguna-s-2.1:free",
  "stepfun/step-3.7-flash:free",
  "poolside/laguna-xs-2.1:free",
] as const;

const GITHUB_COPILOT_ACCOUNT_MODELS = [
  "gpt-4o",
  "gpt-4.1",
  "gpt-4.1-mini",
  "claude-sonnet-4",
  "gemini-2.5-pro",
  "gpt-5-mini",
  "gpt-5.3-codex",
  "gpt-5.4",
  "gpt-5.4-mini",
  "gpt-5.5",
  "gpt-5.6-luna",
  "gpt-5.6-sol",
  "gpt-5.6-terra",
] as const;

function fallbackModel(
  providerId: ProviderId,
  id: string,
  options: { contextWindow?: number; reasoning?: boolean; vision?: boolean } = {},
): ModelInfo {
  return {
    id,
    providerId,
    name: id,
    ...(options.contextWindow ? { contextWindow: options.contextWindow } : {}),
    capabilities: {
      streaming: true,
      tools: true,
      reasoning: options.reasoning ?? false,
      vision: options.vision ?? false,
    },
  };
}

/** Stable account catalogs used when a provider has no model endpoint or live discovery is unavailable. */
export function accountProviderFallbackModels(config: ProviderConfig): ModelInfo[] {
  const provider = config.auth.oauthProvider;
  if (config.kind === "anthropic" && config.auth.method === "oauth") {
    return ANTHROPIC_ACCOUNT_MODELS.map(([id, contextWindow]) => fallbackModel(config.id, id, { contextWindow, reasoning: true, vision: true }));
  }
  if (provider === "xai") {
    return XAI_ACCOUNT_MODELS.map(([id, contextWindow]) => fallbackModel(config.id, id, {
      ...(contextWindow ? { contextWindow } : {}),
      reasoning: !/non-reasoning|build|composer/iu.test(id),
      vision: !/build|composer/iu.test(id),
    }));
  }
  if (provider === "kimi") {
    return KIMI_ACCOUNT_MODELS.map(([id, contextWindow]) => fallbackModel(config.id, id, {
      contextWindow,
      reasoning: id === "k3" || id === "k3[1m]",
      vision: id === "k3" || id === "k3[1m]",
    }));
  }
  if (provider === "nous") {
    return NOUS_ACCOUNT_MODELS.map(id => fallbackModel(config.id, id, { reasoning: true }));
  }
  if (provider === "github-copilot") {
    return GITHUB_COPILOT_ACCOUNT_MODELS.map(id => fallbackModel(config.id, id, {
      reasoning: /^(?:gpt-5|claude|gemini)/iu.test(id),
      vision: true,
    }));
  }
  return [];
}

function openAiCompatibleUsesResponses(config: ProviderConfig, model: string): boolean {
  if (config.auth.oauthProvider === "github-copilot") return GITHUB_COPILOT_RESPONSES_MODELS.has(model);
  return config.auth.method === "oauth"
    && config.auth.oauthProvider === "xai"
    && XAI_RESPONSES_MODELS.has(model);
}

function kimiWireModelId(config: ProviderConfig, model: string): string {
  return config.auth.oauthProvider === "kimi" && model.endsWith("[1m]")
    ? model.slice(0, -"[1m]".length)
    : model;
}

function openAiCompatibleReasoningEffort(config: ProviderConfig, model: string, effort?: string): string | undefined {
  if (config.auth.oauthProvider !== "kimi") return effort;
  if (model !== "k3" && model !== "k3[1m]") return undefined;
  const normalized = effort?.trim().toLowerCase();
  if (!normalized) return "max";
  if (normalized === "none" || normalized === "low") return normalized;
  if (normalized === "medium" || normalized === "high") return "high";
  return "max";
}

function kimiPromptCacheKey(messages: ChatMessage[]): string {
  const firstUser = messages.find(message => message.role === "user")?.content ?? "openwordcode";
  return `owc-${createHash("sha256").update(firstUser, "utf8").digest("hex").slice(0, 32)}`;
}

interface ResponsesStreamOptions {
  config: ProviderConfig;
  fetchImpl: typeof fetch;
  url: string;
  headers: Record<string, string>;
  request: ProviderChatRequest;
}

/** Shared Responses transport used by ChatGPT/Codex accounts and newer Copilot models. */
async function* streamResponses(options: ResponsesStreamOptions): AsyncGenerator<ProviderStreamEvent> {
  const { config, fetchImpl, url, headers, request } = options;
  const instructions = request.messages.filter(message => message.role === "system").map(message => message.content).join("\n\n");
  const body: JsonObject = {
    model: request.model,
    input: responsesInput(request.messages, request.attachments ?? []),
    stream: true,
  };
  if (instructions) body.instructions = instructions;
  if (request.tools?.length) {
    body.tools = responseToolDefinitions(request.tools);
    body.tool_choice = "auto";
  }
  if (request.effort) body.reasoning = { effort: request.effort };
  const response = await fetchImpl(url, {
    method: "POST",
    headers: { "Content-Type": "application/json", Accept: "text/event-stream", ...headers },
    body: JSON.stringify(body),
    signal: request.signal,
  });
  await requireResponse(config, response);
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
  for (const [key, call] of toolCalls) {
    if (call.name && !emittedCalls.has(key)) yield { type: "tool_call", call: { ...call, arguments: call.arguments || "{}" } };
  }
  yield { type: "done" };
}

/** Parse only models the signed-in ChatGPT account explicitly exposes to API clients. */
export function parseChatGptAccountModels(payload: unknown, providerId: ProviderId): ModelInfo[] {
  if (!isObject(payload) || !Array.isArray(payload.models)) return [];
  const seen = new Set<string>();
  const models: ModelInfo[] = [];
  for (const row of payload.models) {
    if (!isObject(row)) continue;
    const id = stringValue(row.slug);
    if (!id || id.length > 200 || !/^[A-Za-z0-9][A-Za-z0-9._:/-]*$/u.test(id)) continue;
    if (row.supported_in_api !== true || row.visibility === "hide" || seen.has(id)) continue;
    const contextWindow = [row.context_window, row.context_length, row.max_context_tokens]
      .find(value => typeof value === "number" && Number.isFinite(value) && value > 0) as number | undefined;
    models.push({
      id,
      providerId,
      name: stringValue(row.display_name) ?? stringValue(row.name) ?? id,
      ...(contextWindow ? { contextWindow: Math.floor(contextWindow) } : {}),
      capabilities: {
        streaming: true,
        tools: row.supports_tools !== false,
        reasoning: row.supports_reasoning !== false && /^(?:gpt-5|o\d)/iu.test(id),
        vision: row.supports_vision !== false,
      },
    });
    seen.add(id);
  }
  return models;
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
    const headers = { Accept: "application/json", ...authHeaders(this.config, credential, this.env) };
    const fallback = accountProviderFallbackModels(this.config);
    // The Kimi Code subscription exposes a documented static catalog, not a reliable
    // OpenAI-compatible /models endpoint. Keep account aliases such as k3[1m] local.
    if (this.config.auth.method === "oauth" && this.config.auth.oauthProvider === "kimi") return fallback;
    try {
      const response = await this.fetchImpl(`${baseUrl}/models`, { headers, signal });
      await requireResponse(this.config, response);
      const payload = await parseJson(response);
      const rows = isObject(payload) ? arrayValue(payload.data) : [];
      const discovered = rows.map(row => modelFromRow(this.config.id, row)).filter((model): model is ModelInfo => model !== null);
      return discovered.length ? discovered : fallback;
    } catch (error) {
      if (fallback.length) return fallback;
      throw error;
    }
  }

  async *streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent> {
    const credential = await readCredential(this.config, this.store, this.env, this.oauth);
    const baseUrl = normalizeBaseUrl(credential?.apiBaseUrl ?? this.config.baseUrl);
    if (openAiCompatibleUsesResponses(this.config, request.model)) {
      yield* streamResponses({
        config: this.config,
        fetchImpl: this.fetchImpl,
        url: `${baseUrl}/responses`,
        headers: authHeaders(this.config, credential, this.env),
        request,
      });
      return;
    }
    const lastUser = lastUserMessage(request.messages);
    const wireModel = kimiWireModelId(this.config, request.model);
    const body: JsonObject = {
      model: wireModel,
      messages: request.messages.map(message => ({
        role: message.role,
        content: openAiMessageContent(message, request.attachments ?? [], message === lastUser),
        ...(message.toolCallId ? { tool_call_id: message.toolCallId } : {}),
        ...(message.toolCalls ? {
          tool_calls: message.toolCalls.map(call => ({
            id: call.id,
            type: "function",
            function: { name: call.name, arguments: call.arguments },
            ...(call.thoughtSignature ? { thought_signature: call.thoughtSignature } : {}),
          })),
        } : {}),
      })),
      stream: true,
    };
    if (request.tools?.length) {
      body.tools = request.tools;
      body.tool_choice = "auto";
    }
    const reasoningEffort = openAiCompatibleReasoningEffort(this.config, request.model, request.effort);
    if (reasoningEffort) body.reasoning_effort = reasoningEffort;
    if (this.config.auth.method === "oauth" && this.config.auth.oauthProvider === "kimi") {
      body.prompt_cache_key = kimiPromptCacheKey(request.messages);
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
          ...(stringValue(rawCall.thought_signature) ? { thoughtSignature: stringValue(rawCall.thought_signature) } : {}),
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
        if (stringValue(rawCall.thought_signature)) existing.thoughtSignature = stringValue(rawCall.thought_signature);
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
    return [{
      id: stringValue(item.id) ?? `tool-${Math.random().toString(16).slice(2)}`,
      name,
      arguments: stringValue(fn.arguments) ?? "{}",
      ...(stringValue(item.thought_signature) ? { thoughtSignature: stringValue(item.thought_signature) } : {}),
    }];
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
    if (this.oauthMode()) return accountProviderFallbackModels(this.config);
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

export const GEMINI_DIRECT_WIRE_RENAMES: Readonly<Record<string, string>> = {
  "gemini-3.7-flash": "gemini-3.7-flash-tiered",
  "gemini-3.6-flash": "gemini-3.6-flash-tiered",
};

export function resolveDirectGeminiWireModelId(modelId: string): string {
  return Object.hasOwn(GEMINI_DIRECT_WIRE_RENAMES, modelId)
    ? GEMINI_DIRECT_WIRE_RENAMES[modelId]!
    : modelId;
}

function googleThinkingLevel(effort: string | undefined): string | undefined {
  if (!effort) return undefined;
  const normalized = effort.trim().toLowerCase();
  if (normalized === "max" || normalized === "xhigh" || normalized === "ultra") return "high";
  return normalized === "low" || normalized === "medium" || normalized === "high" ? normalized : undefined;
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
    const response = await fetchGoogleWithRetry(this.fetchImpl, `${normalizeBaseUrl(this.config.baseUrl)}/models?key=${encodeURIComponent(key)}`, { signal, headers: { Accept: "application/json" } });
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
    const generationConfig: JsonObject = { temperature: 0.2 };
    const thinkingLevel = googleThinkingLevel(request.effort);
    if (thinkingLevel) generationConfig.thinkingConfig = { thinkingLevel };
    const body: JsonObject = { contents, generationConfig };
    if (system) body.systemInstruction = { parts: [{ text: system }] };
    if (request.tools?.length) {
      body.tools = [{ functionDeclarations: request.tools.map(tool => ({ name: tool.function.name, description: tool.function.description, parameters: tool.function.parameters })) }];
      body.toolConfig = { functionCallingConfig: { mode: "AUTO" } };
    }
    const wireModelId = resolveDirectGeminiWireModelId(request.model);
    const response = await fetchGoogleWithRetry(this.fetchImpl, `${normalizeBaseUrl(this.config.baseUrl)}/models/${encodeURIComponent(wireModelId)}:streamGenerateContent?alt=sse&key=${encodeURIComponent(key)}`, {
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
          const fallbackId = `gemini-${name}-${createHash("sha256").update(JSON.stringify(part.functionCall.args ?? {})).digest("hex").slice(0, 12)}`;
          const id = googleToolCallId(stringValue(part.functionCall.id) ?? fallbackId);
          const existing = toolCalls.get(id) ?? { id, name, arguments: "{}" };
          const previous = jsonValue(existing.arguments);
          const previousObject = isObject(previous) ? previous : {};
          const nextArguments = isObject(part.functionCall.args) ? { ...previousObject, ...part.functionCall.args } : previousObject;
          existing.arguments = JSON.stringify(nextArguments);
          const thoughtSignature = stringValue(part.thoughtSignature) ?? stringValue(part.thought_signature);
          if (isLikelyRealThoughtSignature(thoughtSignature)) existing.thoughtSignature = thoughtSignature;
          toolCalls.set(id, existing);
        }
      }
    }
    for (const call of toolCalls.values()) yield { type: "tool_call", call };
    yield { type: "done" };
  }
}

export const GEMINI_FLASH_WIRE_ID = "gemini-3.7-flash-tiered";
export const ANTIGRAVITY_IDE_VERSION = "2.5.5";

/** Match the first-party Antigravity IDE fingerprint required by newer CCA models. */
export function antigravityUserAgent(env: NodeJS.ProcessEnv = process.env): string {
  return env.GOOGLE_ANTIGRAVITY_USER_AGENT?.trim()
    || `antigravity/ide/${ANTIGRAVITY_IDE_VERSION} (os_type=windows; arch=amd64; aidev_client; auth_method=oauth)`;
}

const RETIRED_FLASH_TIERS: Readonly<Record<string, string>> = {
  "gemini-3.6-flash": "medium",
  "gemini-3.6-flash-low": "low",
  "gemini-3.6-flash-medium": "medium",
  "gemini-3.6-flash-high": "high",
  "gemini-3.5-flash-extra-low": "low",
  "gemini-3.5-flash-low": "medium",
  "gemini-3.5-flash-mid": "medium",
  "gemini-3.5-flash-high": "high",
  "gemini-3-flash-agent": "high",
};

export const ANTIGRAVITY_EFFORT_WIRE_MAP: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  "gemini-3.1-pro": {
    low: "gemini-3.1-pro-low",
    high: "gemini-pro-agent",
  },
};

export const ANTIGRAVITY_DEFAULT_EFFORT: Readonly<Record<string, string>> = {
  "gemini-3.1-pro": "high",
};

export const ANTIGRAVITY_THINKING_LEVEL_MODELS: Readonly<Record<string, string>> = {
  "gemini-3.7-flash": "medium",
};

export const ANTIGRAVITY_MODEL_ALIASES: Readonly<Record<string, string>> = {
  "gemini-3.1-pro-high": "gemini-pro-agent",
  "gemini-3.1-pro-preview": "gemini-pro-agent",
  "gemini-3.1-pro-low": "gemini-3.1-pro-low",
  "gemini-pro-agent": "gemini-pro-agent",
  ...Object.fromEntries(Object.keys(RETIRED_FLASH_TIERS).map(id => [id, GEMINI_FLASH_WIRE_ID])),
};

export const ANTIGRAVITY_MODELS = [
  "gemini-3.7-flash",
  "gemini-3.1-pro",
  "gemini-3.1-flash-image",
  "claude-sonnet-4-6",
  "claude-opus-4-6-thinking",
  "gpt-oss-120b-medium",
] as const;

const ANTIGRAVITY_CONTEXT_WINDOWS: Readonly<Record<string, number>> = {
  "gemini-3.7-flash": 1_048_576,
  "gemini-3.1-pro": 1_048_576,
  "gemini-3.1-flash-image": 1_048_576,
  "claude-sonnet-4-6": 250_000,
  "claude-opus-4-6-thinking": 250_000,
  "gpt-oss-120b-medium": 131_072,
};

type AntigravityEffortWireModels = Partial<Record<"low" | "medium" | "high", string>>;

interface AntigravityDiscoveredModel {
  model: ModelInfo;
  wireModelId: string;
  effortWireModelIds?: AntigravityEffortWireModels;
}

interface AntigravityDiscoveredMapping {
  models: ReadonlyMap<string, string>;
  effortModels: ReadonlyMap<string, AntigravityEffortWireModels>;
}

const antigravityDiscoveredMappings = new Map<string, AntigravityDiscoveredMapping>();

function antigravityBaseUrlKey(baseUrl: string | undefined): string | undefined {
  if (!baseUrl?.trim()) return undefined;
  try { return normalizeBaseUrl(baseUrl).toLowerCase(); } catch { return undefined; }
}

function registerAntigravityDiscoveredModels(baseUrl: string, models: readonly AntigravityDiscoveredModel[]): void {
  const key = antigravityBaseUrlKey(baseUrl);
  if (!key) return;
  const wireModels = new Map<string, string>();
  const effortModels = new Map<string, AntigravityEffortWireModels>();
  for (const model of models) {
    wireModels.set(model.model.id, model.wireModelId);
    if (model.effortWireModelIds) effortModels.set(model.model.id, { ...model.effortWireModelIds });
  }
  antigravityDiscoveredMappings.set(key, { models: wireModels, effortModels });
}

function antigravityDiscoveredMapping(baseUrl: string | undefined): AntigravityDiscoveredMapping | undefined {
  const key = antigravityBaseUrlKey(baseUrl);
  return key ? antigravityDiscoveredMappings.get(key) : undefined;
}

function discoveredAntigravityEffortWireModelId(modelId: string, effort: string | undefined, baseUrl: string | undefined): string | undefined {
  const effortMap = antigravityDiscoveredMapping(baseUrl)?.effortModels.get(modelId);
  if (!effortMap) return undefined;
  const requested = effort ? resolveAntigravityThinkingLevel(effort) : undefined;
  if (requested && effortMap[requested as keyof AntigravityEffortWireModels]) return effortMap[requested as keyof AntigravityEffortWireModels];
  const defaultEffort = ANTIGRAVITY_DEFAULT_EFFORT[modelId] ?? ANTIGRAVITY_THINKING_LEVEL_MODELS[modelId];
  if (defaultEffort && effortMap[defaultEffort as keyof AntigravityEffortWireModels]) return effortMap[defaultEffort as keyof AntigravityEffortWireModels];
  return Object.values(effortMap)[0];
}

export function resolveAntigravityThinkingLevel(effort: string): string | undefined {
  return googleThinkingLevel(effort);
}

export function retiredAntigravityFlashTier(modelId: string): string | undefined {
  return Object.hasOwn(RETIRED_FLASH_TIERS, modelId) ? RETIRED_FLASH_TIERS[modelId] : undefined;
}

export function resolveAntigravityWireModelId(modelId: string, baseUrl?: string): string {
  const discovered = antigravityDiscoveredMapping(baseUrl)?.models.get(modelId);
  if (discovered) return discovered;
  if (modelId === "gemini-3.7-flash") return GEMINI_FLASH_WIRE_ID;
  return Object.hasOwn(ANTIGRAVITY_MODEL_ALIASES, modelId)
    ? ANTIGRAVITY_MODEL_ALIASES[modelId]!
    : modelId;
}

export function isAntigravitySuffixModelId(modelId: string): boolean {
  return !(ANTIGRAVITY_MODELS as readonly string[]).includes(modelId);
}

export function resolveAntigravityEffortWireModel(
  modelId: string,
  effort?: string,
  baseUrl?: string,
): { wireModelId: string; thinkingLevel?: string } {
  const discoveredEffortWireModelId = discoveredAntigravityEffortWireModelId(modelId, effort, baseUrl);
  if (discoveredEffortWireModelId) return { wireModelId: discoveredEffortWireModelId };

  const retiredTier = retiredAntigravityFlashTier(modelId);
  if (retiredTier) {
    return {
      wireModelId: GEMINI_FLASH_WIRE_ID,
      thinkingLevel: effort ? resolveAntigravityThinkingLevel(effort) ?? retiredTier : retiredTier,
    };
  }

  if (isAntigravitySuffixModelId(modelId)) {
    return { wireModelId: resolveAntigravityWireModelId(modelId, baseUrl) };
  }

  const defaultThinkingLevel = ANTIGRAVITY_THINKING_LEVEL_MODELS[modelId];
  if (defaultThinkingLevel) {
    return {
      wireModelId: resolveAntigravityWireModelId(modelId, baseUrl),
      thinkingLevel: effort ? resolveAntigravityThinkingLevel(effort) ?? defaultThinkingLevel : defaultThinkingLevel,
    };
  }

  const effortMap = ANTIGRAVITY_EFFORT_WIRE_MAP[modelId];
  if (effortMap) {
    const normalizedEffort = effort ? resolveAntigravityThinkingLevel(effort) : undefined;
    if (normalizedEffort && effortMap[normalizedEffort]) {
      return { wireModelId: effortMap[normalizedEffort]!, thinkingLevel: normalizedEffort };
    }
    const defaultEffort = ANTIGRAVITY_DEFAULT_EFFORT[modelId]!;
    return { wireModelId: effortMap[defaultEffort]! };
  }

  if (/^claude-/i.test(modelId) && effort) {
    const thinkingLevel = resolveAntigravityThinkingLevel(effort);
    return { wireModelId: modelId, ...(thinkingLevel ? { thinkingLevel } : {}) };
  }

  return { wireModelId: resolveAntigravityWireModelId(modelId, baseUrl) };
}

interface AntigravityPart extends JsonObject {
  thought?: boolean;
  thoughtSignature?: string;
  thought_signature?: string;
}

/** Remove signature shapes that Claude-on-Antigravity rejects. */
export function sanitizeAntigravityClaudeSignatures(contents: unknown[]): unknown[] {
  for (const rawContent of contents) {
    if (!isObject(rawContent) || !Array.isArray(rawContent.parts)) continue;
    const parts = rawContent.parts.filter(isObject) as AntigravityPart[];
    if (rawContent.role !== "model") {
      for (const part of parts) {
        delete part.thoughtSignature;
        delete part.thought_signature;
      }
      rawContent.parts = parts;
      continue;
    }
    rawContent.parts = parts.filter(part => {
      if (part.thought !== true) return true;
      return (typeof part.thoughtSignature === "string" && part.thoughtSignature.length > 0)
        || (typeof part.thought_signature === "string" && part.thought_signature.length > 0);
    });
  }
  return contents;
}

function antigravitySessionId(messages: ChatMessage[]): string {
  const firstUserText = messages.find(message => message.role === "user")?.content;
  if (!firstUserText) return `-${Math.floor(Math.random() * 9e18).toString()}`;
  const digest = createHash("sha256").update(firstUserText, "utf8").digest();
  return `-${(digest.readBigUInt64BE(0) & 0x7fffffffffffffffn).toString()}`;
}

function validDiscoveredModelId(value: unknown): value is string {
  return typeof value === "string"
    && value.length > 0
    && value.length <= 200
    && /^[A-Za-z0-9][A-Za-z0-9._:/-]*$/.test(value);
}

function antigravityDisplayModelId(displayName: unknown, wireId: string): string | undefined {
  if (typeof displayName !== "string") return undefined;
  const label = displayName.trim();
  if (!label || label.length > 512) return undefined;
  const id = label.normalize("NFKC")
    .toLowerCase()
    .replace(/\s+/g, "-")
    .replace(/[^a-z0-9._:/-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
  if (!validDiscoveredModelId(id)) return undefined;
  if (id === wireId || id === `${wireId}-thinking`) return wireId;
  return id;
}

function antigravityPickerModelId(
  wireId: string,
  info: JsonObject,
  available: ReadonlyMap<string, JsonObject>,
): string {
  if ((wireId === "gemini-3.1-pro-low" || wireId === "gemini-pro-agent")
    && available.has("gemini-3.1-pro-low")
    && available.has("gemini-pro-agent")) return "gemini-3.1-pro";
  if (wireId.endsWith("-tiered")) {
    const baseId = wireId.slice(0, -"-tiered".length);
    if (validDiscoveredModelId(baseId)) return baseId;
  }
  const effortMatch = /^(.*)-(low|medium|high)$/.exec(wireId);
  if (effortMatch) {
    const baseId = effortMatch[1]!;
    if (["low", "medium", "high"].every(level => available.has(`${baseId}-${level}`))) return baseId;
  }
  const displayId = antigravityDisplayModelId(info.displayName, wireId);
  if (displayId) {
    const displayEffort = /^(.*)-(low|medium|high)$/.exec(displayId);
    if (displayEffort && (ANTIGRAVITY_MODELS as readonly string[]).includes(displayEffort[1]!)) return displayEffort[1]!;
    return displayId;
  }
  return wireId;
}

function completeAntigravityEffortWireModels(
  pickerId: string,
  available: ReadonlyMap<string, JsonObject>,
): AntigravityEffortWireModels | undefined {
  if (pickerId === "gemini-3.1-pro"
    && available.has("gemini-3.1-pro-low")
    && available.has("gemini-pro-agent")) {
    return { low: "gemini-3.1-pro-low", high: "gemini-pro-agent" };
  }
  const suffixes: AntigravityEffortWireModels = {};
  for (const level of ["low", "medium", "high"] as const) {
    const wireId = `${pickerId}-${level}`;
    if (!available.has(wireId)) return undefined;
    suffixes[level] = wireId;
  }
  return suffixes;
}

function staticAntigravityModels(providerId: ProviderId): ModelInfo[] {
  return ANTIGRAVITY_MODELS.map(id => ({
    id,
    providerId,
    name: id,
    contextWindow: ANTIGRAVITY_CONTEXT_WINDOWS[id],
    capabilities: {
      streaming: true,
      tools: true,
      reasoning: /gemini|opus|sonnet/iu.test(id),
      vision: /image|gemini|opus|sonnet/iu.test(id),
    },
  }));
}

function parseAntigravityAvailableModelsDetailed(payload: unknown, providerId: ProviderId): AntigravityDiscoveredModel[] {
  if (!isObject(payload) || !isObject(payload.models) || !Array.isArray(payload.agentModelSorts)) return [];
  const wireIds: string[] = [];
  for (const rawSort of payload.agentModelSorts) {
    if (!isObject(rawSort) || !Array.isArray(rawSort.groups)) return [];
    for (const rawGroup of rawSort.groups) {
      if (!isObject(rawGroup) || !Array.isArray(rawGroup.modelIds)) return [];
      for (const id of rawGroup.modelIds) {
        if (!validDiscoveredModelId(id) || !Object.hasOwn(payload.models, id) || wireIds.length >= 200) return [];
        wireIds.push(id);
      }
    }
  }

  if (Array.isArray(payload.imageGenerationModelIds)
    && payload.imageGenerationModelIds.includes("gemini-3.1-flash-image")
    && Object.hasOwn(payload.models, "gemini-3.1-flash-image")) {
    wireIds.push("gemini-3.1-flash-image");
  }
  if (isObject(payload.tieredModelIds) && Array.isArray(payload.tieredModelIds.flash)) {
    for (const id of payload.tieredModelIds.flash) {
      if (validDiscoveredModelId(id) && Object.hasOwn(payload.models, id)) wireIds.push(id);
    }
  }

  const available = new Map<string, JsonObject>();
  for (const wireId of wireIds) {
    if (available.has(wireId)) continue;
    const alias = ANTIGRAVITY_MODEL_ALIASES[wireId];
    if (alias && alias !== wireId) continue;
    const info = isObject(payload.models[wireId]) ? payload.models[wireId] : undefined;
    if (info) available.set(wireId, info);
  }
  const seen = new Set<string>();
  const models: AntigravityDiscoveredModel[] = [];
  for (const [wireId, info] of available) {
    const id = antigravityPickerModelId(wireId, info, available);
    if (seen.has(id)) continue;
    const maxTokens = typeof info.maxTokens === "number" && Number.isFinite(info.maxTokens) && info.maxTokens > 0
      ? Math.floor(info.maxTokens)
      : ANTIGRAVITY_CONTEXT_WINDOWS[id];
    models.push({
      model: {
        id,
        providerId,
        name: id === wireId ? stringValue(info.displayName) ?? id : id,
        ...(maxTokens ? { contextWindow: maxTokens } : {}),
        capabilities: {
          streaming: true,
          tools: true,
          reasoning: info.supportsThinking === true || /gemini|opus|sonnet/iu.test(id),
          vision: info.supportsImages === true || /image|gemini|opus|sonnet/iu.test(id),
        },
      },
      wireModelId: wireId,
      ...(completeAntigravityEffortWireModels(id, available)
        ? { effortWireModelIds: completeAntigravityEffortWireModels(id, available) }
        : {}),
    });
    seen.add(id);
  }
  return models;
}

export function parseAntigravityAvailableModels(payload: unknown, providerId: ProviderId): ModelInfo[] {
  return parseAntigravityAvailableModelsDetailed(payload, providerId).map(model => model.model);
}

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

  private baseUrl(): string { return normalizeBaseUrl(this.config.baseUrl); }

  private async discoverModels(credential: ResolvedCredential, signal?: AbortSignal): Promise<ModelInfo[]> {
    const response = await fetchGoogleWithRetry(this.fetchImpl, `${this.baseUrl()}/v1internal:fetchAvailableModels`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: `Bearer ${credential.token}`,
        "User-Agent": antigravityUserAgent(this.env),
      },
      body: JSON.stringify({ project: credential.projectId }),
      signal,
    });
    await requireResponse(this.config, response);
    const discovered = parseAntigravityAvailableModelsDetailed(await parseJson(response), this.config.id);
    if (!discovered.length) throw new ProviderError("invalid_provider_response", "Google Antigravity returned no usable agent models");
    registerAntigravityDiscoveredModels(this.baseUrl(), discovered);
    return discovered.map(model => model.model);
  }

  private async sendRequest(
    credential: ResolvedCredential,
    request: ProviderChatRequest,
    resolved: { wireModelId: string; thinkingLevel?: string },
  ): Promise<Response> {
    const system = request.messages.filter(message => message.role === "system").map(message => message.content).join("\n\n");
    const contents = geminiContents(request.messages, request.attachments ?? []);
    const isClaude = /^claude-/i.test(resolved.wireModelId);
    if (isClaude) {
      sanitizeAntigravityClaudeSignatures(contents);
      const last = contents.at(-1);
      if (!isObject(last) || last.role === "model") contents.push({ role: "user", parts: [{ text: "(continue)" }] });
    }
    const requestBody: JsonObject = { contents, sessionId: antigravitySessionId(request.messages) };
    if (system) requestBody.systemInstruction = { parts: [{ text: system }] };
    if (resolved.thinkingLevel) requestBody.generationConfig = { thinkingConfig: { thinkingLevel: resolved.thinkingLevel } };
    if (request.tools?.length) {
      requestBody.tools = [{ functionDeclarations: request.tools.map(tool => ({ name: tool.function.name, description: tool.function.description, parameters: tool.function.parameters })) }];
      requestBody.toolConfig = { functionCallingConfig: { mode: isClaude ? "VALIDATED" : "AUTO" } };
    }
    const envelope = {
      model: resolved.wireModelId,
      userAgent: "antigravity",
      requestType: "agent",
      project: credential.projectId,
      requestId: `agent-${randomUUID()}`,
      request: requestBody,
    };
    return await fetchGoogleWithRetry(this.fetchImpl, `${this.baseUrl()}/v1internal:streamGenerateContent?alt=sse`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "text/event-stream",
        Authorization: `Bearer ${credential.token}`,
        "User-Agent": antigravityUserAgent(this.env),
      },
      body: JSON.stringify(envelope),
      signal: request.signal,
    });
  }

  async listModels(signal?: AbortSignal): Promise<ModelInfo[]> {
    const fallback = staticAntigravityModels(this.config.id);
    try {
      const credential = await this.credential();
      return await this.discoverModels(credential, signal);
    } catch (error) {
      if (signal?.aborted) throw error;
      return fallback;
    }
  }

  async *streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent> {
    const credential = await this.credential();
    if (!antigravityDiscoveredMapping(this.baseUrl())) {
      try { await this.discoverModels(credential, request.signal); } catch (error) {
        if (request.signal?.aborted) throw error;
      }
    }
    let resolved = resolveAntigravityEffortWireModel(request.model, request.effort, this.baseUrl());
    let response = await this.sendRequest(credential, request, resolved);
    if (response.status === 400 || response.status === 404) {
      try { await response.body?.cancel(); } catch { /* best-effort connection cleanup */ }
      const previous = `${resolved.wireModelId}:${resolved.thinkingLevel ?? ""}`;
      try { await this.discoverModels(credential, request.signal); } catch (error) {
        if (request.signal?.aborted) throw error;
      }
      const refreshed = resolveAntigravityEffortWireModel(request.model, request.effort, this.baseUrl());
      if (`${refreshed.wireModelId}:${refreshed.thinkingLevel ?? ""}` !== previous) {
        resolved = refreshed;
        response = await this.sendRequest(credential, request, resolved);
      }
    }
    await requireResponse(this.config, response);
    const toolCalls = new Map<string, ToolCall>();
    for await (const raw of sseJson(response)) {
      if (!isObject(raw)) continue;
      if (isObject(raw.error)) throw new ProviderError("provider_response_error", stringValue(raw.error.message) ?? "Google Antigravity returned an error");
      const root = isObject(raw.response) ? raw.response : raw;
      if (isObject(root.error)) throw new ProviderError("provider_response_error", stringValue(root.error.message) ?? "Google Antigravity returned an error");
      for (const candidate of arrayValue(root.candidates)) {
        if (!isObject(candidate) || !isObject(candidate.content)) continue;
        for (const part of arrayValue(candidate.content.parts)) {
          if (!isObject(part)) continue;
          if (typeof part.text === "string" && part.thought !== true) yield { type: "text", delta: part.text };
          if (isObject(part.functionCall)) {
            const name = stringValue(part.functionCall.name);
            if (!name) continue;
            const id = googleToolCallId(stringValue(part.functionCall.id) ?? `antigravity-${toolCalls.size}`);
            const existing = toolCalls.get(id) ?? { id, name, arguments: "{}" };
            existing.name = name;
            existing.arguments = JSON.stringify(isObject(part.functionCall.args) ? part.functionCall.args : {});
            const thoughtSignature = stringValue(part.thoughtSignature) ?? stringValue(part.thought_signature);
            if (isLikelyRealThoughtSignature(thoughtSignature)) existing.thoughtSignature = thoughtSignature;
            toolCalls.set(id, existing);
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

  async listModels(signal?: AbortSignal): Promise<ModelInfo[]> {
    const fallback: ModelInfo[] = [
      { id: "gpt-5.6-sol", providerId: this.config.id, name: "GPT-5.6 Sol", capabilities: { streaming: true, tools: true, reasoning: true, vision: true } },
      { id: "gpt-5.6-terra", providerId: this.config.id, name: "GPT-5.6 Terra", capabilities: { streaming: true, tools: true, reasoning: true, vision: true } },
      { id: "gpt-5.6-luna", providerId: this.config.id, name: "GPT-5.6 Luna", capabilities: { streaming: true, tools: true, reasoning: true, vision: true } },
    ];
    try {
      const credential = await this.resolveCredential(this.config);
      if (!credential) return fallback;
      const headers: Record<string, string> = { Authorization: `Bearer ${credential.accessToken}`, Accept: "application/json" };
      if (credential.accountId) headers["ChatGPT-Account-Id"] = credential.accountId;
      const response = await this.fetchImpl(`${normalizeBaseUrl(credential.apiBaseUrl ?? this.config.baseUrl)}/models?client_version=0.0.0`, {
        headers,
        redirect: "error",
        signal,
      });
      if (!response.ok) return fallback;
      const discovered = parseChatGptAccountModels(await parseJson(response), this.config.id);
      return discovered.length ? discovered : fallback;
    } catch {
      return fallback;
    }
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
    if (request.effort) body.reasoning = { effort: request.effort };
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      Accept: "text/event-stream",
      Authorization: `Bearer ${credential.accessToken}`,
    };
    if (credential.accountId) headers["ChatGPT-Account-Id"] = credential.accountId;
    const response = await this.fetchImpl(`${normalizeBaseUrl(credential.apiBaseUrl ?? this.config.baseUrl)}/responses`, {
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

function configuredBridgeBaseUrl(env: NodeJS.ProcessEnv): string {
  const configured = Number(env.OPENWORDCODE_BRIDGE_PORT ?? 10_101);
  const port = Number.isInteger(configured) && configured > 0 && configured <= 65_535 ? configured : 10_101;
  return `http://127.0.0.1:${port}/v1`;
}

export function defaultProviderConfigs(env: NodeJS.ProcessEnv = process.env): ProviderConfig[] {
  return [
    {
      id: "openwordcode-bridge", displayName: "OpenWordCode Bridge", kind: "openwordcode-bridge", baseUrl: configuredBridgeBaseUrl(env), enabled: true, local: true,
      auth: { method: "oauth", credentialRef: "provider:openwordcode-account" }, defaultModel: "gpt-5.6-luna", privacyNote: "Uses the local OpenWordCode Bridge. Its configured provider/account determines where document content is sent.",
    },
    {
      id: "openwordcode-account", displayName: "OpenWordCode account", kind: "openai-codex", baseUrl: "https://chatgpt.com/backend-api/codex", enabled: true, local: false, internal: true,
      auth: { method: "oauth", credentialRef: "provider:openwordcode-account" }, defaultModel: "gpt-5.6-luna", privacyNote: "Internal account transport used by the OpenWordCode Bridge.",
    },
    {
      id: "anthropic", displayName: "Anthropic", kind: "anthropic", baseUrl: "https://api.anthropic.com/v1", enabled: true, local: false,
      auth: { method: "oauth", oauthProvider: "anthropic", oauthCredentialRef: "oauth:anthropic" }, defaultModel: "claude-sonnet-5", privacyNote: "Selected document context is sent to Anthropic through the connected Claude account.",
    },
    {
      id: "google-antigravity", displayName: "Google Antigravity", kind: "google-antigravity", baseUrl: "https://daily-cloudcode-pa.googleapis.com", enabled: true, local: false,
      auth: { method: "oauth", oauthProvider: "google-antigravity", oauthCredentialRef: "oauth:google-antigravity" }, defaultModel: "gemini-3.7-flash", privacyNote: "Selected document context is sent to Google Cloud Code Assist through the signed-in Antigravity account.",
    },
    {
      id: "xai", displayName: "xAI", kind: "openai-compatible", baseUrl: "https://api.x.ai/v1", enabled: true, local: false,
      auth: { method: "oauth", oauthProvider: "xai", oauthCredentialRef: "oauth:xai" }, defaultModel: "grok-4.5", privacyNote: "Selected document context is sent to xAI through the connected account.",
    },
    {
      id: "kimi", displayName: "Kimi Code", kind: "openai-compatible", baseUrl: "https://api.kimi.com/coding/v1", enabled: true, local: false,
      auth: { method: "oauth", oauthProvider: "kimi", oauthCredentialRef: "oauth:kimi" }, defaultModel: "kimi-k2.7-code", privacyNote: "Selected document context is sent to Kimi Code through the connected account.",
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
      id: "demo", displayName: "Demo (offline)", kind: "demo", baseUrl: "demo://offline", enabled: true, local: true, internal: true,
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
