import Fastify, { type FastifyInstance, type FastifyReply, type FastifyRequest } from "fastify";
import { randomUUID } from "node:crypto";
import { z } from "zod";
import {
  type ChatAttachment,
  type ChatMessage,
  type ChatToolDefinition,
  type ModelInfo,
  type ProviderConfig,
  type ToolCall,
} from "../../../packages/shared/src/index.js";
import { createProvider, ProviderError, type ProviderRuntime } from "../../../packages/providers/src/index.js";
import { equalSecret, normalizeBaseUrl, safeErrorMessage } from "../../../packages/security/src/index.js";
import { providerConfig } from "./config.js";
import { type CoreState } from "./server.js";

export const OPENWORDCODE_BRIDGE_DEFAULT_PORT = 10_101;

const bridgeBodySchema = z.object({
  model: z.string().trim().min(1).max(200),
  messages: z.array(z.unknown()).min(1).max(40),
  stream: z.boolean().optional(),
  tools: z.array(z.unknown()).max(64).optional(),
}).passthrough();

type RecordValue = Record<string, unknown>;

export interface OpenWordCodeBridgeStatus {
  available: boolean;
  endpoint: string;
  models: number;
  detail: string;
}

export function bridgePort(env: NodeJS.ProcessEnv = process.env): number {
  const value = Number(env.OPENWORDCODE_BRIDGE_PORT ?? OPENWORDCODE_BRIDGE_DEFAULT_PORT);
  return Number.isInteger(value) && value >= 0 && value <= 65_535 ? value : OPENWORDCODE_BRIDGE_DEFAULT_PORT;
}

export function bridgeRootUrl(env: NodeJS.ProcessEnv = process.env): string {
  return `http://127.0.0.1:${bridgePort(env)}`;
}

export function bridgeApiUrl(env: NodeJS.ProcessEnv = process.env): string {
  return `${bridgeRootUrl(env)}/v1`;
}

function isRecord(value: unknown): value is RecordValue {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function errorStatus(error: unknown): number {
  if (error instanceof BridgeRequestError) return error.statusCode;
  if (error instanceof ProviderError && error.status) return error.status;
  return 502;
}

function errorBody(error: unknown): RecordValue {
  const message = safeErrorMessage(error) || "OpenWordCode Bridge could not complete the request";
  return { error: { message, type: error instanceof ProviderError ? error.code : "bridge_error" } };
}

class BridgeRequestError extends Error {
  constructor(readonly statusCode: number, message: string) { super(message); }
}

function configuredBridgeToken(state: CoreState): string | null {
  return state.env.OPENWORDCODE_BRIDGE_TOKEN?.trim() || null;
}

function authorizeBridgeRequest(state: CoreState, request: FastifyRequest): void {
  const expected = configuredBridgeToken(state);
  if (!expected) return;
  const header = request.headers["x-openwordcode-bridge-token"];
  const authorization = request.headers.authorization;
  const actual = Array.isArray(header)
    ? header[0]
    : header || (authorization?.startsWith("Bearer ") ? authorization.slice(7) : undefined);
  if (!actual || !equalSecret(actual, expected)) throw new BridgeRequestError(401, "OpenWordCode Bridge token required");
}

function bridgeProviderIds(state: CoreState): string[] {
  const configured = state.env.OPENWORDCODE_BRIDGE_PROVIDER_ID?.trim();
  if (configured) {
    const provider = state.settings.providers[configured];
    if (!provider || !provider.enabled || provider.id === "openwordcode-bridge") {
      throw new BridgeRequestError(400, `Unknown OpenWordCode Bridge target provider: ${configured}`);
    }
    return [provider.id];
  }

  const providers = Object.values(state.settings.providers).filter(provider => provider.enabled && provider.id !== "openwordcode-bridge");
  const account = providers.find(provider => provider.id === "openwordcode-account");
  return [
    ...(account ? [account.id] : []),
    ...providers.filter(provider => provider.id !== account?.id).map(provider => provider.id),
  ];
}

function runtimeForBridge(state: CoreState, provider: ProviderConfig): ProviderRuntime {
  return createProvider(provider, {
    store: state.credentials,
    env: state.env,
    oauth: candidate => state.auth.resolveOAuth(candidate),
  });
}

interface BridgeRoute {
  provider: ProviderConfig;
  model: ModelInfo;
}

async function modelCatalog(state: CoreState): Promise<{ models: ModelInfo[]; routes: Map<string, BridgeRoute> }> {
  const models: ModelInfo[] = [];
  const routes = new Map<string, BridgeRoute>();
  for (const id of bridgeProviderIds(state)) {
    const provider = providerConfig(state.settings, id);
    try {
      const discovered = await runtimeForBridge(state, provider).listModels();
      for (const model of discovered) {
        if (routes.has(model.id)) continue;
        const normalized = { ...model, providerId: "openwordcode-bridge" as const };
        models.push(normalized);
        routes.set(model.id, { provider, model });
      }
    } catch {
      // One disconnected provider must not take down the bridge catalog.
    }
  }
  return { models, routes };
}

async function routeForModel(state: CoreState, modelId: string): Promise<BridgeRoute> {
  const catalog = await modelCatalog(state);
  const known = catalog.routes.get(modelId);
  if (known) return known;
  const fallbackId = bridgeProviderIds(state)[0];
  if (!fallbackId) throw new BridgeRequestError(503, "No provider is configured for OpenWordCode Bridge");
  const provider = providerConfig(state.settings, fallbackId);
  return { provider, model: { id: modelId, providerId: "openwordcode-bridge", name: modelId, capabilities: { streaming: true, tools: true, vision: true } } };
}

function mimeFromDataUrl(value: string): string | null {
  const match = /^data:([^;,]+);base64,/u.exec(value);
  const mime = match?.[1]?.toLowerCase();
  return mime && ["application/pdf", "image/gif", "image/jpeg", "image/png", "image/webp"].includes(mime) ? mime : null;
}

function attachmentFromDataUrl(value: string, name: string, id: string, mimeHint?: string): ChatAttachment {
  const mimeType = mimeFromDataUrl(value) ?? mimeHint;
  if (!mimeType || !["application/pdf", "image/gif", "image/jpeg", "image/png", "image/webp"].includes(mimeType)) {
    throw new BridgeRequestError(400, "OpenWordCode Bridge accepts only image and PDF data URLs");
  }
  const comma = value.indexOf(",");
  const encoded = comma >= 0 ? value.slice(comma + 1) : "";
  const size = Math.floor(encoded.length * 3 / 4);
  if (size < 1 || size > 6_000_000) throw new BridgeRequestError(413, "Bridge attachment exceeds the 6 MB file limit");
  return { id, name, mimeType: mimeType as ChatAttachment["mimeType"], size, dataUrl: value };
}

function contentTextAndAttachments(content: unknown, messageIndex: number): { text: string; attachments: ChatAttachment[] } {
  if (typeof content === "string") return { text: content, attachments: [] };
  if (!Array.isArray(content)) return { text: "", attachments: [] };
  const text: string[] = [];
  const attachments: ChatAttachment[] = [];
  for (let partIndex = 0; partIndex < content.length; partIndex += 1) {
    const part = content[partIndex];
    if (typeof part === "string") {
      text.push(part);
      continue;
    }
    if (!isRecord(part)) continue;
    const type = stringValue(part.type);
    if (["text", "input_text", "output_text"].includes(type ?? "") && typeof part.text === "string") {
      text.push(part.text);
      continue;
    }
    if (type === "image_url" || type === "input_image") {
      const image = isRecord(part.image_url) ? stringValue(part.image_url.url) : stringValue(part.image_url) ?? stringValue(part.url);
      if (!image) continue;
      if (!image.startsWith("data:")) throw new BridgeRequestError(400, "OpenWordCode Bridge does not fetch remote image URLs; send a data URL instead");
      attachments.push(attachmentFromDataUrl(image, `image-${messageIndex + 1}-${partIndex + 1}`, `bridge-image-${messageIndex}-${partIndex}`));
      continue;
    }
    if (type === "file" || type === "input_file" || type === "document") {
      const file = isRecord(part.file) ? part.file : part;
      const data = stringValue(file.file_data) ?? stringValue(file.data) ?? stringValue(file.file_url);
      if (!data) continue;
      const filename = stringValue(file.filename) ?? stringValue(file.name) ?? `file-${messageIndex + 1}-${partIndex + 1}`;
      const mimeHint = stringValue(file.mime_type) ?? stringValue(part.mime_type)
        ?? (type === "document" || /\.pdf$/iu.test(filename) ? "application/pdf" : undefined);
      const dataUrl = data.startsWith("data:") ? data : `${mimeHint ? `data:${mimeHint}` : "data:application/pdf"};base64,${data}`;
      attachments.push(attachmentFromDataUrl(dataUrl, filename, `bridge-file-${messageIndex}-${partIndex}`, mimeHint));
    }
  }
  return { text: text.join("\n"), attachments };
}

function toolCallsFromMessage(value: unknown): ToolCall[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const calls = value.flatMap((raw, index) => {
    if (!isRecord(raw)) return [];
    const fn = isRecord(raw.function) ? raw.function : raw;
    const name = stringValue(fn.name);
    if (!name) return [];
    return [{ id: stringValue(raw.id) ?? `bridge-call-${index}`, name, arguments: stringValue(fn.arguments) ?? "{}" }];
  });
  return calls.length ? calls : undefined;
}

function sharedMessages(rawMessages: unknown[]): { messages: ChatMessage[]; attachments: ChatAttachment[] } {
  const messages: ChatMessage[] = [];
  const attachments: ChatAttachment[] = [];
  for (let index = 0; index < rawMessages.length; index += 1) {
    const raw = rawMessages[index];
    if (!isRecord(raw)) throw new BridgeRequestError(400, `messages[${index}] must be an object`);
    const rawRole = stringValue(raw.role);
    const role = rawRole === "developer" ? "system" : rawRole;
    if (role === "system") {
      const parsed = contentTextAndAttachments(raw.content, index);
      messages.push({ role: "system", content: parsed.text });
      continue;
    }
    if (role === "assistant") {
      const parsed = contentTextAndAttachments(raw.content, index);
      messages.push({ role: "assistant", content: parsed.text, toolCalls: toolCallsFromMessage(raw.tool_calls) });
      continue;
    }
    if (role === "tool") {
      const parsed = contentTextAndAttachments(raw.content, index);
      messages.push({ role: "tool", toolCallId: stringValue(raw.tool_call_id), content: parsed.text });
      continue;
    }
    if (role !== "user") throw new BridgeRequestError(400, `Unsupported message role: ${rawRole ?? "missing"}`);
    const parsed = contentTextAndAttachments(raw.content, index);
    messages.push({ role: "user", content: parsed.text });
    attachments.push(...parsed.attachments);
  }
  if (!messages.some(message => message.role === "user")) throw new BridgeRequestError(400, "messages must include a user message");
  if (attachments.length > 4) throw new BridgeRequestError(413, "Bridge accepts at most four attachments per request");
  if (attachments.reduce((total, file) => total + file.size, 0) > 12_000_000) throw new BridgeRequestError(413, "Bridge attachments exceed the 12 MB total limit");
  return { messages, attachments };
}

function sharedTools(rawTools: unknown[] | undefined): ChatToolDefinition[] | undefined {
  if (!rawTools?.length) return undefined;
  const tools = rawTools.flatMap((raw, index) => {
    if (!isRecord(raw) || raw.type !== "function") return [];
    const fn = isRecord(raw.function) ? raw.function : raw;
    const name = stringValue(fn.name);
    if (!name) throw new BridgeRequestError(400, `tools[${index}] is missing function.name`);
    const parameters = isRecord(fn.parameters) ? fn.parameters : { type: "object", properties: {} };
    return [{ type: "function" as const, function: { name, description: stringValue(fn.description) ?? "", parameters } }];
  });
  return tools.length ? tools : undefined;
}

function sseFrame(value: unknown): string {
  return `data: ${JSON.stringify(value)}\n\n`;
}

function completionId(): string {
  return `chatcmpl-${randomUUID().replace(/-/gu, "").slice(0, 24)}`;
}

interface CollectedCompletion {
  text: string;
  toolCalls: ToolCall[];
}

async function collectCompletion(runtime: ProviderRuntime, request: { model: string; messages: ChatMessage[]; attachments: ChatAttachment[]; tools?: ChatToolDefinition[]; signal: AbortSignal }): Promise<CollectedCompletion> {
  const text: string[] = [];
  const toolCalls: ToolCall[] = [];
  for await (const event of runtime.streamChat(request)) {
    if (event.type === "text") text.push(event.delta);
    if (event.type === "tool_call") toolCalls.push(event.call);
  }
  return { text: text.join(""), toolCalls };
}

function modelRow(model: ModelInfo): RecordValue {
  return {
    id: model.id,
    object: "model",
    created: 0,
    owned_by: "openwordcode",
    name: model.name,
    context_length: model.contextWindow,
    supports_streaming: model.capabilities.streaming,
    supports_tools: model.capabilities.tools,
    supports_reasoning_effort: model.capabilities.reasoning === true,
    supports_vision: model.capabilities.vision === true,
  };
}

function messageResponse(model: string, result: CollectedCompletion): RecordValue {
  const message: RecordValue = { role: "assistant", content: result.text || null };
  if (result.toolCalls.length) {
    message.tool_calls = result.toolCalls.map(call => ({ id: call.id, type: "function", function: { name: call.name, arguments: call.arguments || "{}" } }));
  }
  return {
    id: completionId(),
    object: "chat.completion",
    created: Math.floor(Date.now() / 1_000),
    model,
    choices: [{ index: 0, message, finish_reason: result.toolCalls.length ? "tool_calls" : "stop" }],
  };
}

async function streamCompletion(reply: FastifyReply, request: FastifyRequest, runtime: ProviderRuntime, body: { model: string; messages: ChatMessage[]; attachments: ChatAttachment[]; tools?: ChatToolDefinition[] }): Promise<void> {
  const id = completionId();
  const created = Math.floor(Date.now() / 1_000);
  const controller = new AbortController();
  request.raw.on("close", () => controller.abort());
  reply.hijack();
  reply.raw.writeHead(200, { "Content-Type": "text/event-stream; charset=utf-8", "Cache-Control": "no-cache, no-transform", Connection: "keep-alive", "X-Accel-Buffering": "no" });
  try {
    reply.raw.write(sseFrame({ id, object: "chat.completion.chunk", created, model: body.model, choices: [{ index: 0, delta: { role: "assistant" }, finish_reason: null }] }));
    let toolIndex = 0;
    for await (const event of runtime.streamChat({ ...body, signal: controller.signal })) {
      if (event.type === "text") {
        reply.raw.write(sseFrame({ id, object: "chat.completion.chunk", created, model: body.model, choices: [{ index: 0, delta: { content: event.delta }, finish_reason: null }] }));
      }
      if (event.type === "tool_call") {
        const index = toolIndex;
        toolIndex += 1;
        reply.raw.write(sseFrame({ id, object: "chat.completion.chunk", created, model: body.model, choices: [{ index: 0, delta: { tool_calls: [{ index, id: event.call.id, type: "function", function: { name: event.call.name, arguments: event.call.arguments || "{}" } }] }, finish_reason: null }] }));
      }
    }
    reply.raw.write(sseFrame({ id, object: "chat.completion.chunk", created, model: body.model, choices: [{ index: 0, delta: {}, finish_reason: toolIndex ? "tool_calls" : "stop" }] }));
    reply.raw.write("data: [DONE]\n\n");
  } catch (error) {
    reply.raw.write(sseFrame(errorBody(error)));
    reply.raw.write("data: [DONE]\n\n");
  } finally {
    if (!reply.raw.writableEnded) reply.raw.end();
  }
}

async function proxyResponses(state: CoreState, request: FastifyRequest, reply: FastifyReply, body: RecordValue): Promise<void> {
  const model = stringValue(body.model);
  if (!model) throw new BridgeRequestError(400, "model is required");
  const route = await routeForModel(state, model);
  if (route.provider.kind !== "openai-codex") throw new BridgeRequestError(400, "Responses passthrough is available only for the OpenWordCode account transport");
  const credential = await state.auth.resolveOAuth(route.provider);
  if (!credential) throw new ProviderError("auth_required", "Sign in to the OpenWordCode account before using this bridge tool", 401);
  const upstream = await fetch(`${normalizeBaseUrl(route.provider.baseUrl)}/responses`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: request.headers.accept?.includes("text/event-stream") ? "text/event-stream" : "application/json",
      Authorization: `Bearer ${credential.accessToken}`,
      ...(credential.accountId ? { "ChatGPT-Account-Id": credential.accountId } : {}),
    },
    body: JSON.stringify(body),
  });
  if (!upstream.ok) {
    const detail = await upstream.text().catch(() => "");
    throw new ProviderError("provider_http_error", detail || `OpenWordCode account returned HTTP ${upstream.status}`, upstream.status);
  }
  if (!upstream.body) throw new BridgeRequestError(502, "OpenWordCode account returned an empty response");
  reply.hijack();
  reply.raw.writeHead(upstream.status, {
    "Content-Type": upstream.headers.get("content-type") ?? "application/json",
    "Cache-Control": "no-cache, no-transform",
    "X-Accel-Buffering": "no",
  });
  const reader = upstream.body.getReader();
  try {
    while (true) {
      const chunk = await reader.read();
      if (chunk.done) break;
      reply.raw.write(Buffer.from(chunk.value));
    }
  } finally {
    reader.releaseLock();
    if (!reply.raw.writableEnded) reply.raw.end();
  }
}

export async function buildBridgeServer(state: CoreState): Promise<FastifyInstance> {
  const app = Fastify({ logger: false, bodyLimit: 32_000_000 });
  app.addHook("onRequest", async request => {
    if (request.url === "/healthz" || request.url === "/health" || request.url === "/readyz") return;
    authorizeBridgeRequest(state, request);
  });
  app.setErrorHandler((error, _request, reply) => reply.code(errorStatus(error)).send(errorBody(error)));
  app.get("/healthz", async () => ({ status: "ok", service: "OpenWordCode Bridge" }));
  app.get("/health", async () => ({ status: "ok", service: "OpenWordCode Bridge" }));
  app.get("/readyz", async () => {
    try {
      const catalog = await modelCatalog(state);
      return { status: catalog.models.length ? "ready" : "pending", service: "OpenWordCode Bridge", models: catalog.models.length };
    } catch (error) {
      return { status: "failed", service: "OpenWordCode Bridge", detail: safeErrorMessage(error) };
    }
  });
  app.get("/v1/models", async () => {
    const catalog = await modelCatalog(state);
    return { object: "list", data: catalog.models.map(modelRow) };
  });
  app.get<{ Params: { id: string } }>("/v1/models/:id", async request => {
    const catalog = await modelCatalog(state);
    const model = catalog.models.find(item => item.id === request.params.id);
    if (!model) throw new BridgeRequestError(404, `Model not found: ${request.params.id}`);
    return modelRow(model);
  });
  app.post("/v1/chat/completions", async (request, reply) => {
    const parsed = bridgeBodySchema.safeParse(request.body);
    if (!parsed.success) throw new BridgeRequestError(400, parsed.error.issues.map(issue => `${issue.path.join(".")}: ${issue.message}`).join("; "));
    const { messages, attachments } = sharedMessages(parsed.data.messages);
    const tools = sharedTools(parsed.data.tools);
    const route = await routeForModel(state, parsed.data.model);
    const body = { model: route.model.id || parsed.data.model, messages, attachments, ...(tools ? { tools } : {}) };
    if (parsed.data.stream !== false) {
      await streamCompletion(reply, request, runtimeForBridge(state, route.provider), body);
      return reply;
    }
    const result = await collectCompletion(runtimeForBridge(state, route.provider), { ...body, signal: new AbortController().signal });
    return messageResponse(parsed.data.model, result);
  });
  app.post("/v1/responses", async (request, reply) => {
    if (!isRecord(request.body)) throw new BridgeRequestError(400, "request body must be an object");
    await proxyResponses(state, request, reply, request.body);
    return reply;
  });
  return app;
}

export async function startBridgeServer(state: CoreState): Promise<FastifyInstance> {
  const app = await buildBridgeServer(state);
  await app.listen({ host: "127.0.0.1", port: bridgePort(state.env) });
  return app;
}

export async function inspectBridge(env: NodeJS.ProcessEnv = process.env): Promise<OpenWordCodeBridgeStatus> {
  const endpoint = bridgeRootUrl(env);
  const headers: HeadersInit = {};
  const token = env.OPENWORDCODE_BRIDGE_TOKEN?.trim();
  if (token) headers["x-openwordcode-bridge-token"] = token;
  try {
    const health = await fetch(`${endpoint}/healthz`, { headers, signal: AbortSignal.timeout(1_500) });
    if (!health.ok) return { available: false, endpoint, models: 0, detail: `OpenWordCode Bridge returned HTTP ${health.status}` };
    const models = await fetch(`${endpoint}/v1/models`, { headers, signal: AbortSignal.timeout(1_500) });
    if (!models.ok) return { available: false, endpoint, models: 0, detail: `OpenWordCode Bridge model catalog returned HTTP ${models.status}` };
    const payload = await models.json() as { data?: unknown };
    return { available: true, endpoint, models: Array.isArray(payload.data) ? payload.data.length : 0, detail: "OpenWordCode Bridge is running locally" };
  } catch (error) {
    return { available: false, endpoint, models: 0, detail: safeErrorMessage(error) || "OpenWordCode Bridge is not running" };
  }
}
