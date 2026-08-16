import Fastify, { type FastifyInstance, type FastifyReply, type FastifyRequest } from "fastify";
import fastifyStatic from "@fastify/static";
import { existsSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, join, resolve } from "node:path";
import { z } from "zod";
import {
  OPENWORDCODE_VERSION,
  type AgentEvent,
  type AgentAction,
  type AgentRequest,
  type ProviderConfig,
  type ProviderId,
} from "../../../packages/shared/src/index.js";
import { EncryptedFileCredentialStore } from "../../../packages/auth/src/index.js";
import { createOpaqueToken, equalSecret, isLoopbackUrl, safeErrorMessage } from "../../../packages/security/src/index.js";
import { createProvider, providerErrorDetail, type ProviderRuntime } from "../../../packages/providers/src/index.js";
import { buildReadOnlyContext, runAgent, type AgentRuntimeOptions } from "../../../packages/agent-core/src/index.js";
import { AuthManager } from "./auth.js";
import { ChatGPTOAuthManager, OAuthConfigurationError } from "./chatgpt-auth.js";
import { oauthProviderCatalog, oauthProviderIsSupported, ProviderOAuthManager } from "./provider-oauth.js";
import { localCliDisplayName, localCliProviderIsSupported } from "./local-cli-auth.js";
import { ChangeStore, StaleChangeError } from "./changes.js";
import { AgentActionStore } from "./actions.js";
import { executeConsoleCommand, validateConsoleCommand } from "./console.js";
import { searchWithOpenWordCodeBridge } from "./web-search.js";
import { inspectBridge, startBridgeServer, type OpenWordCodeBridgeStatus } from "./bridge.js";
import { dataDirectory, loadSettings, providerConfig, saveSettings, type Settings, updateProvider } from "./config.js";

const PORT = 10_200;
const bodyProviderId = z.string().trim().min(1).max(80);
// Some Word desktop builds expose Range.start/end as -1 when the selection
// position is unavailable. Treat that as missing metadata rather than
// rejecting the entire agent request.
const rangePositionSchema = z.preprocess(
  value => typeof value === "number" && Number.isInteger(value) && value >= 0 ? value : undefined,
  z.number().int().min(0).max(100_000_000).optional(),
);

const targetSchema = z.object({ kind: z.enum(["selection", "paragraph", "document", "visual"]), id: z.string().min(1).max(200), beforeText: z.string().max(50_000), beforeFingerprint: z.string().min(1).max(32), paragraphIndex: z.number().int().min(0).max(10_000).optional(), visualKind: z.enum(["inlinePicture", "shape"]).optional(), visualIndex: z.number().int().min(0).max(10_000).optional(), shapeId: z.number().int().min(0).optional() });
const visualElementSchema = z.object({
  id: z.string().min(1).max(200),
  kind: z.enum(["inlinePicture", "shape"]),
  index: z.number().int().min(0).max(10_000),
  shapeId: z.number().int().min(0).optional(),
  shapeType: z.string().max(100).optional(),
  name: z.string().max(500).optional(),
  altTextTitle: z.string().max(1_000).optional(),
  altTextDescription: z.string().max(4_000).optional(),
  hyperlink: z.string().max(4_000).optional(),
  imageFormat: z.string().max(100).optional(),
  mimeType: z.enum(["image/gif", "image/jpeg", "image/png", "image/webp"]).optional(),
  width: z.number().finite().min(-20_000).max(20_000).optional(),
  height: z.number().finite().min(-20_000).max(20_000).optional(),
  x: z.number().finite().min(-20_000).max(20_000).optional(),
  y: z.number().finite().min(-20_000).max(20_000).optional(),
  rotation: z.number().finite().min(-10_000).max(10_000).optional(),
  relativeHorizontalPosition: z.string().max(100).optional(),
  relativeVerticalPosition: z.string().max(100).optional(),
  wrapType: z.string().max(100).optional(),
  anchorParagraphIndex: z.number().int().min(0).max(10_000).optional(),
  anchorText: z.string().max(1_000).optional(),
  rangeStart: rangePositionSchema,
  rangeEnd: rangePositionSchema,
  dataUrl: z.string().max(8_500_000).regex(/^data:image\/(?:gif|jpeg|png|webp);base64,[A-Za-z0-9+/]+=*$/).optional(),
  size: z.number().int().min(1).max(6_291_456).optional(),
  contentAvailable: z.boolean(),
  contentOmittedReason: z.string().max(500).optional(),
});
const documentSchema = z.object({
  documentId: z.string().min(1).max(200),
  title: z.string().max(500).optional(),
  selection: z.object({ text: z.string().max(20_000), isEmpty: z.boolean(), isTable: z.boolean().optional(), tableCount: z.number().int().min(0).max(100).optional(), rangeStart: rangePositionSchema, rangeEnd: rangePositionSchema, selectedVisualElementIds: z.array(z.string().min(1).max(200)).max(40).optional(), target: targetSchema }),
  documentText: z.string().max(50_000),
  paragraphs: z.array(z.object({ id: z.string().min(1).max(200), index: z.number().int().min(0).max(10_000), text: z.string().max(20_000), style: z.string().max(200).optional() })).max(300),
  visualElements: z.array(visualElementSchema).max(40).superRefine((elements, context) => {
    if (elements.filter(element => element.dataUrl).length > 4) context.addIssue({ code: z.ZodIssueCode.custom, message: "too many embedded pictures with pixel data" });
    if (elements.reduce((total, element) => total + (element.size ?? 0), 0) > 12_582_912) context.addIssue({ code: z.ZodIssueCode.custom, message: "embedded picture data exceeds 12 MB" });
  }).optional(),
  visualContentTruncated: z.boolean().optional(),
  outline: z.array(z.object({ id: z.string().min(1).max(200), text: z.string().max(20_000), level: z.number().int().min(1).max(10), index: z.number().int().min(0).max(10_000) })).max(200),
  capabilities: z.object({ canRead: z.boolean(), canWrite: z.boolean(), canComment: z.boolean(), canFormat: z.boolean(), host: z.string().max(100).optional(), platform: z.string().max(100).optional() }),
  truncated: z.boolean().optional(),
});

const attachmentSchema = z.object({
  id: z.string().min(1).max(120),
  name: z.string().min(1).max(260),
  mimeType: z.enum(["application/pdf", "image/gif", "image/jpeg", "image/png", "image/webp"]),
  size: z.number().int().min(1).max(6_000_000),
  dataUrl: z.string().max(8_500_000).regex(/^data:(?:application\/pdf|image\/(?:gif|jpeg|png|webp));base64,[A-Za-z0-9+/]+=*$/),
});
const attachmentsSchema = z.array(attachmentSchema).max(4).superRefine((files, context) => {
  if (files.reduce((total, file) => total + file.size, 0) > 12_000_000) context.addIssue({ code: z.ZodIssueCode.custom, message: "attachments exceed the 12 MB total limit" });
});

const skillSchema = z.object({
  id: z.string().min(1).max(120),
  name: z.string().min(1).max(120),
  description: z.string().max(500).optional().default(""),
  instructions: z.string().min(1).max(30_000),
  isDefault: z.boolean().optional(),
  enabled: z.boolean().optional(),
  author: z.string().max(120).optional(),
  version: z.string().max(40).optional(),
});
const skillsSchema = z.array(skillSchema).max(12);

const agentSchema = z.object({
  providerId: bodyProviderId,
  modelId: z.string().trim().min(1).max(200),
  instruction: z.string().trim().min(1).max(8_000),
  mode: z.enum(["manual", "auto", "skip"]),
  document: documentSchema,
  attachments: attachmentsSchema.optional(),
  conversation: z.array(z.object({ role: z.enum(["system", "user", "assistant", "tool"]), content: z.string().max(20_000), toolCallId: z.string().max(200).optional() })).max(12).optional(),
  skills: skillsSchema.optional(),
  tools: z.object({ webSearch: z.boolean().optional(), console: z.boolean().optional() }).strict().optional(),
});

const authMethodSchema = z.object({ method: z.enum(["environment", "none", "existing-session", "oauth"]), envVar: z.string().regex(/^[A-Za-z_][A-Za-z0-9_]*$/).optional() });
const customProviderSchema = z.object({ id: z.string().trim().regex(/^[a-z0-9][a-z0-9._-]{1,60}$/), displayName: z.string().trim().min(1).max(120), baseUrl: z.string().url(), authMethod: z.enum(["environment", "none"]).default("none"), envVar: z.string().regex(/^[A-Za-z_][A-Za-z0-9_]*$/).optional(), defaultModel: z.string().trim().max(200).optional(), privacyNote: z.string().trim().max(500).optional() });
const settingsPatchSchema = z.object({ selectedProviderId: z.string().min(1).max(80).optional(), selectedModelId: z.string().max(200).optional(), mode: z.enum(["manual", "auto", "skip"]).optional(), theme: z.enum(["light", "dark", "system"]).optional() }).strict();

class ApiError extends Error {
  constructor(readonly statusCode: number, readonly code: string, message: string) { super(message); }
}

export interface CoreState {
  settings: Settings;
  readonly credentials: EncryptedFileCredentialStore;
  readonly auth: AuthManager;
  readonly chatGptAuth: ChatGPTOAuthManager;
  readonly providerOAuth: ProviderOAuthManager;
  readonly changes: ChangeStore;
  readonly actions: AgentActionStore;
  readonly sessionToken: string;
  readonly env: NodeJS.ProcessEnv;
  readonly workspaceRoot: string;
}

export function createCoreState(env: NodeJS.ProcessEnv = process.env): CoreState {
  const settings = loadSettings(env);
  const credentials = new EncryptedFileCredentialStore(join(dataDirectory(env), "credentials"));
  const chatGptAuth = new ChatGPTOAuthManager(credentials, env);
  const providerOAuth = new ProviderOAuthManager(credentials, env);
  return { settings, credentials, chatGptAuth, providerOAuth, auth: new AuthManager(credentials, env, chatGptAuth, providerOAuth), changes: new ChangeStore(), actions: new AgentActionStore(), sessionToken: env.OPENWORDCODE_SESSION_TOKEN?.trim() || createOpaqueToken("owc_session_"), env, workspaceRoot: resolve(process.cwd()) };
}

function header(request: FastifyRequest, name: string): string | undefined {
  const value = request.headers[name];
  return Array.isArray(value) ? value[0] : value;
}

function allowedOrigins(state: CoreState): Set<string> {
  return new Set(state.settings.allowedOrigins);
}

function originAllowed(state: CoreState, origin: string | undefined): boolean {
  return !origin || allowedOrigins(state).has(origin);
}

function requireBody<T>(schema: z.ZodType<T>, body: unknown): T {
  const parsed = schema.safeParse(body);
  if (!parsed.success) throw new ApiError(400, "invalid_request", parsed.error.issues.map(issue => `${issue.path.join(".")}: ${issue.message}`).join("; "));
  return parsed.data;
}

function providerIdFromRequest(request: FastifyRequest<{ Params: { id: string } }>): ProviderId {
  const value = bodyProviderId.safeParse(request.params.id);
  if (!value.success) throw new ApiError(400, "invalid_provider", "provider id is invalid");
  return value.data as ProviderId;
}

function requireProvider(state: CoreState, id: ProviderId): ProviderConfig {
  try { return providerConfig(state.settings, id); } catch { throw new ApiError(404, "provider_not_found", `Unknown provider: ${id}`); }
}

function runtimeFor(state: CoreState, id: ProviderId): ProviderRuntime {
  const config = requireProvider(state, id);
  if (!config.enabled) throw new ApiError(409, "provider_disabled", `${config.displayName} is disabled`);
  return createProvider(config, { store: state.credentials, env: state.env, oauth: provider => state.auth.resolveOAuth(provider) });
}

function providerAuthDto(provider: ProviderConfig, auth: Awaited<ReturnType<AuthManager["status"]>>): Record<string, unknown> {
  return {
    ...provider,
    auth: {
      status: auth.status,
      method: auth.method,
      detail: auth.detail,
      availableMethods: auth.availableMethods,
      credentialConfigured: auth.credentialConfigured,
      environmentConfigured: auth.environmentConfigured,
      ...(auth.source ? { source: auth.source } : {}),
    },
  };
}

async function providerSummary(state: CoreState, provider: ProviderConfig): Promise<Record<string, unknown>> {
  return providerAuthDto(provider, await state.auth.status(provider));
}

function saveState(state: CoreState): void { saveSettings(state.settings, state.env); }

function sendSse(reply: FastifyReply, event: AgentEvent): void {
  if (!reply.raw.writableEnded) reply.raw.write(`data: ${JSON.stringify(event)}\n\n`);
}

function writeSseHeaders(reply: FastifyReply): void {
  reply.hijack();
  reply.raw.writeHead(200, { "Content-Type": "text/event-stream; charset=utf-8", "Cache-Control": "no-cache, no-transform", Connection: "keep-alive", "X-Accel-Buffering": "no" });
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/gu, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[character] ?? character));
}

function oauthCallbackPage(detail: string, success: boolean): string {
  const heading = success ? "OpenWordCode is connected" : "OpenWordCode sign-in failed";
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>${heading}</title><style>body{font:16px system-ui,sans-serif;background:#171514;color:#f5f1ee;display:grid;place-items:center;min-height:100vh;margin:0}main{max-width:420px;margin:24px;padding:28px;border:1px solid #3b3532;border-radius:16px;background:#211e1c}h1{font-size:22px;margin:0 0 12px}p{color:#bdb4ae;line-height:1.5}button{border:0;border-radius:9px;background:#d97752;color:#fff;padding:10px 14px;font:inherit;cursor:pointer}</style></head><body><main><h1>${heading}</h1><p>${escapeHtml(detail)}</p><button onclick="window.close()">Return to Word</button></main></body></html>`;
}

export async function buildServer(state = createCoreState()): Promise<FastifyInstance> {
  const app = Fastify({ logger: false, bodyLimit: 32_000_000 });
  app.decorate("owcState", state);

  app.addHook("onRequest", async (request, reply) => {
    const origin = header(request, "origin");
    const isOAuthCallback = request.url.split("?")[0] === "/oauth/chatgpt/callback";
    if (!isOAuthCallback && !originAllowed(state, origin)) {
      reply.code(403).send({ error: "cross-origin request blocked" });
      return;
    }
    if (origin) {
      reply.header("Access-Control-Allow-Origin", origin);
      reply.header("Vary", "Origin");
    }
    reply.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
    reply.header("Access-Control-Allow-Headers", "Content-Type, X-OpenWordCode-Session, X-OpenWordCode-CSRF");
    reply.header("X-Content-Type-Options", "nosniff");
    reply.header("X-Frame-Options", "DENY");
    if (request.method === "OPTIONS") { reply.code(204).send(); return; }
  });

  const protectedRoute = async (request: FastifyRequest, reply: FastifyReply): Promise<void> => {
    const path = request.url.split("?")[0] ?? request.url;
    if (!path.startsWith("/api/") || path === "/api/bootstrap") return;
    if (!equalSecret(header(request, "x-openwordcode-session"), state.sessionToken)) {
      reply.code(401).send({ error: "OpenWordCode session token required" });
      return;
    }
    if (["POST", "PUT", "PATCH", "DELETE"].includes(request.method) && !equalSecret(header(request, "x-openwordcode-csrf"), state.sessionToken)) {
      reply.code(403).send({ error: "CSRF token required" });
    }
  };
  app.addHook("preHandler", protectedRoute);

  app.setErrorHandler((error, _request, reply) => {
    const apiError = error instanceof ApiError ? error : undefined;
    const statusCode = apiError?.statusCode ?? (typeof (error as { statusCode?: unknown }).statusCode === "number" ? Number((error as { statusCode: number }).statusCode) : 500);
    const code = apiError?.code ?? "internal_error";
    reply.code(statusCode).send({ error: code, message: safeErrorMessage(apiError ?? error) });
  });

  app.get("/health", async () => ({ status: "ok", service: "OpenWordCode", version: OPENWORDCODE_VERSION, port: Number(state.env.OPENWORDCODE_PORT ?? PORT) }));
  app.get("/healthz", async () => ({ status: "ok", service: "OpenWordCode", version: OPENWORDCODE_VERSION }));
  app.get("/api/bootstrap", async (request, reply) => {
    const origin = header(request, "origin");
    if (!originAllowed(state, origin)) throw new ApiError(403, "cross_origin", "Origin is not configured for this Core");
    reply.header("Cache-Control", "no-store");
    return { version: OPENWORDCODE_VERSION, sessionToken: state.sessionToken, apiBase: "", dataDirectory: dataDirectory(state.env) };
  });

  app.post("/api/auth/chatgpt/start", async () => {
    try { return state.chatGptAuth.start(); } catch (error) {
      if (error instanceof OAuthConfigurationError) throw new ApiError(409, "oauth_not_configured", error.message);
      throw error;
    }
  });
  app.post("/api/auth/chatgpt/start-codex-cli-login", async () => {
    try {
      const launch = await state.chatGptAuth.startCodexCliLogin();
      return { ok: true, ...launch };
    } catch (error) {
      if (error instanceof OAuthConfigurationError) throw new ApiError(409, "codex_cli_unavailable", error.message);
      throw error;
    }
  });
  app.post("/api/auth/chatgpt/use-codex-cli", async () => {
    try {
      await state.chatGptAuth.useCodexCli();
      const provider = requireProvider(state, "openwordcode-bridge");
      return { ok: true, auth: await state.auth.status(provider) };
    } catch (error) {
      if (error instanceof OAuthConfigurationError) throw new ApiError(409, "codex_cli_session_unavailable", error.message);
      throw error;
    }
  });
  app.get<{ Querystring: { flowId?: string } }>("/api/auth/chatgpt/status", async request => {
    const flowId = request.query.flowId?.trim();
    if (!flowId) throw new ApiError(400, "flow_id_required", "OAuth flow id is required");
    return state.chatGptAuth.flowStatus(flowId);
  });
  app.post("/api/auth/chatgpt/cancel", async request => {
    const body = requireBody(z.object({ flowId: z.string().trim().min(1).max(120) }), request.body);
    state.chatGptAuth.cancel(body.flowId);
    return { ok: true };
  });
  app.post("/api/auth/chatgpt/disconnect", async () => {
    const provider = requireProvider(state, "openwordcode-bridge");
    const next = await state.auth.disconnect(provider);
    state.settings = updateProvider(state.settings, next);
    saveState(state);
    return { ok: true, auth: await state.auth.status(next) };
  });

  app.get("/api/oauth/providers", async () => ({ providers: oauthProviderCatalog() }));
  app.post<{ Params: { provider: string } }>("/api/oauth/:provider/local-cli/start", async request => {
    const providerId = request.params.provider.trim();
    if (!localCliProviderIsSupported(providerId)) throw new ApiError(409, "local_cli_unavailable", `No supported local CLI connector is available for ${providerId}.`);
    const provider = Object.values(state.settings.providers).find(item => item.auth.oauthProvider === providerId && item.enabled);
    if (!provider) throw new ApiError(404, "oauth_provider_not_found", `No enabled provider is configured for ${providerId}.`);
    try {
      const launch = await state.providerOAuth.startLocalCliLogin(providerId);
      const next = state.auth.activateOAuth(provider);
      state.settings = updateProvider(state.settings, next);
      saveState(state);
      return { ok: true, ...launch, auth: await state.auth.status(next) };
    } catch (error) {
      throw new ApiError(409, "local_cli_start_failed", error instanceof Error ? error.message : `${localCliDisplayName(providerId)} login could not be started`);
    }
  });
  app.post<{ Params: { provider: string } }>("/api/oauth/:provider/local-cli/use", async request => {
    const providerId = request.params.provider.trim();
    if (!localCliProviderIsSupported(providerId)) throw new ApiError(409, "local_cli_unavailable", `No supported local CLI connector is available for ${providerId}.`);
    const provider = Object.values(state.settings.providers).find(item => item.auth.oauthProvider === providerId && item.enabled);
    if (!provider) throw new ApiError(404, "oauth_provider_not_found", `No enabled provider is configured for ${providerId}.`);
    try {
      await state.providerOAuth.useLocalCli(providerId);
      const next = state.auth.activateOAuth(provider);
      state.settings = updateProvider(state.settings, next);
      saveState(state);
      return { ok: true, auth: await state.auth.status(next) };
    } catch (error) {
      throw new ApiError(409, "local_cli_session_unavailable", error instanceof Error ? error.message : `No active ${localCliDisplayName(providerId)} session was found`);
    }
  });
  app.post<{ Params: { provider: string } }>("/api/oauth/:provider/start", async request => {
    const providerId = request.params.provider.trim();
    const entry = oauthProviderCatalog().find(item => item.id === providerId);
    if (!entry) throw new ApiError(404, "oauth_provider_not_found", `Unknown OAuth provider: ${providerId}`);
    if (!oauthProviderIsSupported(providerId)) throw new ApiError(409, "oauth_transport_unavailable", entry.detail);
    try {
      return await state.providerOAuth.start(providerId);
    } catch (error) {
      throw new ApiError(409, "oauth_start_failed", error instanceof Error ? error.message : "Could not start OAuth sign-in");
    }
  });
  app.get<{ Params: { flowId: string } }>("/api/oauth/flows/:flowId", async request => {
    let snapshot;
    try { snapshot = state.providerOAuth.flowStatus(request.params.flowId.trim()); } catch (error) { throw new ApiError(404, "oauth_flow_not_found", error instanceof Error ? error.message : "OAuth flow not found"); }
    if (snapshot.status === "connected") {
      const provider = Object.values(state.settings.providers).find(item => item.auth.oauthProvider === snapshot.providerId && item.enabled);
      if (provider && provider.auth.method !== "oauth") {
        state.settings = updateProvider(state.settings, state.auth.activateOAuth(provider));
        saveState(state);
      }
    }
    return snapshot;
  });
  app.post<{ Params: { flowId: string } }>("/api/oauth/flows/:flowId/complete", async request => {
    const body = requireBody(z.object({ code: z.string().trim().min(1).max(4_096) }), request.body);
    try {
      const snapshot = await state.providerOAuth.completeManualCode(request.params.flowId.trim(), body.code);
      if (snapshot.status === "connected") {
        const provider = Object.values(state.settings.providers).find(item => item.auth.oauthProvider === snapshot.providerId && item.enabled);
        if (provider && provider.auth.method !== "oauth") {
          state.settings = updateProvider(state.settings, state.auth.activateOAuth(provider));
          saveState(state);
        }
      }
      return snapshot;
    } catch (error) {
      throw new ApiError(409, "oauth_code_exchange_failed", error instanceof Error ? error.message : "Could not complete sign-in");
    }
  });
  app.post("/api/oauth/flows/cancel", async request => {
    const body = requireBody(z.object({ flowId: z.string().trim().min(1).max(120) }), request.body);
    state.providerOAuth.cancel(body.flowId);
    return { ok: true };
  });

  app.get<{ Querystring: { code?: string; state?: string; error?: string; error_description?: string } }>("/oauth/chatgpt/callback", async (request, reply) => {
    const result = await state.chatGptAuth.handleCallback(request.query);
    const success = result.status === "connected";
    return reply.type("text/html; charset=utf-8").send(oauthCallbackPage(result.detail, success));
  });

  app.get("/api/providers", async () => ({ providers: await Promise.all(Object.values(state.settings.providers).filter(provider => provider.enabled && !provider.internal).map(provider => providerSummary(state, provider))) }));
  app.post("/api/providers", async request => {
    const body = requireBody(customProviderSchema, request.body);
    if (state.settings.providers[body.id]) throw new ApiError(409, "provider_exists", "A provider with this id already exists");
    const parsedUrl = new URL(body.baseUrl);
    if (!["http:", "https:"].includes(parsedUrl.protocol)) throw new ApiError(400, "invalid_provider_url", "Provider URL must use HTTP or HTTPS");
    const authMethod = body.authMethod ?? "none";
    const provider: ProviderConfig = {
      id: body.id,
      displayName: body.displayName,
      kind: "openai-compatible",
      baseUrl: body.baseUrl.replace(/\/+$/, ""),
      enabled: true,
      local: isLoopbackUrl(body.baseUrl),
      auth: authMethod === "environment" ? { method: "environment", envVar: body.envVar || `${body.id.toUpperCase().replace(/[^A-Z0-9]/g, "_")}_API_KEY` } : { method: authMethod },
      ...(body.defaultModel ? { defaultModel: body.defaultModel } : {}),
      privacyNote: body.privacyNote || (isLoopbackUrl(body.baseUrl) ? "Document content is sent to the configured local endpoint." : "Document content is sent to the configured OpenAI-compatible endpoint."),
    };
    state.settings = updateProvider(state.settings, provider);
    saveState(state);
    return { provider: await providerSummary(state, provider) };
  });
  app.get<{ Params: { id: string } }>("/api/providers/:id/auth", async request => {
    const provider = requireProvider(state, providerIdFromRequest(request));
    return { providerId: provider.id, auth: await state.auth.status(provider) };
  });
  app.get<{ Params: { id: string } }>("/api/providers/:id/models", async request => {
    const id = providerIdFromRequest(request);
    const runtime = runtimeFor(state, id);
    try { return { providerId: id, models: await runtime.listModels() }; } catch (error) { throw new ApiError(502, "model_discovery_failed", providerErrorDetail(error)); }
  });
  app.post<{ Params: { id: string } }>("/api/providers/:id/test", async request => {
    const id = providerIdFromRequest(request);
    const runtime = runtimeFor(state, id);
    const started = Date.now();
    try { const models = await runtime.listModels(); return { ok: true, providerId: id, modelCount: models.length, durationMs: Date.now() - started }; } catch (error) { throw new ApiError(502, "provider_test_failed", providerErrorDetail(error)); }
  });
  app.put<{ Params: { id: string } }>("/api/providers/:id/auth", async request => {
    const id = providerIdFromRequest(request);
    const body = requireBody(authMethodSchema, request.body);
    const provider = requireProvider(state, id);
    const next = state.auth.setMethod(provider, body.method, body.envVar);
    if (body.method !== "oauth" && provider.auth.oauthProvider) await state.providerOAuth.clearLocalCli(provider.auth.oauthProvider);
    state.settings = updateProvider(state.settings, next);
    saveState(state);
    return { ok: true, providerId: id, auth: await state.auth.status(next) };
  });
  app.post<{ Params: { id: string } }>("/api/providers/:id/disconnect", async request => {
    const id = providerIdFromRequest(request);
    const provider = requireProvider(state, id);
    const next = await state.auth.disconnect(provider);
    state.settings = updateProvider(state.settings, next);
    saveState(state);
    return { ok: true, providerId: id, auth: await state.auth.status(next) };
  });

  app.get("/api/openwordcode/bridge/status", async (): Promise<OpenWordCodeBridgeStatus> => inspectBridge(state.env));

  app.get("/api/models", async () => {
    const rows: unknown[] = [];
    for (const provider of Object.values(state.settings.providers).filter(item => item.enabled && !item.internal)) {
      try { rows.push(...await runtimeFor(state, provider.id).listModels()); } catch { /* disconnected providers are represented in /api/providers */ }
    }
    return { models: rows };
  });

  app.get("/api/settings", async () => ({ settings: { version: state.settings.version, selectedProviderId: state.settings.selectedProviderId, selectedModelId: state.settings.selectedModelId, mode: state.settings.mode, theme: state.settings.theme }, credentialStore: state.credentials.kind }));
  app.put("/api/settings", async request => {
    const patch = requireBody(settingsPatchSchema, request.body);
    if (patch.selectedProviderId) requireProvider(state, patch.selectedProviderId as ProviderId);
    state.settings = { ...state.settings, ...patch };
    saveState(state);
    return { settings: state.settings };
  });
  app.get("/api/diagnostics", async () => ({ version: OPENWORDCODE_VERSION, core: "connected", dataDirectory: dataDirectory(state.env), credentialStore: state.credentials.kind, selectedProviderId: state.settings.selectedProviderId, selectedModelId: state.settings.selectedModelId }));

  app.get("/api/changes", async () => ({ changes: state.changes.list() }));
  app.post<{ Params: { id: string } }>("/api/changes/:id/reject", async request => ({ change: state.changes.reject(request.params.id) }));
  app.post<{ Params: { id: string } }>("/api/changes/:id/approve", async request => {
    const body = requireBody(z.object({ currentBefore: z.string().max(50_000) }), request.body);
    try { return { change: state.changes.approve(request.params.id, body.currentBefore) }; } catch (error) { if (error instanceof StaleChangeError) throw new ApiError(409, error.code, error.message); throw error; }
  });
  app.post<{ Params: { id: string } }>("/api/changes/:id/complete", async request => {
    const body = requireBody(z.object({ success: z.boolean(), reason: z.string().max(500).optional() }), request.body);
    return { change: state.changes.complete(request.params.id, body.success, body.reason) };
  });

  app.get("/api/actions", async () => ({ actions: state.actions.list() }));
  app.post<{ Params: { id: string } }>("/api/actions/:id/reject", async request => ({ action: state.actions.reject(request.params.id) }));
  app.post<{ Params: { id: string } }>("/api/actions/:id/approve", async request => {
    const approved = state.actions.approve(request.params.id);
    const validation = validateConsoleCommand(approved.command);
    if (!validation.safe) return { action: state.actions.transition(approved.id, "failed", { failureReason: validation.reason }) };
    const result = await executeConsoleCommand({ command: approved.command, workingDirectory: approved.workingDirectory, workspaceRoot: state.workspaceRoot });
    return { action: state.actions.transition(approved.id, result.ok ? "completed" : "failed", { output: result.output, ...(result.failureReason ? { failureReason: result.failureReason } : {}) }) };
  });

  app.post("/api/agent", async (request, reply) => {
    const body = requireBody(agentSchema, request.body) as AgentRequest;
    const document = buildReadOnlyContext(body.document);
    const provider = runtimeFor(state, body.providerId as ProviderId);
    const controller = new AbortController();
    request.raw.on("close", () => controller.abort());
    writeSseHeaders(reply);
    sendSse(reply, { type: "status", message: `Using ${provider.config.displayName} · ${body.modelId}` });
    const onEvent: AgentRuntimeOptions["onEvent"] = event => {
      if (event.type === "text") sendSse(reply, { type: "token", delta: event.delta });
      if (event.type === "tool") sendSse(reply, { type: "tool", name: event.name, state: event.state, ...(event.detail ? { detail: event.detail } : {}) });
      if (event.type === "action") {
        state.actions.add(event.action);
        sendSse(reply, { type: "action", action: event.action });
      }
    };
    try {
      const bridgeProvider = state.settings.providers["openwordcode-bridge"];
      const result = await runAgent({
        provider,
        modelId: body.modelId,
        instruction: body.instruction,
        mode: body.mode,
        document,
        attachments: body.attachments,
        conversation: body.conversation,
        skills: body.skills,
        webSearchEnabled: body.tools?.webSearch === true,
        consoleEnabled: body.tools?.console === true,
        searchWeb: body.tools?.webSearch === true && bridgeProvider
          ? query => searchWithOpenWordCodeBridge({ baseUrl: bridgeProvider.baseUrl, model: "gpt-5.6-luna", query, signal: controller.signal })
          : undefined,
        runConsole: body.tools?.console === true
          ? async requestBody => {
            const validation = validateConsoleCommand(requestBody.command);
            const action: AgentAction = {
              id: `action_${randomUUID()}`,
              type: "console",
              command: requestBody.command,
              workingDirectory: requestBody.workingDirectory?.trim() || ".",
              reason: requestBody.reason,
              status: "pending",
              createdAt: new Date().toISOString(),
              ...(validation.safe ? {} : { failureReason: validation.reason }),
            };
            if (requestBody.mode === "manual" || !validation.safe) {
              state.actions.add(action);
              return { action };
            }
            const execution = await executeConsoleCommand({ command: requestBody.command, workingDirectory: requestBody.workingDirectory, workspaceRoot: state.workspaceRoot, signal: controller.signal });
            const completed = { ...action, status: execution.ok ? "completed" as const : "failed" as const, output: execution.output, ...(execution.failureReason ? { failureReason: execution.failureReason } : {}) };
            state.actions.add(completed);
            return { action: completed, output: execution.output };
          }
          : undefined,
        signal: controller.signal,
        onEvent,
      });
      const storedChanges = result.changes.map(change => state.changes.add(change));
      for (const change of storedChanges) sendSse(reply, { type: "proposal", change });
      sendSse(reply, { type: "done", answer: result.answer, changes: storedChanges, actions: result.actions, truncated: result.truncated });
    } catch (error) {
      const message = controller.signal.aborted ? "Generation cancelled" : safeErrorMessage(error);
      sendSse(reply, { type: "error", code: controller.signal.aborted ? "cancelled" : "agent_failed", message });
    } finally {
      if (!reply.raw.writableEnded) reply.raw.end();
    }
  });

  const webRoot = resolve(dirname(fileURLToPath(import.meta.url)), "../../word-addin/dist");
  if (existsSync(webRoot)) await app.register(fastifyStatic, { root: webRoot, prefix: "/", wildcard: false });
  app.setNotFoundHandler(async (request, reply) => {
    if (request.method === "GET" && existsSync(join(webRoot, "index.html"))) return reply.sendFile("index.html");
    return reply.code(404).send({ error: "not_found" });
  });
  return app;
}

export async function startServer(state = createCoreState()): Promise<FastifyInstance> {
  const bridge = await startBridgeServer(state);
  const app = await buildServer(state);
  const configuredPort = Number(state.env.OPENWORDCODE_PORT ?? PORT);
  app.addHook("onClose", async () => { await bridge.close(); });
  try {
    await app.listen({ host: "127.0.0.1", port: configuredPort });
  } catch (error) {
    await bridge.close();
    throw error;
  }
  return app;
}
