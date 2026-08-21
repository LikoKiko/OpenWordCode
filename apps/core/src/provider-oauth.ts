import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import { createHash, randomBytes, randomUUID } from "node:crypto";
import type { CredentialStore } from "../../../packages/auth/src/index.js";
import type { ProviderConfig } from "../../../packages/shared/src/index.js";
import { ANTIGRAVITY_IDE_VERSION, antigravityUserAgent, type ProviderOAuthCredential } from "../../../packages/providers/src/index.js";
import {
  ensureAntigravityProject,
  isLocalCliCredentialLive,
  launchLocalCliLogin,
  localCliDisplayName,
  localCliProviderIsSupported,
  localCliSource,
  readLocalCliCredential,
  refreshLocalCliCredential,
  type LocalCliCredential,
  type LocalCliLoginLaunch,
  type LocalCliProviderId,
  type LocalCliSource,
} from "./local-cli-auth.js";

/**
 * OAuth here is deliberately provider-owned. The add-in never receives a token;
 * Core exchanges the code and writes the result to the encrypted credential store.
 *
 * Some provider catalogs expose more names than OpenWordCode can safely route
 * today. Those names remain visible in the catalog so the limitation is explicit
 * instead of pretending a generic Bearer header is a working transport.
 */
export type SupportedOAuthProviderId =
  | "anthropic"
  | "xai"
  | "kimi"
  | "google-antigravity"
  | "github-copilot"
  | "nous";

export type OAuthProviderCatalogId = SupportedOAuthProviderId | "command-code" | "kiro" | "cursor";

export interface OAuthProviderCatalogEntry {
  id: OAuthProviderCatalogId;
  displayName: string;
  supported: boolean;
  flow: "browser" | "device" | "custom";
  detail: string;
}

export const OAUTH_PROVIDER_CATALOG: readonly OAuthProviderCatalogEntry[] = [
  { id: "anthropic", displayName: "Claude", supported: true, flow: "browser", detail: "Claude account OAuth for Claude Code-compatible access. Experimental outside Anthropic's own clients." },
  { id: "xai", displayName: "xAI / Grok", supported: true, flow: "browser", detail: "xAI account OAuth using PKCE." },
  { id: "kimi", displayName: "Kimi", supported: true, flow: "device", detail: "Kimi Code device authorization." },
  { id: "google-antigravity", displayName: "Google Antigravity", supported: true, flow: "browser", detail: "Google OAuth plus Cloud Code Assist project discovery." },
  { id: "github-copilot", displayName: "GitHub Copilot", supported: true, flow: "device", detail: "GitHub device authorization followed by Copilot token exchange." },
  { id: "nous", displayName: "Nous Portal", supported: true, flow: "device", detail: "Nous Portal device authorization. Refresh tokens are rotated and refreshed lazily." },
  { id: "command-code", displayName: "Command Code", supported: false, flow: "custom", detail: "Requires the Command Code proprietary callback and inference transport." },
  { id: "kiro", displayName: "Kiro", supported: false, flow: "custom", detail: "Requires the AWS CodeWhisperer/Kiro transport and account metadata." },
  { id: "cursor", displayName: "Cursor", supported: false, flow: "custom", detail: "Requires Cursor's proprietary HTTP/2 and protobuf transport." },
];

const SUPPORTED = new Set<SupportedOAuthProviderId>([
  "anthropic",
  "xai",
  "kimi",
  "google-antigravity",
  "github-copilot",
  "nous",
]);

const ANTHROPIC_CLIENT_ID = Buffer.from("OWQxYzI1MGEtZTYxYi00NGQ5LTg4ZWQtNTk0NGQxOTYyZjVl", "base64").toString("utf8");
const ANTHROPIC_AUTHORIZE_URL = "https://claude.ai/oauth/authorize";
const ANTHROPIC_TOKEN_URL = "https://api.anthropic.com/v1/oauth/token";
const ANTHROPIC_OAUTH_BETA = "claude-code-20250219,oauth-2025-04-20";
const ANTHROPIC_SYSTEM_IDENTITY = "You are a Claude agent, built on Anthropic's Claude Agent SDK.";

const XAI_DISCOVERY_URL = "https://auth.x.ai/.well-known/openid-configuration";
const XAI_CLIENT_ID = "b1a00492-073a-47ea-816f-4c329264a828";
const XAI_SCOPE = "openid profile email offline_access grok-cli:access api:access";

const DEFAULT_GOOGLE_CLIENT_ID = String.fromCharCode(49,48,55,49,48,48,54,48,54,48,53,57,49,45,116,109,104,115,115,105,110,50,104,50,49,108,99,114,101,50,51,53,118,116,111,108,111,106,104,52,103,52,48,51,101,112,46,97,112,112,115,46,103,111,111,103,108,101,117,115,101,114,99,111,110,116,101,110,116,46,99,111,109);
const DEFAULT_GOOGLE_CLIENT_SECRET = String.fromCharCode(71,79,67,83,80,88,45,75,53,56,70,87,82,52,56,54,76,100,76,74,49,109,76,66,56,115,88,67,52,122,54,113,68,65,102);
const GOOGLE_AUTHORIZE_URL = "https://accounts.google.com/o/oauth2/v2/auth";
const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";
const GOOGLE_CLOUD_CODE_API = "https://cloudcode-pa.googleapis.com";
const GOOGLE_CLOUD_CODE_DAILY_API = "https://daily-cloudcode-pa.googleapis.com";
const GOOGLE_API_VERSION = "v1internal";
const GOOGLE_SCOPES = [
  "https://www.googleapis.com/auth/cloud-platform",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/userinfo.profile",
  "https://www.googleapis.com/auth/cclog",
  "https://www.googleapis.com/auth/experimentsandconfigs",
].join(" ");

const KIMI_CLIENT_ID = "17e5f671-d194-4dfb-9706-5516cb48c098";
const KIMI_OAUTH_HOST = "https://auth.kimi.com";
const KIMI_CLI_VERSION = "0.14.0";

const GITHUB_CLIENT_ID = "Iv1.b507a08c87ecfe98";
const GITHUB_DEVICE_URL = "https://github.com/login/device/code";
const GITHUB_TOKEN_URL = "https://github.com/login/oauth/access_token";
const GITHUB_COPILOT_TOKEN_URL = "https://api.github.com/copilot_internal/v2/token";
const GITHUB_USER_URL = "https://api.github.com/user";
const GITHUB_COPILOT_BASE_URL = "https://api.githubcopilot.com";

const NOUS_BASE_URL = "https://portal.nousresearch.com";
const NOUS_CLIENT_ID = "hermes-cli";
const NOUS_INFERENCE_BASE_URL = "https://inference-api.nousresearch.com/v1";
const OAUTH_REQUEST_TIMEOUT_MS = 30_000;
const MAX_DEVICE_POLL_INTERVAL_MS = 30_000;

type BrowserOAuthProviderId = "anthropic" | "xai" | "google-antigravity";

const LOOPBACK: Record<BrowserOAuthProviderId, { port: number; authorizePath: string; hostname: "localhost" | "127.0.0.1" }> = {
  // Anthropic's public Claude Code OAuth client has localhost registered. The
  // hostname is part of the OAuth redirect URI and cannot be substituted with
  // another spelling of loopback (Anthropic rejects 127.0.0.1 here).
  anthropic: { port: 54545, authorizePath: "/callback", hostname: "localhost" },
  xai: { port: 56121, authorizePath: "/callback", hostname: "127.0.0.1" },
  "google-antigravity": { port: 51121, authorizePath: "/callback", hostname: "127.0.0.1" },
};

export function oauthRedirectUri(providerId: BrowserOAuthProviderId): string {
  const callback = LOOPBACK[providerId];
  return `http://${callback.hostname}:${callback.port}${callback.authorizePath}`;
}

interface StoredOAuthCredential {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  accountId?: string;
  email?: string;
  projectId?: string;
  apiBaseUrl?: string;
}

interface TokenPayload {
  access_token?: unknown;
  refresh_token?: unknown;
  expires_in?: unknown;
  id_token?: unknown;
  expires_at?: unknown;
  refresh_in?: unknown;
  error?: unknown;
  error_description?: unknown;
  token?: unknown;
  endpoints?: { api?: unknown };
}

interface DeviceFlowData {
  kind: "kimi" | "github-copilot" | "nous";
  deviceCode: string;
  intervalMs: number;
  deadline: number;
  verificationUrl: string;
  userCode: string;
}

interface DevicePollResult {
  credential: StoredOAuthCredential | null;
  slowDown: boolean;
  retryAfterMs?: number;
}

interface OAuthFlow {
  id: string;
  providerId: SupportedOAuthProviderId;
  status: "pending" | "connected" | "error" | "cancelled";
  detail: string;
  createdAt: string;
  expiresAt: string;
  controller: AbortController;
  authorizeUrl?: string;
  verificationUrl?: string;
  userCode?: string;
  state?: string;
  verifier?: string;
  redirectUri?: string;
  server?: Server;
  servers?: Server[];
  device?: DeviceFlowData;
}

export interface OAuthFlowSnapshot {
  flowId: string;
  providerId: SupportedOAuthProviderId;
  status: OAuthFlow["status"];
  detail: string;
  createdAt: string;
  expiresAt: string;
  authorizeUrl?: string;
  verificationUrl?: string;
  userCode?: string;
}

export interface OAuthStatus {
  status: "login-required" | "connected" | "expired" | "unsupported" | "error";
  detail: string;
  credentialConfigured: boolean;
  source?: "oauth" | LocalCliSource;
}

function isSupported(value: string): value is SupportedOAuthProviderId {
  return SUPPORTED.has(value as SupportedOAuthProviderId);
}

export function oauthProviderIsSupported(value: string | undefined): value is SupportedOAuthProviderId {
  return typeof value === "string" && isSupported(value);
}

export function oauthProviderCatalog(): OAuthProviderCatalogEntry[] {
  return OAUTH_PROVIDER_CATALOG.map(entry => ({ ...entry }));
}

function randomSecret(): string {
  return randomBytes(32).toString("base64url");
}

function pkceChallenge(verifier: string): string {
  return createHash("sha256").update(verifier, "utf8").digest("base64url");
}

function decodeJwt(token: string | undefined): Record<string, unknown> | undefined {
  if (!token) return undefined;
  const part = token.split(".")[1];
  if (!part) return undefined;
  try { return JSON.parse(Buffer.from(part, "base64url").toString("utf8")) as Record<string, unknown>; } catch { return undefined; }
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function requestSignal(signal?: AbortSignal): AbortSignal {
  const timeout = AbortSignal.timeout(OAUTH_REQUEST_TIMEOUT_MS);
  return signal ? AbortSignal.any([signal, timeout]) : timeout;
}

export function validateXaiOAuthEndpoint(rawUrl: string): string {
  const parsed = new URL(rawUrl);
  const host = parsed.hostname.toLowerCase();
  if (parsed.protocol !== "https:"
    || parsed.username
    || parsed.password
    || (parsed.port && parsed.port !== "443")
    || (host !== "x.ai" && !host.endsWith(".x.ai"))) {
    throw new Error(`xAI OAuth discovery returned an unexpected endpoint: ${rawUrl}`);
  }
  return parsed.toString();
}

export function nextDevicePollInterval(currentMs: number, retryAfterMs?: number): number {
  return Math.min(MAX_DEVICE_POLL_INTERVAL_MS, Math.max(currentMs + 5_000, retryAfterMs ?? 0));
}

function expiresAt(payload: TokenPayload, fallbackMs = 60 * 60 * 1000): number {
  if (typeof payload.expires_at === "number" && Number.isFinite(payload.expires_at)) {
    return payload.expires_at > 10_000_000_000 ? payload.expires_at - 120_000 : payload.expires_at * 1000 - 120_000;
  }
  const seconds = typeof payload.expires_in === "number" && Number.isFinite(payload.expires_in) && payload.expires_in >= 0
    ? payload.expires_in
    : fallbackMs / 1000;
  return Date.now() + seconds * 1000 - 120_000;
}

function callbackPage(success: boolean, detail: string): string {
  const heading = success ? "OpenWordCode is connected" : "OpenWordCode sign-in failed";
  const escaped = detail.replace(/[&<>"']/gu, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[character] ?? character));
  return `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>${heading}</title><style>body{font:16px system-ui,sans-serif;background:#171514;color:#f5f1ee;display:grid;place-items:center;min-height:100vh;margin:0}main{max-width:440px;margin:24px;padding:28px;border:1px solid #3b3532;border-radius:16px;background:#211e1c}h1{font-size:22px;margin:0 0 12px}p{color:#bdb4ae;line-height:1.5}button{border:0;border-radius:9px;background:#d97752;color:#fff;padding:10px 14px;font:inherit;cursor:pointer}</style></head><body><main><h1>${heading}</h1><p>${escaped}</p><button onclick="window.close()">Return to Word</button></main></body></html>`;
}

function sendHtml(response: ServerResponse, success: boolean, detail: string): void {
  response.statusCode = success ? 200 : 400;
  response.setHeader("Content-Type", "text/html; charset=utf-8");
  response.end(callbackPage(success, detail));
}

function requestHeaders(providerId: SupportedOAuthProviderId): Record<string, string> {
  if (providerId === "kimi") {
    return {
      "User-Agent": `KimiCLI/${KIMI_CLI_VERSION}`,
      "X-Msh-Platform": "kimi_code_cli",
      "X-Msh-Version": KIMI_CLI_VERSION,
      "X-Msh-Device-Id": "openwordcode",
    };
  }
  if (providerId === "github-copilot") {
    return {
      "Editor-Version": "openwordcode/0.1.0",
      "Editor-Plugin-Version": "openwordcode/0.1.0",
      "Copilot-Integration-Id": "vscode-chat",
      "User-Agent": "openwordcode",
    };
  }
  return {};
}

async function readJson(response: Response): Promise<TokenPayload & Record<string, unknown>> {
  const value: unknown = await response.json().catch(() => ({}));
  return value && typeof value === "object" && !Array.isArray(value) ? value as TokenPayload & Record<string, unknown> : {};
}

async function postForm(fetchImpl: typeof fetch, url: string, values: Record<string, string>, headers: Record<string, string> = {}, signal?: AbortSignal): Promise<TokenPayload & Record<string, unknown>> {
  const response = await fetchImpl(url, {
    method: "POST",
    headers: { Accept: "application/json", "Content-Type": "application/x-www-form-urlencoded", ...headers },
    body: new URLSearchParams(values).toString(),
    redirect: "error",
    signal: requestSignal(signal),
  });
  const payload = await readJson(response);
  if (!response.ok) {
    const code = stringValue(payload.error);
    throw new Error(`${url} rejected OAuth (${response.status}${code ? `: ${code}` : ""})`);
  }
  return payload;
}

async function postJson(fetchImpl: typeof fetch, url: string, values: Record<string, unknown>, signal?: AbortSignal): Promise<TokenPayload & Record<string, unknown>> {
  const response = await fetchImpl(url, {
    method: "POST",
    headers: { Accept: "application/json", "Content-Type": "application/json" },
    body: JSON.stringify(values),
    redirect: "error",
    signal: requestSignal(signal),
  });
  const payload = await readJson(response);
  if (!response.ok) {
    const code = stringValue(payload.error);
    throw new Error(`${url} rejected OAuth (${response.status}${code ? `: ${code}` : ""})`);
  }
  return payload;
}

function asStored(providerId: SupportedOAuthProviderId, payload: TokenPayload & Record<string, unknown>, refreshFallback?: string, extras: Partial<StoredOAuthCredential> = {}): StoredOAuthCredential {
  let accessToken = stringValue(payload.access_token);
  if (providerId === "github-copilot") accessToken = stringValue(payload.token);
  if (!accessToken) throw new Error(`${providerId} OAuth response did not include an access token`);
  const refreshToken = stringValue(payload.refresh_token) ?? refreshFallback;
  if (["anthropic", "xai", "kimi", "google-antigravity", "nous"].includes(providerId) && !refreshToken) {
    throw new Error(`${providerId} OAuth response did not include a refresh token`);
  }
  const claims = decodeJwt(accessToken);
  const idTokenClaims = decodeJwt(stringValue(payload.id_token));
  const accountId = extras.accountId ?? stringValue(claims?.sub) ?? stringValue(idTokenClaims?.sub);
  const email = extras.email ?? stringValue(claims?.email) ?? stringValue(idTokenClaims?.email);
  return {
    accessToken,
    ...(refreshToken ? { refreshToken } : {}),
    expiresAt: extras.expiresAt ?? expiresAt(payload, providerId === "nous" ? 12 * 60 * 60 * 1000 : undefined),
    ...(accountId ? { accountId } : {}),
    ...(email ? { email: email.toLowerCase() } : {}),
    ...(extras.projectId ? { projectId: extras.projectId } : {}),
    ...(extras.apiBaseUrl ? { apiBaseUrl: extras.apiBaseUrl } : {}),
  };
}

function asNousStored(payload: TokenPayload & Record<string, unknown>, submittedRefreshToken?: string): StoredOAuthCredential {
  const rotatedRefreshToken = stringValue(payload.refresh_token);
  if (submittedRefreshToken && (!rotatedRefreshToken || rotatedRefreshToken === submittedRefreshToken)) {
    throw new Error("Nous Portal did not rotate its single-use refresh token. Sign in again rather than replaying the consumed token.");
  }
  return asStored("nous", payload, undefined, { apiBaseUrl: NOUS_INFERENCE_BASE_URL });
}

function toProviderCredential(value: { accessToken: string; accountId?: string; projectId?: string; apiBaseUrl?: string }): ProviderOAuthCredential {
  return {
    accessToken: value.accessToken,
    ...(value.accountId ? { accountId: value.accountId } : {}),
    ...(value.projectId ? { projectId: value.projectId } : {}),
    ...(value.apiBaseUrl ? { apiBaseUrl: value.apiBaseUrl } : {}),
  };
}

function validStored(value: unknown): value is StoredOAuthCredential {
  return Boolean(value && typeof value === "object" && !Array.isArray(value)
    && typeof (value as StoredOAuthCredential).accessToken === "string"
    && typeof (value as StoredOAuthCredential).expiresAt === "number");
}

function sleep(ms: number, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal.aborted) { reject(signal.reason ?? new Error("OAuth flow cancelled")); return; }
    const timer = setTimeout(resolve, ms);
    signal.addEventListener("abort", () => { clearTimeout(timer); reject(signal.reason ?? new Error("OAuth flow cancelled")); }, { once: true });
  });
}

export class ProviderOAuthManager {
  private readonly flows = new Map<string, OAuthFlow>();
  private readonly activeFlowByProvider = new Map<SupportedOAuthProviderId, string>();
  private readonly refreshes = new Map<string, Promise<StoredOAuthCredential>>();
  private readonly localCliRefreshes = new Map<LocalCliProviderId, Promise<LocalCliCredential>>();
  private readonly localCliCache = new Map<LocalCliProviderId, LocalCliCredential>();
  private readonly fetchImpl: typeof fetch;

  constructor(private readonly store: CredentialStore, private readonly env: NodeJS.ProcessEnv = process.env, fetchImpl: typeof fetch = fetch) {
    this.fetchImpl = fetchImpl;
  }

  private googleClientId(): string { return this.env.GOOGLE_ANTIGRAVITY_CLIENT_ID?.trim() || DEFAULT_GOOGLE_CLIENT_ID; }

  private googleClientSecret(): string {
    return this.env.GOOGLE_ANTIGRAVITY_CLIENT_SECRET?.trim() || DEFAULT_GOOGLE_CLIENT_SECRET;
  }

  credentialRef(provider: ProviderConfig): string {
    return provider.auth.oauthCredentialRef ?? `oauth:${provider.auth.oauthProvider ?? provider.id}`;
  }

  async status(provider: ProviderConfig): Promise<OAuthStatus> {
    const providerId = provider.auth.oauthProvider;
    if (!oauthProviderIsSupported(providerId)) return { status: "unsupported", detail: "This OAuth provider requires a custom transport that OpenWordCode has not enabled yet.", credentialConfigured: false };
    if (localCliProviderIsSupported(providerId) && await this.isLocalCliSelected(providerId)) {
      const source = localCliSource(providerId);
      const persisted = await readLocalCliCredential(providerId, this.env);
      if (persisted && isLocalCliCredentialLive(persisted)) this.localCliCache.set(providerId, persisted);
      const cached = this.localCliCache.get(providerId);
      const credential = persisted && isLocalCliCredentialLive(persisted) ? persisted : cached && isLocalCliCredentialLive(cached) ? cached : persisted ?? cached;
      if (!credential) return { status: "login-required", detail: `Sign in with ${localCliDisplayName(providerId)}, then connect the session here.`, credentialConfigured: false, ...(source ? { source } : {}) };
      if (!isLocalCliCredentialLive(credential)) {
        if (credential.refreshToken) return { status: "connected", detail: `${localCliDisplayName(providerId)} session will renew automatically when used.`, credentialConfigured: true, ...(source ? { source } : {}) };
        return { status: "expired", detail: `${localCliDisplayName(providerId)} session expired. Sign in again with the CLI, then reconnect it here.`, credentialConfigured: true, ...(source ? { source } : {}) };
      }
      const detail = credential.email ? `Using your ${localCliDisplayName(providerId)} session (${credential.email})` : `Using your ${localCliDisplayName(providerId)} session`;
      return { status: "connected", detail, credentialConfigured: true, ...(source ? { source } : {}) };
    }
    const stored = await this.read(provider.auth.oauthCredentialRef ?? `oauth:${providerId}`);
    if (!stored) return { status: "login-required", detail: `Sign in with your ${provider.displayName} account.`, credentialConfigured: false };
    if (stored.expiresAt <= Date.now()) {
      if (stored.refreshToken) return { status: "connected", detail: stored.email ? `Connected as ${stored.email}; the session will renew automatically.` : "OAuth session will renew automatically when used.", credentialConfigured: true };
      return { status: "expired", detail: stored.email ? `${stored.email} needs to sign in again.` : "The saved OAuth session has expired.", credentialConfigured: true };
    }
    return { status: "connected", detail: stored.email ? `Connected as ${stored.email}` : "OAuth account connected", credentialConfigured: true };
  }

  async resolve(provider: ProviderConfig): Promise<ProviderOAuthCredential | null> {
    const providerId = provider.auth.oauthProvider;
    if (!oauthProviderIsSupported(providerId)) return null;
    if (localCliProviderIsSupported(providerId) && await this.isLocalCliSelected(providerId)) {
      let credential = await this.readUsableLocalCliCredential(providerId);
      if (credential && providerId === "google-antigravity") credential = await ensureAntigravityProject(credential, this.fetchImpl);
      return credential ? toProviderCredential(credential) : null;
    }
    const ref = provider.auth.oauthCredentialRef ?? `oauth:${providerId}`;
    const stored = await this.read(ref);
    if (!stored) return null;
    if (stored.expiresAt > Date.now() + 30_000) return toProviderCredential(stored);
    if (!stored.refreshToken) return null;
    const inFlight = this.refreshes.get(ref);
    if (inFlight) return toProviderCredential(await inFlight);
    const refresh = this.refreshStored(providerId, stored, ref);
    this.refreshes.set(ref, refresh);
    try { return toProviderCredential(await refresh); } finally { this.refreshes.delete(ref); }
  }

  async disconnect(provider: ProviderConfig): Promise<void> {
    const providerId = provider.auth.oauthProvider;
    if (!providerId) return;
    await this.store.remove(provider.auth.oauthCredentialRef ?? `oauth:${providerId}`);
    if (localCliProviderIsSupported(providerId)) {
      await this.store.remove(this.localCliSourceRef(providerId));
      this.localCliCache.delete(providerId);
    }
    for (const flow of this.flows.values()) if (flow.providerId === providerId) this.cancel(flow.id);
  }

  async startLocalCliLogin(providerId: string): Promise<LocalCliLoginLaunch> {
    if (!localCliProviderIsSupported(providerId)) throw new Error(`${providerId} does not have a supported local CLI connector`);
    const source = localCliSource(providerId);
    if (!source) throw new Error(`${providerId} does not have a supported local CLI connector`);
    const launch = await launchLocalCliLogin(providerId, this.env);
    await this.store.set(this.localCliSourceRef(providerId), source);
    return launch;
  }

  async useLocalCli(providerId: string): Promise<void> {
    if (!localCliProviderIsSupported(providerId)) throw new Error(`${providerId} does not have a supported local CLI connector`);
    const source = localCliSource(providerId);
    if (!source) throw new Error(`${providerId} does not have a supported local CLI connector`);
    const credential = await this.readUsableLocalCliCredential(providerId);
    if (!credential || !isLocalCliCredentialLive(credential)) throw new Error(`No active ${localCliDisplayName(providerId)} session was found. Sign in with the CLI, then connect the session here.`);
    await this.store.set(this.localCliSourceRef(providerId), source);
  }

  async clearLocalCli(providerId: string): Promise<void> {
    if (!localCliProviderIsSupported(providerId)) return;
    await this.store.remove(this.localCliSourceRef(providerId));
    this.localCliCache.delete(providerId);
  }

  async start(providerId: string): Promise<OAuthFlowSnapshot> {
    if (!oauthProviderIsSupported(providerId)) throw new Error(`OAuth for ${providerId} requires a provider-specific transport and is not enabled in this build`);
    const activeFlowId = this.activeFlowByProvider.get(providerId);
    const activeFlow = activeFlowId ? this.flows.get(activeFlowId) : undefined;
    if (activeFlow?.status === "pending" && Date.parse(activeFlow.expiresAt) > Date.now()) return this.snapshot(activeFlow);
    if (activeFlow?.status === "pending") this.fail(activeFlow, "OAuth flow timed out");
    if (localCliProviderIsSupported(providerId)) await this.store.remove(this.localCliSourceRef(providerId));
    const flow: OAuthFlow = {
      id: `oauth_${randomUUID()}`,
      providerId,
      status: "pending",
      detail: "Preparing sign-in…",
      createdAt: new Date().toISOString(),
      expiresAt: new Date(Date.now() + 15 * 60_000).toISOString(),
      controller: new AbortController(),
    };
    this.flows.set(flow.id, flow);
    this.activeFlowByProvider.set(providerId, flow.id);
    try {
      if (providerId === "kimi" || providerId === "github-copilot" || providerId === "nous") return await this.startDevice(flow);
      return await this.startBrowser(flow);
    } catch (error) {
      this.fail(flow, error instanceof Error ? error.message : "Could not start OAuth");
      throw error;
    }
  }

  flowStatus(flowId: string): OAuthFlowSnapshot {
    const flow = this.flows.get(flowId);
    if (!flow) throw new Error("OAuth flow not found or expired");
    if (flow.status === "pending" && Date.parse(flow.expiresAt) <= Date.now()) this.fail(flow, "OAuth flow timed out");
    return this.snapshot(flow);
  }

  async completeManualCode(flowId: string, code: string): Promise<OAuthFlowSnapshot> {
    const flow = this.flows.get(flowId);
    if (!flow) throw new Error("OAuth flow not found or expired");
    if (flow.providerId !== "xai") throw new Error("Manual authorization codes are only supported for xAI sign-in");
    if (flow.status !== "pending") return this.snapshot(flow);
    const value = code.trim();
    if (!value) throw new Error("Paste the authorization code from the xAI sign-in page");
    try {
      const credential = await this.exchangeBrowser(flow, value);
      await this.saveFlowCredential(flow, credential);
      flow.status = "connected";
      flow.detail = credential.email ? `Connected as ${credential.email}` : `${this.displayName(flow.providerId)} account connected`;
      this.closeServer(flow);
      return this.snapshot(flow);
    } catch (error) {
      this.fail(flow, error instanceof Error ? error.message : "xAI authorization-code exchange failed");
      throw error;
    }
  }

  cancel(flowId: string): void {
    const flow = this.flows.get(flowId);
    if (!flow || flow.status !== "pending") return;
    flow.controller.abort(new Error("OAuth flow cancelled"));
    flow.status = "cancelled";
    flow.detail = "Sign-in cancelled";
    this.closeServer(flow);
  }

  private snapshot(flow: OAuthFlow): OAuthFlowSnapshot {
    return {
      flowId: flow.id,
      providerId: flow.providerId,
      status: flow.status,
      detail: flow.detail,
      createdAt: flow.createdAt,
      expiresAt: flow.expiresAt,
      ...(flow.authorizeUrl ? { authorizeUrl: flow.authorizeUrl } : {}),
      ...(flow.verificationUrl ? { verificationUrl: flow.verificationUrl } : {}),
      ...(flow.userCode ? { userCode: flow.userCode } : {}),
    };
  }

  private async read(ref: string): Promise<StoredOAuthCredential | null> {
    const raw = await this.store.get(ref);
    if (!raw) return null;
    try {
      const parsed: unknown = JSON.parse(raw);
      return validStored(parsed) ? parsed : null;
    } catch { return null; }
  }

  private async save(ref: string, credential: StoredOAuthCredential): Promise<void> {
    await this.store.set(ref, JSON.stringify(credential));
  }

  private assertCurrentFlow(flow: OAuthFlow): void {
    if (flow.status !== "pending"
      || flow.controller.signal.aborted
      || this.activeFlowByProvider.get(flow.providerId) !== flow.id) {
      throw new Error("OAuth sign-in was superseded by a newer attempt");
    }
  }

  private async saveFlowCredential(flow: OAuthFlow, credential: StoredOAuthCredential): Promise<void> {
    await this.store.set(`oauth:${flow.providerId}`, JSON.stringify(credential), {
      assertBeforePersist: () => this.assertCurrentFlow(flow),
    });
  }

  private localCliSourceRef(providerId: LocalCliProviderId): string {
    return `provider:local-cli:${providerId}`;
  }

  private async isLocalCliSelected(providerId: LocalCliProviderId): Promise<boolean> {
    return (await this.store.get(this.localCliSourceRef(providerId))) === localCliSource(providerId);
  }

  private async readUsableLocalCliCredential(providerId: LocalCliProviderId): Promise<LocalCliCredential | null> {
    const persisted = await readLocalCliCredential(providerId, this.env);
    if (persisted && isLocalCliCredentialLive(persisted)) {
      this.localCliCache.set(providerId, persisted);
      return persisted;
    }
    const cached = this.localCliCache.get(providerId);
    if (cached && isLocalCliCredentialLive(cached)) return cached;
    const previous = persisted ?? cached;
    if (!previous?.refreshToken) return previous ?? null;
    const inFlight = this.localCliRefreshes.get(providerId);
    if (inFlight) return inFlight;
    const refresh = refreshLocalCliCredential(providerId, previous, this.env, this.fetchImpl);
    this.localCliRefreshes.set(providerId, refresh);
    try {
      const next = await refresh;
      this.localCliCache.set(providerId, next);
      return next;
    } finally {
      this.localCliRefreshes.delete(providerId);
    }
  }

  private async startBrowser(flow: OAuthFlow): Promise<OAuthFlowSnapshot> {
    const callback = LOOPBACK[flow.providerId as BrowserOAuthProviderId];
    const verifier = randomSecret();
    const state = randomSecret();
    const redirectUri = oauthRedirectUri(flow.providerId as BrowserOAuthProviderId);
    flow.verifier = verifier;
    flow.state = state;
    flow.redirectUri = redirectUri;
    const handler = (request: IncomingMessage, response: ServerResponse): void => { void this.handleBrowserCallback(flow, request, response); };
    flow.servers = await this.listen(handler, callback.port, callback.hostname);
    flow.server = flow.servers[0];
    try {
      const challenge = pkceChallenge(verifier);
      let url: string;
      if (flow.providerId === "anthropic") {
        url = `${ANTHROPIC_AUTHORIZE_URL}?${new URLSearchParams({ code: "true", client_id: ANTHROPIC_CLIENT_ID, response_type: "code", redirect_uri: redirectUri, scope: "org:create_api_key user:profile user:inference", code_challenge: challenge, code_challenge_method: "S256", state }).toString()}`;
      } else if (flow.providerId === "xai") {
        const discovery = await this.discoverXai(flow.controller.signal);
        url = `${discovery.authorize}?${new URLSearchParams({ response_type: "code", client_id: XAI_CLIENT_ID, redirect_uri: redirectUri, scope: XAI_SCOPE, code_challenge: challenge, code_challenge_method: "S256", state, nonce: randomUUID() }).toString()}`;
      } else {
        url = `${GOOGLE_AUTHORIZE_URL}?${new URLSearchParams({ response_type: "code", client_id: this.googleClientId(), redirect_uri: redirectUri, scope: GOOGLE_SCOPES, code_challenge: challenge, code_challenge_method: "S256", access_type: "offline", prompt: "consent", state }).toString()}`;
      }
      flow.authorizeUrl = url;
      flow.detail = `Complete ${this.displayName(flow.providerId)} sign-in in your browser.`;
      return this.snapshot(flow);
    } catch (error) {
      this.closeServer(flow);
      throw error;
    }
  }

  private async listen(handler: (request: IncomingMessage, response: ServerResponse) => void, port: number, hostname: "localhost" | "127.0.0.1"): Promise<Server[]> {
    const hosts = hostname === "localhost" ? ["127.0.0.1", "::1"] : ["127.0.0.1"];
    const servers: Server[] = [];
    try {
      for (const host of hosts) {
        const server = createServer(handler);
        try {
          await new Promise<void>((resolve, reject) => {
            const onError = (error: Error): void => { server.off("listening", onListening); reject(error); };
            const onListening = (): void => { server.off("error", onError); resolve(); };
            server.once("error", onError);
            server.once("listening", onListening);
            server.listen(port, host);
          });
          servers.push(server);
        } catch (error) {
          try { server.close(); } catch {}
          // IPv6 is optional on some developer machines. Keep the IPv4
          // callback working there, but never ignore a port conflict because
          // localhost could otherwise resolve to an unrelated process.
          const code = (error as NodeJS.ErrnoException).code;
          if (host === "::1" && (code === "EAFNOSUPPORT" || code === "ENOTSUP" || code === "EINVAL")) continue;
          throw error;
        }
      }
      return servers;
    } catch {
      for (const server of servers) {
        try { server.close(); } catch {}
      }
      throw new Error(`Could not open the ${port} OAuth callback port. Close any older OpenWordCode sign-in window and try again.`);
    }
  }

  private async handleBrowserCallback(flow: OAuthFlow, request: IncomingMessage, response: ServerResponse): Promise<void> {
    if (request.method !== "GET") { response.statusCode = 405; response.end("Method not allowed"); return; }
    const path = new URL(request.url ?? "/", "http://127.0.0.1").pathname;
    if (path !== "/callback") { response.statusCode = 404; response.end("Not found"); return; }
    if (flow.status !== "pending") { sendHtml(response, false, flow.detail); return; }
    const query = new URL(request.url ?? "/", "http://127.0.0.1").searchParams;
    const returnedState = query.get("state");
    const code = query.get("code");
    const oauthError = query.get("error");
    if (!returnedState || returnedState !== flow.state) {
      this.fail(flow, "OAuth state validation failed. Start sign-in again.");
      sendHtml(response, false, flow.detail);
      return;
    }
    if (oauthError || !code) {
      this.fail(flow, oauthError ? `Provider rejected sign-in: ${oauthError}` : "OAuth callback did not include an authorization code.");
      sendHtml(response, false, flow.detail);
      return;
    }
    try {
      const credential = await this.exchangeBrowser(flow, code);
      await this.saveFlowCredential(flow, credential);
      flow.status = "connected";
      flow.detail = credential.email ? `Connected as ${credential.email}` : `${this.displayName(flow.providerId)} account connected`;
      sendHtml(response, true, flow.detail);
    } catch (error) {
      this.fail(flow, error instanceof Error ? error.message : "OAuth exchange failed");
      sendHtml(response, false, flow.detail);
    } finally {
      this.closeServer(flow);
    }
  }

  private async exchangeBrowser(flow: OAuthFlow, code: string): Promise<StoredOAuthCredential> {
    const redirectUri = flow.redirectUri ?? "";
    const verifier = flow.verifier ?? "";
    if (!verifier) throw new Error("OAuth PKCE verifier is missing");
    if (flow.providerId === "anthropic") {
      let exchangeCode = code;
      let exchangeState = flow.state;
      const hash = code.indexOf("#");
      if (hash >= 0) {
        exchangeCode = code.slice(0, hash);
        const frag = code.slice(hash + 1);
        if (frag.length > 0) exchangeState = frag;
      }
      const payload = await postJson(this.fetchImpl, ANTHROPIC_TOKEN_URL, { grant_type: "authorization_code", client_id: ANTHROPIC_CLIENT_ID, code: exchangeCode, state: exchangeState, redirect_uri: redirectUri, code_verifier: verifier }, flow.controller.signal);
      return asStored(flow.providerId, payload);
    }
    if (flow.providerId === "xai") {
      const discovery = await this.discoverXai(flow.controller.signal);
      const payload = await postForm(this.fetchImpl, discovery.token, { grant_type: "authorization_code", client_id: XAI_CLIENT_ID, code, redirect_uri: redirectUri, code_verifier: verifier }, {}, flow.controller.signal);
      return asStored(flow.providerId, payload);
    }
    const payload = await postForm(this.fetchImpl, GOOGLE_TOKEN_URL, { grant_type: "authorization_code", client_id: this.googleClientId(), client_secret: this.googleClientSecret(), code, redirect_uri: redirectUri, code_verifier: verifier }, {}, flow.controller.signal);
    const base = asStored(flow.providerId, payload);
    const projectId = await this.discoverGoogleProject(base.accessToken, flow.controller.signal);
    if (!projectId) throw new Error("Google sign-in succeeded, but Cloud Code Assist did not return a project for this account.");
    return { ...base, projectId };
  }

  private async startDevice(flow: OAuthFlow): Promise<OAuthFlowSnapshot> {
    const providerId = flow.providerId;
    if (providerId !== "kimi" && providerId !== "github-copilot" && providerId !== "nous") throw new Error("This OAuth provider does not use a device flow");
    const device = await this.requestDevice(providerId, flow.controller.signal);
    flow.device = device;
    flow.verificationUrl = device.verificationUrl;
    flow.userCode = device.userCode;
    flow.detail = `Open the verification page and enter ${device.userCode}.`;
    void this.pollDevice(flow).catch(error => this.fail(flow, error instanceof Error ? error.message : "Device sign-in failed"));
    return this.snapshot(flow);
  }

  private async requestDevice(providerId: "kimi" | "github-copilot" | "nous", signal: AbortSignal): Promise<DeviceFlowData> {
    if (providerId === "kimi") {
      const payload = await postForm(this.fetchImpl, `${KIMI_OAUTH_HOST}/api/oauth/device_authorization`, { client_id: KIMI_CLIENT_ID }, requestHeaders(providerId), signal);
      const deviceCode = stringValue(payload.device_code);
      const userCode = stringValue(payload.user_code);
      const verificationUrl = stringValue(payload.verification_uri_complete) ?? stringValue(payload.verification_uri);
      if (!deviceCode || !userCode || !verificationUrl) throw new Error("Kimi device authorization response was incomplete");
      return { kind: providerId, deviceCode, userCode, verificationUrl, intervalMs: this.interval(payload), deadline: Date.now() + this.ttl(payload) };
    }
    if (providerId === "nous") {
      const payload = await postForm(this.fetchImpl, `${NOUS_BASE_URL}/api/oauth/device/code`, { client_id: NOUS_CLIENT_ID, scope: "inference:invoke" }, {}, signal);
      const deviceCode = stringValue(payload.device_code);
      const userCode = stringValue(payload.user_code);
      const verificationUrl = stringValue(payload.verification_uri_complete) ?? stringValue(payload.verification_uri);
      if (!deviceCode || !userCode || !verificationUrl) throw new Error("Nous device authorization response was incomplete");
      return { kind: providerId, deviceCode, userCode, verificationUrl, intervalMs: this.interval(payload), deadline: Date.now() + this.ttl(payload) };
    }
    const payload = await postForm(this.fetchImpl, GITHUB_DEVICE_URL, { client_id: GITHUB_CLIENT_ID, scope: "read:user" }, { "User-Agent": "openwordcode" }, signal);
    const deviceCode = stringValue(payload.device_code);
    const userCode = stringValue(payload.user_code);
    if (!deviceCode || !userCode) throw new Error("GitHub device authorization response was incomplete");
    return { kind: providerId, deviceCode, userCode, verificationUrl: "https://github.com/login/device", intervalMs: this.interval(payload, 5_000), deadline: Date.now() + this.ttl(payload) };
  }

  private interval(payload: TokenPayload & Record<string, unknown>, fallback = 5_000): number {
    return typeof payload.interval === "number" && payload.interval > 0 ? payload.interval * 1000 : fallback;
  }

  private ttl(payload: TokenPayload & Record<string, unknown>): number {
    return typeof payload.expires_in === "number" && payload.expires_in > 0 ? payload.expires_in * 1000 : 15 * 60_000;
  }

  private async pollDevice(flow: OAuthFlow): Promise<void> {
    const device = flow.device;
    if (!device) throw new Error("Device flow was not initialized");
    let intervalMs = Math.max(1_000, device.intervalMs);
    while (Date.now() < device.deadline && flow.status === "pending") {
      await sleep(intervalMs, flow.controller.signal);
      if (Date.now() >= device.deadline || flow.status !== "pending") break;
      const result = device.kind === "github-copilot"
        ? await this.pollGithub(flow, device)
        : await this.pollStandardDevice(flow, device);
      if (result.credential) {
        await this.saveFlowCredential(flow, result.credential);
        flow.status = "connected";
        flow.detail = result.credential.email ? `Connected as ${result.credential.email}` : `${this.displayName(flow.providerId)} account connected`;
        return;
      }
      if (result.slowDown) intervalMs = nextDevicePollInterval(intervalMs, result.retryAfterMs);
    }
    if (flow.status === "pending") this.fail(flow, "Device sign-in timed out. Start again to receive a new code.");
  }

  private async pollStandardDevice(flow: OAuthFlow, device: DeviceFlowData): Promise<DevicePollResult> {
    const url = device.kind === "kimi" ? `${KIMI_OAUTH_HOST}/api/oauth/token` : `${NOUS_BASE_URL}/api/oauth/token`;
    const values = { client_id: device.kind === "kimi" ? KIMI_CLIENT_ID : NOUS_CLIENT_ID, device_code: device.deviceCode, grant_type: "urn:ietf:params:oauth:grant-type:device_code" };
    const payload = await this.postDevice(url, values, requestHeaders(device.kind), flow.controller.signal);
    const error = stringValue(payload.error);
    if (error === "authorization_pending") return { credential: null, slowDown: false };
    if (error === "slow_down") return { credential: null, slowDown: true, retryAfterMs: this.interval(payload, 0) || undefined };
    if (error === "expired_token" || error === "access_denied") throw new Error(`${this.displayName(flow.providerId)} device sign-in ${error.replace("_", " ")}`);
    if (error) throw new Error(`${this.displayName(flow.providerId)} device sign-in failed: ${error}`);
    if (!stringValue(payload.access_token)) return { credential: null, slowDown: false };
    return {
      credential: device.kind === "nous" ? asNousStored(payload) : asStored(flow.providerId, payload),
      slowDown: false,
    };
  }

  private async pollGithub(flow: OAuthFlow, device: DeviceFlowData): Promise<DevicePollResult> {
    const payload = await this.postDevice(GITHUB_TOKEN_URL, { client_id: GITHUB_CLIENT_ID, device_code: device.deviceCode, grant_type: "urn:ietf:params:oauth:grant-type:device_code" }, { "User-Agent": "openwordcode" }, flow.controller.signal);
    const error = stringValue(payload.error);
    if (error === "authorization_pending") return { credential: null, slowDown: false };
    if (error === "slow_down") return { credential: null, slowDown: true, retryAfterMs: this.interval(payload, 0) || undefined };
    if (error === "expired_token" || error === "access_denied") throw new Error(`GitHub device sign-in ${error.replace("_", " ")}`);
    if (error) throw new Error(`GitHub device sign-in failed: ${error}`);
    const githubToken = stringValue(payload.access_token);
    if (!githubToken) return { credential: null, slowDown: false };
    const exchanged = await this.exchangeCopilot(githubToken, flow.controller.signal);
    return { credential: { ...exchanged, refreshToken: stringValue(payload.refresh_token) ?? githubToken }, slowDown: false };
  }

  private async postDevice(url: string, values: Record<string, string>, headers: Record<string, string>, signal: AbortSignal): Promise<TokenPayload & Record<string, unknown>> {
    const response = await this.fetchImpl(url, { method: "POST", headers: { Accept: "application/json", "Content-Type": "application/x-www-form-urlencoded", ...headers }, body: new URLSearchParams(values).toString(), redirect: "error", signal: requestSignal(signal) });
    const payload = await readJson(response);
    if (!response.ok && !stringValue(payload.error)) throw new Error(`OAuth device poll failed (${response.status})`);
    return payload;
  }

  private async exchangeCopilot(githubToken: string, signal: AbortSignal): Promise<StoredOAuthCredential> {
    const response = await this.fetchImpl(GITHUB_COPILOT_TOKEN_URL, { headers: { ...requestHeaders("github-copilot"), Accept: "application/json", Authorization: `token ${githubToken}` }, redirect: "error", signal: requestSignal(signal) });
    const payload = await readJson(response);
    if (!response.ok) throw new Error(`GitHub Copilot token exchange failed (${response.status})`);
    const copilotToken = stringValue(payload.token);
    if (!copilotToken) throw new Error("GitHub Copilot token exchange did not return a token");
    const identityResponse = await this.fetchImpl(GITHUB_USER_URL, { headers: { Accept: "application/vnd.github+json", Authorization: `Bearer ${githubToken}`, "User-Agent": "openwordcode", "X-GitHub-Api-Version": "2022-11-28" }, redirect: "error", signal: requestSignal(signal) });
    const identity = identityResponse.ok ? await readJson(identityResponse) : {};
    const accountId = typeof identity.id === "number" ? String(identity.id) : stringValue(identity.login);
    const email = stringValue(identity.email);
    const apiBaseUrl = this.validateCopilotBase(stringValue(payload.endpoints && typeof payload.endpoints === "object" ? payload.endpoints.api : undefined));
    return { accessToken: copilotToken, expiresAt: expiresAt(payload, 25 * 60 * 1000), apiBaseUrl: apiBaseUrl ?? GITHUB_COPILOT_BASE_URL, ...(accountId ? { accountId } : {}), ...(email ? { email } : {}) };
  }

  private validateCopilotBase(value: string | undefined): string | undefined {
    if (!value) return undefined;
    try {
      const parsed = new URL(value);
      const host = parsed.hostname.toLowerCase();
      if (parsed.protocol !== "https:" || parsed.username || parsed.password || (parsed.port && parsed.port !== "443")) return undefined;
      if (host !== "api.githubcopilot.com" && !host.endsWith(".githubcopilot.com")) return undefined;
      return `https://${host}`;
    } catch { return undefined; }
  }

  private async discoverXai(signal: AbortSignal): Promise<{ authorize: string; token: string }> {
    const response = await this.fetchImpl(XAI_DISCOVERY_URL, { headers: { Accept: "application/json" }, redirect: "error", signal: requestSignal(signal) });
    const payload = await readJson(response);
    const authorize = stringValue(payload.authorization_endpoint);
    const token = stringValue(payload.token_endpoint);
    if (!response.ok || !authorize || !token) throw new Error("xAI OAuth discovery did not return valid endpoints");
    return { authorize: validateXaiOAuthEndpoint(authorize), token: validateXaiOAuthEndpoint(token) };
  }

  private async discoverGoogleProject(accessToken: string, signal: AbortSignal): Promise<string | undefined> {
    const headers = { Authorization: `Bearer ${accessToken}`, Accept: "*/*", "Content-Type": "application/json", "User-Agent": antigravityUserAgent(this.env) };
    const first = await this.fetchImpl(`${GOOGLE_CLOUD_CODE_API}/${GOOGLE_API_VERSION}:loadCodeAssist`, { method: "POST", headers, body: JSON.stringify({ metadata: { ideType: "ANTIGRAVITY" } }), redirect: "error", signal: requestSignal(signal) });
    const firstPayload = await readJson(first);
    const found = this.extractProject(firstPayload);
    if (found) return found;
    for (let attempt = 0; attempt < 3; attempt++) {
      const response = await this.fetchImpl(`${GOOGLE_CLOUD_CODE_DAILY_API}/${GOOGLE_API_VERSION}:onboardUser`, { method: "POST", headers, body: JSON.stringify({ tier_id: "free-tier", metadata: { ide_type: "ANTIGRAVITY", ide_name: "antigravity", ide_version: ANTIGRAVITY_IDE_VERSION } }), redirect: "error", signal: requestSignal(signal) });
      const payload = await readJson(response);
      const project = this.extractProject(payload.response && typeof payload.response === "object" ? payload.response as Record<string, unknown> : payload);
      if (project && (payload.done === true || response.ok)) return project;
      await sleep(500, signal);
    }
    return undefined;
  }

  private extractProject(payload: Record<string, unknown>): string | undefined {
    for (const key of ["cloudaicompanionProject", "projectId", "project"]) {
      const value = payload[key];
      if (typeof value === "string" && value.trim()) return value.trim();
      if (value && typeof value === "object" && typeof (value as { id?: unknown }).id === "string") return (value as { id: string }).id;
    }
    return undefined;
  }

  private async refreshStored(providerId: SupportedOAuthProviderId, previous: StoredOAuthCredential, ref: string): Promise<StoredOAuthCredential> {
    const refreshToken = previous.refreshToken;
    if (!refreshToken) throw new Error(`${this.displayName(providerId)} OAuth session expired; sign in again`);
    const signal = AbortSignal.timeout(30_000);
    let refreshed: StoredOAuthCredential;
    if (providerId === "anthropic") {
      const payload = await postJson(this.fetchImpl, ANTHROPIC_TOKEN_URL, { grant_type: "refresh_token", client_id: ANTHROPIC_CLIENT_ID, refresh_token: refreshToken }, signal);
      refreshed = asStored(providerId, payload, refreshToken);
    } else if (providerId === "xai") {
      const discovery = await this.discoverXai(signal);
      const payload = await postForm(this.fetchImpl, discovery.token, { grant_type: "refresh_token", client_id: XAI_CLIENT_ID, refresh_token: refreshToken }, {}, signal);
      refreshed = asStored(providerId, payload, refreshToken);
    } else if (providerId === "google-antigravity") {
      const payload = await postForm(this.fetchImpl, GOOGLE_TOKEN_URL, { grant_type: "refresh_token", client_id: this.googleClientId(), client_secret: this.googleClientSecret(), refresh_token: refreshToken }, {}, signal);
      const next = asStored(providerId, payload, refreshToken);
      refreshed = { ...next, projectId: await this.discoverGoogleProject(next.accessToken, signal) ?? previous.projectId };
    } else if (providerId === "kimi") {
      const payload = await postForm(this.fetchImpl, `${KIMI_OAUTH_HOST}/api/oauth/token`, { grant_type: "refresh_token", refresh_token: refreshToken, client_id: KIMI_CLIENT_ID }, requestHeaders(providerId), signal);
      refreshed = asStored(providerId, payload, refreshToken);
    } else if (providerId === "nous") {
      const response = await this.fetchImpl(`${NOUS_BASE_URL}/api/oauth/token`, { method: "POST", headers: { Accept: "application/json", "Content-Type": "application/x-www-form-urlencoded", "x-nous-refresh-token": refreshToken }, body: new URLSearchParams({ grant_type: "refresh_token", client_id: NOUS_CLIENT_ID }).toString(), signal });
      const payload = await readJson(response);
      if (!response.ok) throw new Error(`Nous OAuth refresh failed (${response.status})`);
      refreshed = asNousStored(payload, refreshToken);
    } else {
      let githubToken = refreshToken;
      if (refreshToken.startsWith("ghr_")) {
        const payload = await postForm(this.fetchImpl, GITHUB_TOKEN_URL, { client_id: GITHUB_CLIENT_ID, grant_type: "refresh_token", refresh_token: refreshToken }, { "User-Agent": "openwordcode" }, signal);
        githubToken = stringValue(payload.access_token) ?? "";
        if (!githubToken) throw new Error("GitHub OAuth refresh did not return an access token");
        refreshed = await this.exchangeCopilot(githubToken, signal);
        refreshed.refreshToken = stringValue(payload.refresh_token) ?? refreshToken;
      } else {
        refreshed = await this.exchangeCopilot(githubToken, signal);
      }
    }
    await this.save(ref, refreshed);
    return refreshed;
  }

  private fail(flow: OAuthFlow, detail: string): void {
    if (flow.status === "connected") return;
    flow.status = "error";
    flow.detail = detail;
    this.closeServer(flow);
  }

  private closeServer(flow: OAuthFlow): void {
    const servers = flow.servers ?? (flow.server ? [flow.server] : []);
    flow.server = undefined;
    flow.servers = undefined;
    for (const server of servers) {
      try { server.close(); } catch {}
    }
  }

  private displayName(providerId: SupportedOAuthProviderId): string {
    return OAUTH_PROVIDER_CATALOG.find(entry => entry.id === providerId)?.displayName ?? providerId;
  }

  /** Provider-specific headers used by Anthropic's Claude Code-compatible OAuth request. */
  oauthHeaders(provider: ProviderConfig, token: string): Record<string, string> {
    if (provider.auth.oauthProvider !== "anthropic") return {};
    return {
      "anthropic-beta": ANTHROPIC_OAUTH_BETA,
      "X-App": "cli",
      "X-Stainless-Retry-Count": "0",
      "X-Stainless-Runtime": "node",
      "X-Stainless-Lang": "js",
      "X-Stainless-Timeout": "600",
      "X-Stainless-Package-Version": "0.74.0",
      "X-Claude-Code-Session-Id": createHash("sha256").update(`claude-code-session:${token}`).digest("hex").slice(0, 32),
      "x-client-request-id": randomUUID(),
    };
  }

  oauthSystemInstruction(provider: ProviderConfig): string | undefined {
    return provider.auth.oauthProvider === "anthropic" ? ANTHROPIC_SYSTEM_IDENTITY : undefined;
  }
}
