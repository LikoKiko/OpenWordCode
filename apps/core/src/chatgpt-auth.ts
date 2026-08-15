import { createHash, randomBytes, randomUUID } from "node:crypto";
import { type CredentialStore } from "../../../packages/auth/src/index.js";

const DEFAULT_AUTHORIZE_URL = "https://auth.openai.com/oauth/authorize";
const DEFAULT_TOKEN_URL = "https://auth.openai.com/oauth/token";
const DEFAULT_SCOPE = "openid profile email offline_access";
const DEFAULT_PORT = 10_200;
const CREDENTIAL_REF = "provider:openwordcode-account";
const FLOW_TTL_MS = 10 * 60 * 1_000;

export interface OAuthProviderCredential {
  accessToken: string;
  accountId?: string;
}

interface StoredOAuthCredential extends OAuthProviderCredential {
  refreshToken?: string;
  expiresAt?: number;
  email?: string;
  tokenType?: string;
  scope?: string;
}

interface PendingFlow {
  id: string;
  state: string;
  verifier: string;
  createdAt: number;
  status: "pending" | "connected" | "error";
  detail?: string;
}

export type ChatGPTOAuthFlowStatus = "pending" | "connected" | "error";

export interface ChatGPTOAuthStatus {
  status: "login-required" | "connected" | "expired" | "unsupported" | "error";
  detail: string;
  credentialConfigured: boolean;
  configured: boolean;
  email?: string;
}

export interface ChatGPTOAuthStart {
  flowId: string;
  authorizeUrl: string;
  redirectUri: string;
  expiresAt: string;
}

export interface ChatGPTOAuthFlowResult {
  flowId: string;
  status: ChatGPTOAuthFlowStatus;
  detail: string;
}

export class OAuthConfigurationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "OAuthConfigurationError";
  }
}

function base64Url(value: Uint8Array): string {
  return Buffer.from(value).toString("base64url");
}

function createVerifier(): string {
  return base64Url(randomBytes(48));
}

function createChallenge(verifier: string): string {
  return createHash("sha256").update(verifier).digest("base64url");
}

function randomState(): string {
  return base64Url(randomBytes(32));
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function numberValue(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function recordValue(value: unknown): Record<string, unknown> | null {
  return typeof value === "object" && value !== null && !Array.isArray(value) ? value as Record<string, unknown> : null;
}

function jwtClaims(token: string): Record<string, unknown> {
  const payload = token.split(".")[1];
  if (!payload) return {};
  try {
    const parsed = JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) as unknown;
    return recordValue(parsed) ?? {};
  } catch {
    return {};
  }
}

function accountIdFromClaims(claims: Record<string, unknown>): string | undefined {
  const direct = stringValue(claims.chatgpt_account_id) ?? stringValue(claims.account_id);
  if (direct) return direct;
  for (const key of ["https://api.openai.com/auth", "https://api.openai.com/profile"]) {
    const nested = recordValue(claims[key]);
    const accountId = nested ? stringValue(nested.chatgpt_account_id) ?? stringValue(nested.account_id) : undefined;
    if (accountId) return accountId;
  }
  const organizations = Array.isArray(claims.organizations) ? claims.organizations : [];
  const first = recordValue(organizations[0]);
  return first ? stringValue(first.id) : undefined;
}

function emailFromClaims(claims: Record<string, unknown>): string | undefined {
  return stringValue(claims.email);
}

function parseStoredCredential(raw: string | null): StoredOAuthCredential | null {
  if (!raw) return null;
  try {
    const value = recordValue(JSON.parse(raw));
    const accessToken = value ? stringValue(value.accessToken) : undefined;
    if (!accessToken) return null;
    return {
      accessToken,
      ...(value && stringValue(value.refreshToken) ? { refreshToken: stringValue(value.refreshToken) } : {}),
      ...(value && numberValue(value.expiresAt) !== undefined ? { expiresAt: numberValue(value.expiresAt) } : {}),
      ...(value && stringValue(value.accountId) ? { accountId: stringValue(value.accountId) } : {}),
      ...(value && stringValue(value.email) ? { email: stringValue(value.email) } : {}),
      ...(value && stringValue(value.tokenType) ? { tokenType: stringValue(value.tokenType) } : {}),
      ...(value && stringValue(value.scope) ? { scope: stringValue(value.scope) } : {}),
    };
  } catch {
    return null;
  }
}

function safeError(value: unknown): string {
  return value instanceof Error ? value.message : "OAuth request failed";
}

function localRedirectUri(value: string): string {
  let parsed: URL;
  try { parsed = new URL(value); } catch { throw new OAuthConfigurationError("The OpenWordCode OAuth redirect URI is invalid."); }
  if (parsed.protocol !== "http:" || !["localhost", "127.0.0.1", "[::1]"].includes(parsed.hostname)) {
    throw new OAuthConfigurationError("The OpenWordCode OAuth redirect URI must point to localhost.");
  }
  return parsed.toString().replace(/\/$/u, "");
}

/**
 * App-owned PKCE flow for an OAuth provider configured by the product owner.
 *
 * The manager deliberately does not read browser cookies, unrelated application files, or
 * ChatGPT passwords. The provider client id and redirect registration belong
 * to OpenWordCode and are supplied through the Core environment.
 */
export class ChatGPTOAuthManager {
  private readonly flows = new Map<string, PendingFlow>();
  private refreshPromise: Promise<OAuthProviderCredential> | null = null;

  constructor(private readonly store: CredentialStore, private readonly env: NodeJS.ProcessEnv = process.env) {}

  private clientId(): string | undefined {
    return this.env.OPENWORDCODE_OPENAI_OAUTH_CLIENT_ID?.trim()
      || this.env.OPENWORDCODE_CHATGPT_OAUTH_CLIENT_ID?.trim();
  }

  private redirectUri(): string {
    const configured = this.env.OPENWORDCODE_OPENAI_OAUTH_REDIRECT_URI?.trim();
    const port = Number(this.env.OPENWORDCODE_PORT ?? DEFAULT_PORT) || DEFAULT_PORT;
    return localRedirectUri(configured || `http://localhost:${port}/oauth/chatgpt/callback`);
  }

  private authorizeUrl(): string {
    return this.env.OPENWORDCODE_OPENAI_OAUTH_AUTHORIZE_URL?.trim() || DEFAULT_AUTHORIZE_URL;
  }

  private tokenUrl(): string {
    return this.env.OPENWORDCODE_OPENAI_OAUTH_TOKEN_URL?.trim() || DEFAULT_TOKEN_URL;
  }

  private scope(): string {
    return this.env.OPENWORDCODE_OPENAI_OAUTH_SCOPE?.trim() || DEFAULT_SCOPE;
  }

  isConfigured(): boolean {
    return Boolean(this.clientId());
  }

  async status(credentialRef = CREDENTIAL_REF): Promise<ChatGPTOAuthStatus> {
    const configured = this.isConfigured();
    const credential = parseStoredCredential(await this.store.get(credentialRef));
    if (!configured) {
      return {
        status: "unsupported",
        detail: "OpenWordCode account sign-in is not configured for this build. Set an app-owned OAuth client id in the Core environment.",
        credentialConfigured: Boolean(credential),
        configured: false,
        ...(credential?.email ? { email: credential.email } : {}),
      };
    }
    if (!credential) return { status: "login-required", detail: "Sign in to your OpenWordCode account to use the Bridge.", credentialConfigured: false, configured: true };
    if (credential.expiresAt !== undefined && credential.expiresAt <= Date.now() && !credential.refreshToken) {
      return { status: "expired", detail: "The OpenWordCode account sign-in expired. Sign in again to reconnect.", credentialConfigured: true, configured: true, ...(credential.email ? { email: credential.email } : {}) };
    }
    return { status: "connected", detail: credential.email ? `Signed in as ${credential.email}` : "OpenWordCode account connected", credentialConfigured: true, configured: true, ...(credential.email ? { email: credential.email } : {}) };
  }

  start(): ChatGPTOAuthStart {
    const clientId = this.clientId();
    if (!clientId) throw new OAuthConfigurationError("OpenWordCode account sign-in is not configured. Set OPENWORDCODE_OPENAI_OAUTH_CLIENT_ID for the Core process.");
    const redirectUri = this.redirectUri();
    const verifier = createVerifier();
    const state = randomState();
    const flowId = `oauth_${randomUUID()}`;
    const params = new URLSearchParams({
      response_type: "code",
      client_id: clientId,
      redirect_uri: redirectUri,
      scope: this.scope(),
      code_challenge: createChallenge(verifier),
      code_challenge_method: "S256",
      state,
      prompt: "login",
    });
    const flow: PendingFlow = { id: flowId, state, verifier, createdAt: Date.now(), status: "pending" };
    this.pruneFlows();
    this.flows.set(flowId, flow);
    return { flowId, authorizeUrl: `${this.authorizeUrl()}?${params.toString()}`, redirectUri, expiresAt: new Date(flow.createdAt + FLOW_TTL_MS).toISOString() };
  }

  flowStatus(flowId: string): ChatGPTOAuthFlowResult {
    const flow = this.flows.get(flowId);
    if (!flow) return { flowId, status: "error", detail: "The sign-in attempt expired. Start again." };
    if (flow.status === "connected") return { flowId, status: "connected", detail: flow.detail ?? "OpenWordCode account connected." };
    if (flow.status === "error") return { flowId, status: "error", detail: flow.detail ?? "ChatGPT sign-in failed." };
    return { flowId, status: "pending", detail: "Complete OpenWordCode account sign-in in your browser." };
  }

  cancel(flowId: string): void {
    this.flows.delete(flowId);
  }

  async handleCallback(query: { code?: string; state?: string; error?: string; error_description?: string }): Promise<ChatGPTOAuthFlowResult> {
    this.pruneFlows();
    const flow = [...this.flows.values()].find(candidate => candidate.state === query.state);
    if (!flow) return { flowId: "unknown", status: "error", detail: "The sign-in state is invalid or expired." };
    if (query.error) {
      flow.status = "error";
      flow.detail = query.error_description?.trim() || query.error.trim();
      return this.flowStatus(flow.id);
    }
    if (!query.code) {
      flow.status = "error";
      flow.detail = "The account provider did not return an authorization code.";
      return this.flowStatus(flow.id);
    }
    try {
      const credential = await this.exchangeCode(query.code, flow.verifier);
      await this.storeCredential(credential);
      flow.status = "connected";
      flow.detail = credential.email ? `Signed in as ${credential.email}` : "OpenWordCode account connected.";
    } catch (error) {
      flow.status = "error";
      flow.detail = safeError(error);
    }
    return this.flowStatus(flow.id);
  }

  async disconnect(credentialRef = CREDENTIAL_REF): Promise<void> {
    await this.store.remove(credentialRef);
  }

  async resolve(credentialRef = CREDENTIAL_REF): Promise<OAuthProviderCredential | null> {
    const stored = parseStoredCredential(await this.store.get(credentialRef));
    if (!stored) return null;
    if (stored.expiresAt === undefined || stored.expiresAt > Date.now() + 60_000) return { accessToken: stored.accessToken, ...(stored.accountId ? { accountId: stored.accountId } : {}) };
    if (!stored.refreshToken) return null;
    if (this.refreshPromise) return this.refreshPromise;
    this.refreshPromise = this.refresh(stored).finally(() => { this.refreshPromise = null; });
    return this.refreshPromise;
  }

  private async exchangeCode(code: string, verifier: string): Promise<StoredOAuthCredential> {
    const clientId = this.clientId();
    if (!clientId) throw new OAuthConfigurationError("OpenWordCode OAuth client id is missing.");
    const response = await fetch(this.tokenUrl(), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
      body: new URLSearchParams({ grant_type: "authorization_code", client_id: clientId, code, redirect_uri: this.redirectUri(), code_verifier: verifier }).toString(),
    });
    return this.parseTokenResponse(response);
  }

  private async refresh(previous: StoredOAuthCredential): Promise<OAuthProviderCredential> {
    const clientId = this.clientId();
    if (!clientId || !previous.refreshToken) throw new Error("OpenWordCode account sign-in needs to be completed again.");
    const response = await fetch(this.tokenUrl(), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
      body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, refresh_token: previous.refreshToken }).toString(),
    });
    const next = await this.parseTokenResponse(response);
    const merged: StoredOAuthCredential = { ...previous, ...next, refreshToken: next.refreshToken ?? previous.refreshToken, accountId: next.accountId ?? previous.accountId, email: next.email ?? previous.email };
    await this.storeCredential(merged);
    return { accessToken: merged.accessToken, ...(merged.accountId ? { accountId: merged.accountId } : {}) };
  }

  private async parseTokenResponse(response: Response): Promise<StoredOAuthCredential> {
    let payload: unknown = null;
    try { payload = await response.json() as unknown; } catch { /* handled below */ }
    if (!response.ok) {
      const body = recordValue(payload);
      throw new Error(stringValue(body?.error_description) ?? stringValue(body?.error) ?? `OAuth token request failed (${response.status})`);
    }
    const value = recordValue(payload);
    const accessToken = value ? stringValue(value.access_token) : undefined;
    if (!accessToken) throw new Error("The account OAuth provider returned no access token.");
    const claims = jwtClaims(accessToken);
    const expiresIn = value ? numberValue(value.expires_in) : undefined;
    return {
      accessToken,
      ...(value && stringValue(value.refresh_token) ? { refreshToken: stringValue(value.refresh_token) } : {}),
      ...(expiresIn !== undefined ? { expiresAt: Date.now() + expiresIn * 1_000 } : {}),
      ...(value && (stringValue(value.account_id) ?? stringValue(value.chatgpt_account_id) ?? accountIdFromClaims(claims)) ? { accountId: stringValue(value.account_id) ?? stringValue(value.chatgpt_account_id) ?? accountIdFromClaims(claims) } : {}),
      ...(value && (stringValue(value.email) ?? emailFromClaims(claims)) ? { email: stringValue(value.email) ?? emailFromClaims(claims) } : {}),
      ...(value && stringValue(value.token_type) ? { tokenType: stringValue(value.token_type) } : {}),
      ...(value && stringValue(value.scope) ? { scope: stringValue(value.scope) } : {}),
    };
  }

  private async storeCredential(value: StoredOAuthCredential): Promise<void> {
    await this.store.set(CREDENTIAL_REF, JSON.stringify(value));
  }

  private pruneFlows(): void {
    const cutoff = Date.now() - FLOW_TTL_MS;
    for (const [id, flow] of this.flows) if (flow.createdAt < cutoff) this.flows.delete(id);
  }
}

export { CREDENTIAL_REF as OPENWORDCODE_ACCOUNT_CREDENTIAL_REF };
