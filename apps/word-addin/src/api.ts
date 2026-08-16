import type { AgentAction, AgentEvent, AgentRequest, ModelInfo, ProposedChange, ProviderSummary, DocumentSnapshot } from "../../../packages/shared/src/index.js";

let sessionToken = "";

function errorMessage(text: string, fallback: string): string {
  try {
    const payload = JSON.parse(text) as { message?: unknown; error?: unknown };
    if (typeof payload.message === "string" && payload.message) return payload.message;
    if (typeof payload.error === "string" && payload.error) return payload.error;
  } catch {
    // Keep the original response when it is not JSON.
  }
  return text || fallback;
}

export async function initializeCore(): Promise<{ version: string }> {
  const response = await fetch("/api/bootstrap", { cache: "no-store", headers: { Origin: window.location.origin } });
  if (!response.ok) throw new Error(errorMessage(await response.text(), `Core bootstrap failed (${response.status})`));
  const payload = await response.json() as { sessionToken: string; version: string };
  if (!payload.sessionToken) throw new Error("Core bootstrap returned no session token");
  sessionToken = payload.sessionToken;
  return { version: payload.version };
}

function headers(json = false): HeadersInit {
  return { ...(json ? { "Content-Type": "application/json" } : {}), "x-openwordcode-session": sessionToken, "x-openwordcode-csrf": sessionToken };
}

async function request<T>(path: string, init: RequestInit = {}, retry = true): Promise<T> {
  const response = await fetch(path, { ...init, headers: { ...headers(Boolean(init.body)), ...(init.headers ?? {}) } });
  if (response.status === 401 && retry) {
    sessionToken = "";
    await initializeCore();
    return request<T>(path, init, false);
  }
  if (!response.ok) {
    const text = await response.text();
    throw new Error(errorMessage(text, `Core request failed (${response.status})`));
  }
  return response.json() as Promise<T>;
}

export async function getHealth(): Promise<{ status: string; version: string }> {
  const response = await fetch("/health");
  if (!response.ok) throw new Error("Core is offline");
  return response.json() as Promise<{ status: string; version: string }>;
}

export async function getProviders(): Promise<ProviderSummary[]> {
  const payload = await request<{ providers: ProviderSummary[] }>("/api/providers");
  return payload.providers;
}

export async function getModels(providerId: string): Promise<ModelInfo[]> {
  const payload = await request<{ models: ModelInfo[] }>(`/api/providers/${encodeURIComponent(providerId)}/models`);
  return payload.models;
}

export async function getSettings(): Promise<{ selectedProviderId: string; selectedModelId?: string; mode: AgentRequest["mode"]; theme: "light" | "dark" | "system" }> {
  const payload = await request<{ settings: { selectedProviderId: string; selectedModelId?: string; mode: AgentRequest["mode"]; theme: "light" | "dark" | "system" } }>("/api/settings");
  return payload.settings;
}

export async function saveSettings(patch: Partial<Pick<AgentRequest, "mode">> & { selectedProviderId?: string; selectedModelId?: string; theme?: "light" | "dark" | "system" }): Promise<void> {
  await request("/api/settings", { method: "PUT", body: JSON.stringify(patch) });
}

export interface ChatGPTLoginStart {
  flowId: string;
  authorizeUrl: string;
  redirectUri: string;
  expiresAt: string;
}

export interface ChatGPTLoginStatus {
  flowId: string;
  status: "pending" | "connected" | "error";
  detail: string;
}

export interface CodexCliLoginLaunch {
  ok: true;
  executable: string;
  command: string;
}

export async function startChatGPTLogin(): Promise<ChatGPTLoginStart> {
  return request<ChatGPTLoginStart>("/api/auth/chatgpt/start", { method: "POST", body: "{}" });
}

export async function startCodexCliLogin(): Promise<CodexCliLoginLaunch> {
  return request<CodexCliLoginLaunch>("/api/auth/chatgpt/start-codex-cli-login", { method: "POST", body: "{}" });
}

export async function useCodexCliSession(): Promise<void> {
  await request("/api/auth/chatgpt/use-codex-cli", { method: "POST", body: "{}" });
}

export async function getChatGPTLoginStatus(flowId: string): Promise<ChatGPTLoginStatus> {
  return request<ChatGPTLoginStatus>(`/api/auth/chatgpt/status?flowId=${encodeURIComponent(flowId)}`);
}

export async function cancelChatGPTLogin(flowId: string): Promise<void> {
  await request("/api/auth/chatgpt/cancel", { method: "POST", body: JSON.stringify({ flowId }) });
}

export async function disconnectChatGPT(): Promise<void> {
  await request("/api/auth/chatgpt/disconnect", { method: "POST", body: "{}" });
}

export interface OAuthProviderOption {
  id: string;
  displayName: string;
  supported: boolean;
  flow: "browser" | "device" | "custom";
  detail: string;
}

export interface OAuthLoginStart {
  flowId: string;
  providerId: string;
  status: "pending" | "connected" | "error" | "cancelled";
  detail: string;
  createdAt: string;
  expiresAt: string;
  authorizeUrl?: string;
  verificationUrl?: string;
  userCode?: string;
}

export type OAuthLoginStatus = OAuthLoginStart;

export interface LocalCliLoginLaunch {
  ok: true;
  providerId: string;
  executable: string;
  command: string;
}

export async function getOAuthProviders(): Promise<OAuthProviderOption[]> {
  const payload = await request<{ providers: OAuthProviderOption[] }>("/api/oauth/providers");
  return payload.providers;
}

export async function startOAuthLogin(providerId: string): Promise<OAuthLoginStart> {
  return request<OAuthLoginStart>(`/api/oauth/${encodeURIComponent(providerId)}/start`, { method: "POST", body: "{}" });
}

export async function startLocalCliLogin(providerId: string): Promise<LocalCliLoginLaunch> {
  return request<LocalCliLoginLaunch>(`/api/oauth/${encodeURIComponent(providerId)}/local-cli/start`, { method: "POST", body: "{}" });
}

export async function useLocalCliSession(providerId: string): Promise<void> {
  await request(`/api/oauth/${encodeURIComponent(providerId)}/local-cli/use`, { method: "POST", body: "{}" });
}

export async function getOAuthLoginStatus(flowId: string): Promise<OAuthLoginStatus> {
  return request<OAuthLoginStatus>(`/api/oauth/flows/${encodeURIComponent(flowId)}`);
}

export async function completeOAuthLogin(flowId: string, code: string): Promise<OAuthLoginStatus> {
  return request<OAuthLoginStatus>(`/api/oauth/flows/${encodeURIComponent(flowId)}/complete`, { method: "POST", body: JSON.stringify({ code }) });
}

export async function cancelOAuthLogin(flowId: string): Promise<void> {
  await request("/api/oauth/flows/cancel", { method: "POST", body: JSON.stringify({ flowId }) });
}

export async function disconnectOAuth(providerId: string): Promise<void> {
  await request(`/api/providers/${encodeURIComponent(providerId)}/disconnect`, { method: "POST", body: "{}" });
}

export async function testProvider(providerId: string): Promise<{ ok: boolean; modelCount: number; durationMs: number }> {
  return request(`/api/providers/${encodeURIComponent(providerId)}/test`, { method: "POST", body: "{}" });
}

export async function getBridgeStatus(): Promise<{ available: boolean; endpoint: string; models: number; detail: string }> {
  return request("/api/openwordcode/bridge/status");
}

export async function approveChange(id: string, currentBefore: string): Promise<ProposedChange> {
  const payload = await request<{ change: ProposedChange }>(`/api/changes/${encodeURIComponent(id)}/approve`, { method: "POST", body: JSON.stringify({ currentBefore }) });
  return payload.change;
}

export async function completeChange(id: string, success: boolean, reason?: string): Promise<ProposedChange> {
  const payload = await request<{ change: ProposedChange }>(`/api/changes/${encodeURIComponent(id)}/complete`, { method: "POST", body: JSON.stringify({ success, reason }) });
  return payload.change;
}

export async function approveConsoleAction(id: string): Promise<AgentAction> {
  const payload = await request<{ action: AgentAction }>(`/api/actions/${encodeURIComponent(id)}/approve`, { method: "POST", body: "{}" });
  return payload.action;
}

export async function rejectConsoleAction(id: string): Promise<AgentAction> {
  const payload = await request<{ action: AgentAction }>(`/api/actions/${encodeURIComponent(id)}/reject`, { method: "POST", body: "{}" });
  return payload.action;
}

export async function streamAgent(requestBody: AgentRequest, onEvent: (event: AgentEvent) => void, signal: AbortSignal, retry = true): Promise<void> {
  const response = await fetch("/api/agent", { method: "POST", headers: headers(true), body: JSON.stringify(requestBody), signal });
  if (response.status === 401 && retry) {
    sessionToken = "";
    await initializeCore();
    return streamAgent(requestBody, onEvent, signal, false);
  }
  if (!response.ok) throw new Error(errorMessage(await response.text(), `Core request failed (${response.status})`));
  if (!response.body) throw new Error("Core returned an empty stream");
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  while (true) {
    const chunk = await reader.read();
    if (chunk.done) break;
    buffer += decoder.decode(chunk.value, { stream: true });
    const frames = buffer.split(/\r?\n\r?\n/);
    buffer = frames.pop() ?? "";
    for (const frame of frames) {
      const line = frame.split(/\r?\n/).find(item => item.startsWith("data:"));
      if (!line) continue;
      try { onEvent(JSON.parse(line.slice(5).trim()) as AgentEvent); } catch { /* ignore malformed frame */ }
    }
  }
}

export type { DocumentSnapshot };
