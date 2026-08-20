import { describe, expect, it, vi } from "vitest";
import { MemoryCredentialStore, type CredentialStore, type CredentialWriteOptions } from "../packages/auth/src/index.js";
import type { ProviderConfig } from "../packages/shared/src/index.js";
import {
  nextDevicePollInterval,
  oauthProviderCatalog,
  ProviderOAuthManager,
  validateXaiOAuthEndpoint,
} from "../apps/core/src/provider-oauth.js";

const nousProvider: ProviderConfig = {
  id: "nous",
  displayName: "Nous Portal",
  kind: "openai-compatible",
  baseUrl: "https://inference-api.nousresearch.com/v1",
  enabled: true,
  local: false,
  auth: { method: "oauth", oauthProvider: "nous", oauthCredentialRef: "oauth:nous" },
  defaultModel: "tencent/hy3:free",
  privacyNote: "test",
};

describe("provider OAuth reliability", () => {
  it("publishes every enabled account connector as supported", () => {
    const supported = oauthProviderCatalog().filter(provider => provider.supported).map(provider => provider.id);
    expect(supported).toEqual(["anthropic", "xai", "kimi", "google-antigravity", "github-copilot", "nous"]);
  });

  it("only accepts HTTPS xAI discovery endpoints on x.ai", () => {
    expect(validateXaiOAuthEndpoint("https://accounts.x.ai/oauth2/authorize")).toBe("https://accounts.x.ai/oauth2/authorize");
    expect(() => validateXaiOAuthEndpoint("https://x.ai.evil.example/token")).toThrow("unexpected endpoint");
    expect(() => validateXaiOAuthEndpoint("http://auth.x.ai/token")).toThrow("unexpected endpoint");
  });

  it("implements RFC 8628 slow-down without unbounded polling", () => {
    expect(nextDevicePollInterval(5_000)).toBe(10_000);
    expect(nextDevicePollInterval(5_000, 20_000)).toBe(20_000);
    expect(nextDevicePollInterval(29_000, 60_000)).toBe(30_000);
  });

  it("treats an expired access token with a refresh token as renewable", async () => {
    const store = new MemoryCredentialStore();
    await store.set("oauth:nous", JSON.stringify({ accessToken: "old-access", refreshToken: "refresh-a", expiresAt: Date.now() - 1_000 }));
    const manager = new ProviderOAuthManager(store, {}, vi.fn() as unknown as typeof fetch);
    await expect(manager.status(nousProvider)).resolves.toMatchObject({ status: "connected", credentialConfigured: true });
  });

  it("persists the rotated Nous refresh token", async () => {
    const store = new MemoryCredentialStore();
    await store.set("oauth:nous", JSON.stringify({ accessToken: "old-access", refreshToken: "refresh-a", expiresAt: Date.now() - 1_000 }));
    const fetchImpl = vi.fn(async () => new Response(JSON.stringify({ access_token: "new-access", refresh_token: "refresh-b", expires_in: 3_600 }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    })) as unknown as typeof fetch;
    const manager = new ProviderOAuthManager(store, {}, fetchImpl);

    await expect(manager.resolve(nousProvider)).resolves.toMatchObject({ accessToken: "new-access", apiBaseUrl: "https://inference-api.nousresearch.com/v1" });
    expect(JSON.parse((await store.get("oauth:nous"))!) as Record<string, unknown>).toMatchObject({ refreshToken: "refresh-b" });
  });

  it("never replays a consumed Nous refresh token", async () => {
    const store = new MemoryCredentialStore();
    await store.set("oauth:nous", JSON.stringify({ accessToken: "old-access", refreshToken: "refresh-a", expiresAt: Date.now() - 1_000 }));
    const fetchImpl = vi.fn(async () => new Response(JSON.stringify({ access_token: "new-access", refresh_token: "refresh-a", expires_in: 3_600 }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    })) as unknown as typeof fetch;
    const manager = new ProviderOAuthManager(store, {}, fetchImpl);

    await expect(manager.resolve(nousProvider)).rejects.toThrow("did not rotate");
    expect(JSON.parse((await store.get("oauth:nous"))!) as Record<string, unknown>).toMatchObject({ accessToken: "old-access", refreshToken: "refresh-a" });
  });

  it("does not let a cancelled OAuth flow overwrite its replacement", async () => {
    let releaseBlockedWrite!: () => void;
    let signalBlockedWrite!: () => void;
    const blockedWrite = new Promise<void>(resolve => { signalBlockedWrite = resolve; });
    const writeGate = new Promise<void>(resolve => { releaseBlockedWrite = resolve; });
    class DelayedCredentialStore implements CredentialStore {
      readonly kind = "memory" as const;
      private readonly values = new Map<string, string>();
      private delayNextOAuthWrite = true;
      async get(id: string): Promise<string | null> { return this.values.get(id) ?? null; }
      async set(id: string, secret: string, options: CredentialWriteOptions = {}): Promise<void> {
        if (id === "oauth:xai" && this.delayNextOAuthWrite) {
          this.delayNextOAuthWrite = false;
          signalBlockedWrite();
          await writeGate;
        }
        options.assertBeforePersist?.();
        this.values.set(id, secret);
      }
      async remove(id: string): Promise<void> { this.values.delete(id); }
    }
    const store = new DelayedCredentialStore();
    const fetchImpl = vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
      const url = String(input);
      if (url.includes(".well-known/openid-configuration")) {
        return new Response(JSON.stringify({
          authorization_endpoint: "https://accounts.x.ai/oauth2/authorize",
          token_endpoint: "https://accounts.x.ai/oauth2/token",
        }), { status: 200, headers: { "Content-Type": "application/json" } });
      }
      const body = new URLSearchParams(typeof init?.body === "string" ? init.body : "");
      const code = body.get("code") ?? "unknown";
      return new Response(JSON.stringify({ access_token: `access-${code}`, refresh_token: `refresh-${code}`, expires_in: 3_600 }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }) as unknown as typeof fetch;
    const manager = new ProviderOAuthManager(store, {}, fetchImpl);
    let firstFlowId: string | undefined;
    let secondFlowId: string | undefined;
    try {
      const first = await manager.start("xai");
      firstFlowId = first.flowId;
      const staleCompletion = manager.completeManualCode(first.flowId, "old");
      await blockedWrite;
      manager.cancel(first.flowId);

      const second = await manager.start("xai");
      secondFlowId = second.flowId;
      releaseBlockedWrite();
      await expect(staleCompletion).rejects.toThrow("superseded");
      await expect(manager.completeManualCode(second.flowId, "new")).resolves.toMatchObject({ status: "connected" });

      expect(JSON.parse((await store.get("oauth:xai"))!) as Record<string, unknown>).toMatchObject({ accessToken: "access-new" });
    } finally {
      releaseBlockedWrite();
      if (firstFlowId) manager.cancel(firstFlowId);
      if (secondFlowId) manager.cancel(secondFlowId);
    }
  });
});
