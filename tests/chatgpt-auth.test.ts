import { describe, expect, it, vi } from "vitest";
import { MemoryCredentialStore } from "../packages/auth/src/index.js";
import { ChatGPTOAuthManager } from "../apps/core/src/chatgpt-auth.js";

describe("ChatGPT OAuth manager", () => {
  it("creates a PKCE flow, exchanges the callback code, and keeps tokens out of status", async () => {
    const store = new MemoryCredentialStore();
    const manager = new ChatGPTOAuthManager(store, {
      OPENWORDCODE_OPENAI_OAUTH_CLIENT_ID: "openwordcode-client",
      OPENWORDCODE_OPENAI_OAUTH_REDIRECT_URI: "http://localhost:10200/oauth/chatgpt/callback",
    });
    const start = manager.start();
    const query = new URL(start.authorizeUrl).searchParams;
    expect(query.get("client_id")).toBe("openwordcode-client");
    expect(query.get("code_challenge_method")).toBe("S256");
    expect(query.get("state")).toBeTruthy();

    const payload = Buffer.from(JSON.stringify({ email: "user@example.com", chatgpt_account_id: "acct-test" })).toString("base64url");
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({ access_token: `header.${payload}.signature`, refresh_token: "refresh-token", expires_in: 3_600 }), { headers: { "content-type": "application/json" } })));
    try {
      const callback = await manager.handleCallback({ code: "authorization-code", state: query.get("state") ?? undefined });
      expect(callback.status).toBe("connected");
      expect((await manager.status()).detail).toContain("user@example.com");
      expect((await manager.resolve())).toEqual({ accessToken: `header.${payload}.signature`, accountId: "acct-test" });
      const stored = await store.get("provider:openwordcode-account");
      expect(stored).not.toContain("authorization-code");
      expect(stored).toContain("refresh-token");
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
