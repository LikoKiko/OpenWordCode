import { describe, expect, it } from "vitest";
import { join } from "node:path";
import {
  claudeCliCredentialsPath,
  isLocalCliCredentialLive,
  kimiCliCredentialsPath,
  localCliSource,
  parseAntigravityCliCredential,
  parseClaudeCliCredential,
  parseKimiCliCredential,
} from "../apps/core/src/local-cli-auth.js";

describe("local CLI auth adapters", () => {
  it("parses Claude Code's nested OAuth record without requiring an email", () => {
    const credential = parseClaudeCliCredential(JSON.stringify({ claudeAiOauth: { accessToken: "claude-access", refreshToken: "claude-refresh", expiresAt: 1_900_000_000_000 } }));
    expect(credential).toMatchObject({ accessToken: "claude-access", refreshToken: "claude-refresh", expiresAt: 1_900_000_000_000 });
  });

  it("parses Kimi Code's persisted token shape and normalizes seconds", () => {
    const credential = parseKimiCliCredential(JSON.stringify({ access_token: "kimi-access", refresh_token: "kimi-refresh", expires_at: 1_900_000_000 }));
    expect(credential).toMatchObject({ accessToken: "kimi-access", refreshToken: "kimi-refresh", expiresAt: 1_900_000_000_000 });
  });

  it("parses Antigravity keyring JSON and project metadata", () => {
    const credential = parseAntigravityCliCredential(JSON.stringify({ access_token: "google-access", refresh_token: "google-refresh", expiry_date: 1_900_000_000_000, project_id: "project-123" }));
    expect(credential).toMatchObject({ accessToken: "google-access", refreshToken: "google-refresh", projectId: "project-123" });
    expect(isLocalCliCredentialLive(credential!, 1_800_000_000_000)).toBe(true);
  });

  it("uses provider-owned sources and configurable CLI data roots", () => {
    expect(localCliSource("anthropic")).toBe("claude-cli");
    expect(localCliSource("kimi")).toBe("kimi-cli");
    expect(localCliSource("google-antigravity")).toBe("antigravity-cli");
    expect(claudeCliCredentialsPath({ HOME: "/tmp/openwordcode-test" })).toBe(join("/tmp/openwordcode-test", ".claude", ".credentials.json"));
    expect(kimiCliCredentialsPath({ HOME: "/tmp/openwordcode-test", KIMI_SHARE_DIR: "/tmp/kimi" })).toBe(join("/tmp/kimi", "credentials", "kimi-code.json"));
  });
});
