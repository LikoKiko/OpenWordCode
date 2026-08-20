import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { loadSettings } from "../apps/core/src/config.js";

const created: string[] = [];
afterEach(() => { for (const path of created.splice(0)) rmSync(path, { recursive: true, force: true }); });

describe("configuration migration", () => {
  it("defaults to the OpenWordCode Bridge account setup", () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-defaults-test-"));
    created.push(directory);
    const settings = loadSettings({ ...process.env, OPENWORDCODE_DATA_DIR: directory });
    expect(settings.selectedProviderId).toBe("openwordcode-bridge");
    expect(settings.selectedModelId).toBe("gpt-5.6-luna");
    expect(settings.providers.openai).toBeUndefined();
    expect(settings.providers.openrouter).toBeUndefined();
    expect(settings.providers.google).toBeUndefined();
    expect(settings.providers.anthropic?.auth).toMatchObject({ method: "oauth", oauthProvider: "anthropic" });
    expect(settings.providers.anthropic?.defaultModel).toBe("claude-sonnet-5");
    expect(settings.providers["google-antigravity"]?.defaultModel).toBe("gemini-3.7-flash");
    expect(settings.providers.xai?.auth).toMatchObject({ method: "oauth", oauthProvider: "xai" });
    expect(settings.providers.xai?.defaultModel).toBe("grok-4.5");
  });

  it("routes the default and persisted Bridge provider through the configured port", () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-bridge-port-test-"));
    created.push(directory);
    const env = { ...process.env, OPENWORDCODE_DATA_DIR: directory, OPENWORDCODE_BRIDGE_PORT: "11101" };

    const fresh = loadSettings(env);
    expect(fresh.providers["openwordcode-bridge"]?.baseUrl).toBe("http://127.0.0.1:11101/v1");

    writeFileSync(join(directory, "config.json"), JSON.stringify({
      ...fresh,
      providers: {
        ...fresh.providers,
        "openwordcode-bridge": {
          ...fresh.providers["openwordcode-bridge"],
          baseUrl: "http://127.0.0.1:10101/v1",
        },
      },
    }), "utf8");

    const migrated = loadSettings(env);
    expect(migrated.providers["openwordcode-bridge"]?.baseUrl).toBe("http://127.0.0.1:11101/v1");
  });

  it("moves old legacy bridge entries to the first-party Bridge", () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-config-test-"));
    created.push(directory);
    writeFileSync(join(directory, "config.json"), JSON.stringify({
      version: 1,
      selectedProviderId: "legacy-bridge",
      selectedModelId: "gpt-5.6-luna",
      mode: "manual",
      theme: "system",
      providers: {
        "legacy-bridge": {
          id: "legacy-bridge",
          displayName: "Legacy Bridge",
          kind: "openai-compatible",
          baseUrl: "http://127.0.0.1:10100/v1",
          enabled: true,
          local: true,
          auth: { method: "none" },
          defaultModel: "gpt-5.6-luna",
          privacyNote: "legacy",
        },
      },
      allowedOrigins: ["http://localhost:10200"],
    }), "utf8");

    const settings = loadSettings({ ...process.env, OPENWORDCODE_DATA_DIR: directory });
    expect(settings.selectedProviderId).toBe("openwordcode-bridge");
    expect(settings.selectedModelId).toBe("gpt-5.6-luna");
    expect(settings.providers["legacy-bridge"]).toBeUndefined();
    expect(settings.providers["openwordcode-bridge"]).toMatchObject({ kind: "openwordcode-bridge", baseUrl: "http://127.0.0.1:10101/v1" });
  });

  it("updates shipped provider defaults without changing unrelated provider settings", () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-model-defaults-test-"));
    created.push(directory);
    writeFileSync(join(directory, "config.json"), JSON.stringify({
      version: 1,
      selectedProviderId: "xai",
      selectedModelId: "grok-4-1-fast-reasoning",
      mode: "manual",
      theme: "system",
      providers: {
        xai: {
          id: "xai",
          displayName: "xAI",
          kind: "openai-compatible",
          baseUrl: "https://api.x.ai/v1",
          enabled: true,
          local: false,
          auth: { method: "environment", envVar: "XAI_API_KEY", oauthProvider: "xai", oauthCredentialRef: "oauth:xai" },
          defaultModel: "grok-4-1-fast-reasoning",
          privacyNote: "test",
        },
      },
      allowedOrigins: ["http://localhost:10200"],
    }), "utf8");
    const settings = loadSettings({ ...process.env, OPENWORDCODE_DATA_DIR: directory });
    expect(settings.selectedProviderId).toBe("xai");
    expect(settings.selectedModelId).toBe("grok-4.5");
    expect(settings.providers.xai?.defaultModel).toBe("grok-4.5");
    expect(settings.providers.xai?.baseUrl).toBe("https://api.x.ai/v1");
    expect(settings.providers.xai?.auth).toEqual({ method: "oauth", oauthProvider: "xai", oauthCredentialRef: "oauth:xai" });
  });

  it("removes shipped API-only providers from persisted account lists", () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-account-only-test-"));
    created.push(directory);
    writeFileSync(join(directory, "config.json"), JSON.stringify({
      version: 1,
      selectedProviderId: "google",
      selectedModelId: "gemini-3.5-flash",
      mode: "manual",
      theme: "system",
      providers: {
        openai: { id: "openai", displayName: "OpenAI", kind: "openai-compatible", baseUrl: "https://api.openai.com/v1", enabled: true, local: false, auth: { method: "api-key", credentialRef: "provider:openai" }, privacyNote: "legacy" },
        openrouter: { id: "openrouter", displayName: "OpenRouter", kind: "openai-compatible", baseUrl: "https://openrouter.ai/api/v1", enabled: true, local: false, auth: { method: "environment", envVar: "OPENROUTER_API_KEY" }, privacyNote: "legacy" },
        google: { id: "google", displayName: "Google", kind: "gemini", baseUrl: "https://generativelanguage.googleapis.com/v1beta", enabled: true, local: false, auth: { method: "environment", envVar: "GEMINI_API_KEY" }, privacyNote: "legacy" },
      },
      allowedOrigins: ["http://localhost:10200"],
    }), "utf8");

    const settings = loadSettings({ ...process.env, OPENWORDCODE_DATA_DIR: directory });
    expect(settings.selectedProviderId).toBe("openwordcode-bridge");
    expect(settings.selectedModelId).toBe("gpt-5.6-luna");
    expect(settings.providers.openai).toBeUndefined();
    expect(settings.providers.openrouter).toBeUndefined();
    expect(settings.providers.google).toBeUndefined();
  });

  it("migrates retired Antigravity wire ids back to stable picker ids", () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-antigravity-model-test-"));
    created.push(directory);
    writeFileSync(join(directory, "config.json"), JSON.stringify({
      version: 1,
      selectedProviderId: "google-antigravity",
      selectedModelId: "gemini-3.6-flash-high",
      mode: "manual",
      theme: "system",
      providers: {
        "google-antigravity": {
          id: "google-antigravity",
          displayName: "Google Antigravity",
          kind: "google-antigravity",
          baseUrl: "https://daily-cloudcode-pa.googleapis.com",
          enabled: true,
          local: false,
          auth: { method: "oauth", oauthProvider: "google-antigravity", oauthCredentialRef: "oauth:google-antigravity" },
          defaultModel: "gemini-pro-agent",
          privacyNote: "test",
        },
      },
      allowedOrigins: ["http://localhost:10200"],
    }), "utf8");

    const settings = loadSettings({ ...process.env, OPENWORDCODE_DATA_DIR: directory });
    expect(settings.selectedModelId).toBe("gemini-3.7-flash");
    expect(settings.providers["google-antigravity"]?.defaultModel).toBe("gemini-3.1-pro");
  });
});
