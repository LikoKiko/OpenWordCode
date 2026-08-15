import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { loadSettings } from "../apps/core/src/config.js";

const created: string[] = [];
afterEach(() => { for (const path of created.splice(0)) rmSync(path, { recursive: true, force: true }); });

describe("configuration migration", () => {
  it("defaults to OpenAI API-key setup", () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-defaults-test-"));
    created.push(directory);
    const settings = loadSettings({ ...process.env, OPENWORDCODE_DATA_DIR: directory });
    expect(settings.selectedProviderId).toBe("openai");
    expect(settings.selectedModelId).toBe("gpt-4.1-mini");
    expect(settings.providers.openai?.auth).toMatchObject({ method: "api-key", credentialRef: "provider:openai" });
  });

  it("moves the old bridge entry to the API-key provider", () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-config-test-"));
    created.push(directory);
    writeFileSync(join(directory, "config.json"), JSON.stringify({
      version: 1,
      selectedProviderId: "opencodex",
      selectedModelId: "gpt-5.6-luna",
      mode: "manual",
      theme: "system",
      providers: {
        opencodex: {
          id: "opencodex",
          displayName: "OpenCodex Bridge",
          kind: "opencodex",
          baseUrl: "http://127.0.0.1:10100/v1",
          enabled: true,
          local: true,
          auth: { method: "opencodex" },
          defaultModel: "gpt-5.6-luna",
          privacyNote: "legacy",
        },
      },
      allowedOrigins: ["http://localhost:10200"],
    }), "utf8");

    const settings = loadSettings({ ...process.env, OPENWORDCODE_DATA_DIR: directory });
    expect(settings.selectedProviderId).toBe("openai");
    expect(settings.selectedModelId).toBe("gpt-4.1-mini");
    expect(settings.providers["opencodex"]).toBeUndefined();
    expect(settings.providers["openwordcode-bridge"]).toMatchObject({ kind: "openwordcode-bridge", baseUrl: "http://127.0.0.1:10101/v1" });
  });
});
