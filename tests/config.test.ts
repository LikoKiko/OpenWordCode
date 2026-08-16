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
    expect(settings.providers.openai?.auth).toMatchObject({ method: "api-key", credentialRef: "provider:openai" });
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
});
