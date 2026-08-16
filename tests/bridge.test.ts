import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { buildBridgeServer } from "../apps/core/src/bridge.js";
import { createCoreState } from "../apps/core/src/server.js";

const created: string[] = [];
afterEach(() => { for (const path of created.splice(0)) rmSync(path, { recursive: true, force: true }); });

describe("OpenWordCode Bridge", () => {
  it("owns the local compatibility surface and routes chat completions standalone", async () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-bridge-test-"));
    created.push(directory);
    const state = createCoreState({
      ...process.env,
      OPENWORDCODE_DATA_DIR: directory,
      OPENWORDCODE_BRIDGE_PROVIDER_ID: "demo",
    });
    const app = await buildBridgeServer(state);
    const health = await app.inject({ method: "GET", url: "/healthz" });
    expect(health.statusCode).toBe(200);
    expect(health.json()).toMatchObject({ service: "OpenWordCode Bridge" });

    const models = await app.inject({ method: "GET", url: "/v1/models" });
    expect(models.statusCode).toBe(200);
    expect(models.json()).toMatchObject({ data: [{ id: "demo-rewrite", owned_by: "openwordcode" }] });

    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "demo-rewrite", stream: false, messages: [{ role: "user", content: "Fix grammar" }] },
    });
    expect(response.statusCode).toBe(200);
    expect(response.json()).toMatchObject({ object: "chat.completion", choices: [{ message: { role: "assistant" } }] });

    const streamed = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "demo-rewrite", stream: true, messages: [{ role: "user", content: "Fix grammar" }] },
    });
    expect(streamed.statusCode).toBe(200);
    expect(streamed.headers["content-type"]).toContain("text/event-stream");
    expect(streamed.body).toContain("[DONE]");
    await app.close();
  });
});
