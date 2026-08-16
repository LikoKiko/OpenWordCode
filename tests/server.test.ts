import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { buildServer, createCoreState } from "../apps/core/src/server.js";

const created: string[] = [];
afterEach(() => { for (const path of created.splice(0)) rmSync(path, { recursive: true, force: true }); });

describe("Core API", () => {
  it("protects management routes and streams a demo agent proposal", async () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-test-"));
    created.push(directory);
    const state = createCoreState({ ...process.env, OPENWORDCODE_DATA_DIR: directory, OPENWORDCODE_PORT: "0", OPENWORDCODE_OPENAI_OAUTH_CLIENT_ID: undefined, OPENWORDCODE_CHATGPT_OAUTH_CLIENT_ID: undefined });
    const app = await buildServer(state);
    const health = await app.inject({ method: "GET", url: "/health" });
    expect(health.statusCode).toBe(200);
    const denied = await app.inject({ method: "GET", url: "/api/providers", headers: { origin: "http://127.0.0.1:10200" } });
    expect(denied.statusCode).toBe(401);
    const bootstrap = await app.inject({ method: "GET", url: "/api/bootstrap", headers: { origin: "http://127.0.0.1:10200" } });
    expect(bootstrap.statusCode).toBe(200);
    const token = (bootstrap.json() as { sessionToken: string }).sessionToken;
    const headers = { origin: "http://127.0.0.1:10200", "x-openwordcode-session": token, "x-openwordcode-csrf": token };
    const providers = await app.inject({ method: "GET", url: "/api/providers", headers });
    expect(providers.statusCode).toBe(200);
    const oauthProviders = await app.inject({ method: "GET", url: "/api/oauth/providers", headers });
    expect(oauthProviders.statusCode).toBe(200);
    expect((oauthProviders.json() as { providers: Array<{ id: string; supported: boolean }> }).providers).toEqual(expect.arrayContaining([
      expect.objectContaining({ id: "anthropic", supported: true }),
      expect.objectContaining({ id: "google-antigravity", supported: true }),
      expect.objectContaining({ id: "cursor", supported: false }),
    ]));
    const chatgpt = (providers.json() as { providers: Array<{ id: string; auth: { status: string; method: string } }> }).providers.find(provider => provider.id === "openwordcode-bridge");
    expect(chatgpt?.auth.method).toBe("oauth");
    expect(chatgpt?.auth.status).toBe("login-required");
    const login = await app.inject({ method: "POST", url: "/api/auth/chatgpt/start", headers: { ...headers, "content-type": "application/json" }, payload: {} });
    expect(login.statusCode).toBe(200);
    expect((login.json() as { authorizeUrl: string }).authorizeUrl).toContain("https://auth.openai.com/oauth/authorize");
    const text = "The customers needs to submit the form.";
    const agent = await app.inject({ method: "POST", url: "/api/agent", headers: { ...headers, "content-type": "application/json" }, payload: { providerId: "demo", modelId: "demo-rewrite", instruction: "Fix grammar", mode: "manual", document: { documentId: "test", selection: { text, isEmpty: false, rangeStart: -1, rangeEnd: -1, target: { kind: "selection", id: "selection", beforeText: text, beforeFingerprint: "ignored" } }, documentText: text, paragraphs: [{ id: "p0", index: 0, text }], outline: [], capabilities: { canRead: true, canWrite: true, canComment: false, canFormat: false } } } });
    expect(agent.statusCode).toBe(200);
    expect(agent.body).toContain('"type":"proposal"');
    expect(agent.body).toContain("The customer needs to submit the form.");
    await app.close();
  });

  it("blocks unconfigured origins", async () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-origin-test-"));
    created.push(directory);
    const app = await buildServer(createCoreState({ ...process.env, OPENWORDCODE_DATA_DIR: directory }));
    const response = await app.inject({ method: "GET", url: "/api/bootstrap", headers: { origin: "https://malicious.example" } });
    expect(response.statusCode).toBe(403);
    await app.close();
  });

  it("accepts bounded image and PDF review attachments", async () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-attachments-test-"));
    created.push(directory);
    const app = await buildServer(createCoreState({ ...process.env, OPENWORDCODE_DATA_DIR: directory }));
    const bootstrap = await app.inject({ method: "GET", url: "/api/bootstrap", headers: { origin: "http://127.0.0.1:10200" } });
    const token = (bootstrap.json() as { sessionToken: string }).sessionToken;
    const headers = { origin: "http://127.0.0.1:10200", "x-openwordcode-session": token, "x-openwordcode-csrf": token, "content-type": "application/json" };
    const text = "Review the attachment.";
    const response = await app.inject({ method: "POST", url: "/api/agent", headers, payload: { providerId: "demo", modelId: "demo-rewrite", instruction: text, mode: "manual", attachments: [{ id: "image-1", name: "chart.png", mimeType: "image/png", size: 1, dataUrl: "data:image/png;base64,AA==" }, { id: "pdf-1", name: "brief.pdf", mimeType: "application/pdf", size: 1, dataUrl: "data:application/pdf;base64,AA==" }], document: { documentId: "test", selection: { text: "", isEmpty: true, selectedVisualElementIds: ["inline-picture-0"], target: { kind: "selection", id: "selection", beforeText: "", beforeFingerprint: "ignored" } }, documentText: "", paragraphs: [], visualElements: [{ id: "inline-picture-0", kind: "inlinePicture", index: 0, width: 72, height: 72, mimeType: "image/png", size: 1, dataUrl: "data:image/png;base64,AA==", contentAvailable: true }], outline: [], capabilities: { canRead: true, canWrite: true, canComment: false, canFormat: false } } } });
    expect(response.statusCode).toBe(200);
    expect(response.body).toContain('"type":"done"');
    await app.close();
  });
});
