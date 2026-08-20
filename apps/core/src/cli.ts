import { startServer } from "./server.js";
import { bridgeRootUrl } from "./bridge.js";
import { OPENWORDCODE_VERSION } from "../../../packages/shared/src/index.js";

const port = Number(process.env.OPENWORDCODE_PORT ?? 10_200);
const base = `http://127.0.0.1:${port}`;

async function jsonRequest<T>(path: string, init: RequestInit = {}): Promise<T> {
  // Native CLI requests do not carry a browser Origin. This keeps alternate Core
  // ports working while the server still rejects untrusted browser origins.
  const bootstrap = await fetch(`${base}/api/bootstrap`);
  if (!bootstrap.ok) throw new Error(`Core bootstrap failed with HTTP ${bootstrap.status}`);
  const body = await bootstrap.json() as { sessionToken: string };
  const response = await fetch(`${base}${path}`, {
    ...init,
    headers: { "x-openwordcode-session": body.sessionToken, "x-openwordcode-csrf": body.sessionToken, ...(init.headers ?? {}) },
  });
  if (!response.ok) throw new Error(await response.text());
  return response.json() as Promise<T>;
}

async function main(): Promise<void> {
  const command = process.argv[2] ?? "start";
  if (command === "start") {
    console.log(`OpenWordCode ${OPENWORDCODE_VERSION} Core starting on ${base}`);
    await startServer();
    console.log(`OpenWordCode Core is ready at ${base}/health`);
    console.log(`OpenWordCode Bridge is ready at ${bridgeRootUrl()}/healthz`);
    return;
  }
  if (command === "status") {
    try { const response = await fetch(`${base}/health`); console.log(JSON.stringify(await response.json(), null, 2)); }
    catch { console.error(`OpenWordCode Core is offline at ${base}`); process.exitCode = 1; }
    return;
  }
  if (command === "doctor") {
    try {
      const health = await fetch(`${base}/health`);
      const bridge = await jsonRequest<Record<string, unknown>>("/api/openwordcode/bridge/status");
      console.log(JSON.stringify({ health: await health.json(), openwordcodeBridge: bridge }, null, 2));
    } catch (error) { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1; }
    return;
  }
  if (command === "providers") {
    try { console.log(JSON.stringify(await jsonRequest("/api/providers"), null, 2)); }
    catch (error) { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1; }
    return;
  }
  if (command === "models") {
    try { console.log(JSON.stringify(await jsonRequest("/api/models"), null, 2)); }
    catch (error) { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1; }
    return;
  }
  console.log("Usage: npm run start | npm run status | npm run doctor | npm run providers | npm run models");
}

void main();
