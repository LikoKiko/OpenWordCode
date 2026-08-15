import { safeErrorMessage } from "../../../packages/security/src/index.js";

const SEARCH_TIMEOUT_MS = 60_000;

function responsesUrl(baseUrl: string): string {
  const normalized = baseUrl.replace(/\/+$/u, "");
  if (/\/responses$/iu.test(normalized)) return normalized;
  if (/\/v1$/iu.test(normalized)) return `${normalized}/responses`;
  return `${normalized}/v1/responses`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function textFromResponse(payload: unknown): string {
  if (!isRecord(payload)) return "";
  if (typeof payload.output_text === "string") return payload.output_text;
  const output = Array.isArray(payload.output) ? payload.output : [];
  const chunks: string[] = [];
  for (const item of output) {
    if (!isRecord(item) || !Array.isArray(item.content)) continue;
    for (const part of item.content) {
      if (!isRecord(part)) continue;
      if (typeof part.text === "string") chunks.push(part.text);
      if (typeof part.value === "string") chunks.push(part.value);
    }
  }
  return chunks.join("").trim();
}

async function readResponseEvents(response: Response): Promise<unknown[]> {
  if (!response.body) return [];
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  const events: unknown[] = [];
  let buffer = "";
  let bytes = 0;
  try {
    while (true) {
      const chunk = await reader.read();
      if (chunk.done) break;
      bytes += chunk.value.byteLength;
      if (bytes > 2_000_000) throw new Error("web-search response exceeded the safety limit");
      buffer += decoder.decode(chunk.value, { stream: true });
      const frames = buffer.split(/\r?\n\r?\n/u);
      buffer = frames.pop() ?? "";
      for (const frame of frames) {
        const data = frame.split(/\r?\n/u).filter(line => line.startsWith("data:")).map(line => line.slice(5).trim()).join("\n");
        if (!data || data === "[DONE]") continue;
        try { events.push(JSON.parse(data) as unknown); } catch { /* ignore keep-alive or malformed frames */ }
      }
    }
    const data = buffer.split(/\r?\n/u).filter(line => line.startsWith("data:")).map(line => line.slice(5).trim()).join("\n");
    if (data && data !== "[DONE]") {
      try { events.push(JSON.parse(data) as unknown); } catch { /* incomplete terminal frame */ }
    }
  } finally {
    reader.releaseLock();
  }
  return events;
}

function textFromEvents(events: unknown[]): string {
  const deltas: string[] = [];
  for (const event of events) {
    if (!isRecord(event)) continue;
    if (event.type === "response.output_text.delta" && typeof event.delta === "string") deltas.push(event.delta);
  }
  if (deltas.length) return deltas.join("").trim();
  return events.map(textFromResponse).find(Boolean) ?? "";
}

function sourcesFromResponse(payload: unknown): Array<{ title: string; url: string }> {
  const sources: Array<{ title: string; url: string }> = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) { value.forEach(visit); return; }
    if (!isRecord(value)) return;
    const url = typeof value.url === "string" ? value.url : typeof value.href === "string" ? value.href : "";
    const title = typeof value.title === "string" ? value.title : typeof value.name === "string" ? value.name : "Source";
    if (url.startsWith("http://") || url.startsWith("https://")) {
      if (!sources.some(source => source.url === url)) sources.push({ title, url });
    }
    Object.values(value).forEach(visit);
  };
  visit(payload);
  return sources.slice(0, 12);
}

export async function searchWithOpenWordCodeBridge(options: { baseUrl: string; model?: string; query: string; signal?: AbortSignal }): Promise<string> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), SEARCH_TIMEOUT_MS);
  const abort = () => controller.abort();
  if (options.signal?.aborted) abort();
  else options.signal?.addEventListener("abort", abort, { once: true });
  try {
    const response = await fetch(responsesUrl(options.baseUrl), {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "text/event-stream" },
      body: JSON.stringify({
        model: options.model || "gpt-5.6-luna",
        instructions: "Search the live web for the user's query. Answer concisely and finish with a Sources section containing title and URL for every source used.",
        input: [{ type: "message", role: "user", content: [{ type: "input_text", text: options.query }] }],
        tools: [{ type: "web_search" }],
        tool_choice: "auto",
        store: false,
        stream: true,
      }),
      signal: controller.signal,
    });
    if (!response.ok) return JSON.stringify({ error: `OpenWordCode Bridge web search returned HTTP ${response.status}.` });
    const events = await readResponseEvents(response);
    const text = textFromEvents(events);
    const sources = sourcesFromResponse(events);
    if (!text && !sources.length) return JSON.stringify({ error: "OpenWordCode Bridge returned no web-search result." });
    return JSON.stringify({ query: options.query, answer: text || "The search returned sources without a summary.", sources });
  } catch (error) {
    return JSON.stringify({ error: `Web search is unavailable: ${safeErrorMessage(error)}` });
  } finally {
    clearTimeout(timeout);
    options.signal?.removeEventListener("abort", abort);
  }
}
