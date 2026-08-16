import { spawn, type ChildProcess } from "node:child_process";
import { access, readdir, readFile, stat } from "node:fs/promises";
import { homedir } from "node:os";
import { delimiter, isAbsolute, join, resolve } from "node:path";

const CODEX_APP_SERVER_TIMEOUT_MS = 8_000;
const CODEX_CLI_LOGIN_COMMAND = "login";

export interface CodexCliCredential {
  accessToken: string;
  accountId?: string;
  email?: string;
  expiresAt?: number;
}

export interface CodexCliLoginLaunch {
  executable: string;
  command: string;
}

export interface CodexCliRefreshResult {
  refreshed: boolean;
  detail?: string;
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function recordValue(value: unknown): Record<string, unknown> | undefined {
  return typeof value === "object" && value !== null && !Array.isArray(value)
    ? value as Record<string, unknown>
    : undefined;
}

function userHome(env: NodeJS.ProcessEnv, fallback: string): string {
  return env.USERPROFILE?.trim() || env.HOME?.trim() || fallback;
}

function expandHome(value: string, home: string): string {
  if (value === "~") return home;
  if (value.startsWith("~/") || value.startsWith("~\\")) return join(home, value.slice(2));
  return value;
}

function expandPath(value: string, home: string): string {
  const expanded = expandHome(value.trim(), home);
  return isAbsolute(expanded) ? expanded : resolve(expanded);
}

export function codexCliAuthPath(env: NodeJS.ProcessEnv = process.env, home = homedir()): string {
  const configured = env.CODEX_HOME?.trim();
  if (configured) {
    const expanded = expandHome(configured, userHome(env, home));
    return join(isAbsolute(expanded) ? expanded : resolve(expanded), "auth.json");
  }
  return join(userHome(env, home), ".codex", "auth.json");
}

async function isFile(path: string): Promise<boolean> {
  try {
    return (await stat(path)).isFile();
  } catch {
    return false;
  }
}

async function accessibleFile(path: string): Promise<string | null> {
  if (!(await isFile(path))) return null;
  try {
    await access(path);
    return path;
  } catch {
    return null;
  }
}

function executableName(): string {
  return process.platform === "win32" ? "codex.exe" : "codex";
}

function configuredExecutable(env: NodeJS.ProcessEnv, home: string): string | undefined {
  const configured = env.CODEX_CLI_PATH?.trim();
  if (!configured) return undefined;
  if (configured.includes("/") || configured.includes("\\") || isAbsolute(configured)) return expandPath(configured, home);
  return undefined;
}

async function localWindowsExecutables(env: NodeJS.ProcessEnv, home: string): Promise<string[]> {
  if (process.platform !== "win32") return [];
  const localAppData = env.LOCALAPPDATA?.trim() || join(userHome(env, home), "AppData", "Local");
  const roots = [
    join(localAppData, "OpenAI", "Codex", "bin"),
    join(localAppData, "OpenAI", "Codex"),
    join(localAppData, "Programs", "OpenAI Codex"),
    join(localAppData, "Programs", "Codex"),
  ];
  const candidates: Array<{ path: string; modified: number }> = [];
  for (const root of roots) {
    const direct = join(root, executableName());
    const directFile = await accessibleFile(direct);
    if (directFile) candidates.push({ path: directFile, modified: 0 });
    try {
      const entries = await readdir(root, { withFileTypes: true });
      for (const entry of entries) {
        if (!entry.isDirectory()) continue;
        const candidate = join(root, entry.name, executableName());
        const file = await accessibleFile(candidate);
        if (!file) continue;
        let modified = 0;
        try { modified = (await stat(file)).mtimeMs; } catch { /* best-effort ordering */ }
        candidates.push({ path: file, modified });
      }
    } catch {
      // The install root may not exist on this platform or for this user.
    }
  }
  return [...new Map(candidates.map(item => [item.path, item])).values()]
    .sort((left, right) => right.modified - left.modified)
    .map(item => item.path);
}

async function pathExecutables(env: NodeJS.ProcessEnv): Promise<string[]> {
  const name = executableName();
  const paths = (env.PATH ?? "").split(delimiter).map(value => value.trim()).filter(Boolean);
  const candidates: string[] = [];
  for (const directory of paths) {
    const candidate = await accessibleFile(join(directory, name));
    if (candidate) candidates.push(candidate);
  }
  return candidates;
}

/**
 * Finds an executable that OpenWordCode can actually start. On Windows the
 * Microsoft Store alias can appear on PATH but return Access Denied, so the
 * per-user OpenAI Codex installation is intentionally preferred.
 */
export async function findCodexCliExecutable(
  env: NodeJS.ProcessEnv = process.env,
  home = homedir(),
): Promise<string | null> {
  const configured = configuredExecutable(env, userHome(env, home));
  if (configured) return accessibleFile(configured);
  const local = await localWindowsExecutables(env, home);
  if (local.length) return local[0] ?? null;
  return (await pathExecutables(env))[0] ?? null;
}

function commandFor(executable: string): string {
  return process.platform === "win32" ? `"${executable.replaceAll('"', '\\"')}" ${CODEX_CLI_LOGIN_COMMAND}` : `${executable} ${CODEX_CLI_LOGIN_COMMAND}`;
}

function spawnDetached(executable: string, args: string[]): Promise<void> {
  return new Promise((resolvePromise, rejectPromise) => {
    let settled = false;
    let child: ChildProcess;
    try {
      child = spawn(executable, args, {
        detached: true,
        stdio: "ignore",
        windowsHide: false,
      });
    } catch (error) {
      rejectPromise(error);
      return;
    }
    child.once("error", error => {
      if (settled) return;
      settled = true;
      rejectPromise(error);
    });
    child.once("spawn", () => {
      if (settled) return;
      settled = true;
      child.unref();
      resolvePromise();
    });
  });
}

/** Starts the official Codex browser login after an explicit user action. */
export async function launchCodexCliLogin(
  env: NodeJS.ProcessEnv = process.env,
  home = homedir(),
): Promise<CodexCliLoginLaunch> {
  const executable = await findCodexCliExecutable(env, home);
  if (!executable) {
    throw new Error("The Codex CLI was not found. Install or update Codex, then try again, or set CODEX_CLI_PATH to the codex executable.");
  }
  await spawnDetached(executable, [CODEX_CLI_LOGIN_COMMAND]);
  return { executable, command: commandFor(executable) };
}

function jwtClaims(token: string): Record<string, unknown> {
  const payload = token.split(".")[1];
  if (!payload) return {};
  try {
    return recordValue(JSON.parse(Buffer.from(payload, "base64url").toString("utf8"))) ?? {};
  } catch {
    return {};
  }
}

function accountIdFromClaims(claims: Record<string, unknown>): string | undefined {
  const direct = stringValue(claims.chatgpt_account_id) ?? stringValue(claims.account_id);
  if (direct) return direct;
  for (const key of ["https://api.openai.com/auth", "https://api.openai.com/profile"]) {
    const nested = recordValue(claims[key]);
    const accountId = nested ? stringValue(nested.chatgpt_account_id) ?? stringValue(nested.account_id) : undefined;
    if (accountId) return accountId;
  }
  const organizations = Array.isArray(claims.organizations) ? claims.organizations : [];
  return stringValue(recordValue(organizations[0])?.id);
}

export function parseCodexCliCredential(raw: string): CodexCliCredential | null {
  try {
    const root = recordValue(JSON.parse(raw));
    const tokens = root ? recordValue(root.tokens) : undefined;
    const accessToken = tokens ? stringValue(tokens.access_token) : undefined;
    if (!accessToken) return null;
    const claims = jwtClaims(accessToken);
    const exp = typeof claims.exp === "number" && Number.isFinite(claims.exp) ? claims.exp * 1_000 : undefined;
    const accountId = stringValue(tokens?.account_id) ?? accountIdFromClaims(claims);
    const email = stringValue(claims.email)?.toLowerCase();
    return {
      accessToken,
      ...(accountId ? { accountId } : {}),
      ...(email ? { email } : {}),
      ...(exp !== undefined && Number.isFinite(exp) ? { expiresAt: exp } : {}),
    };
  } catch {
    return null;
  }
}

export async function readCodexCliCredential(
  env: NodeJS.ProcessEnv = process.env,
  read = readFile,
): Promise<CodexCliCredential | null> {
  try {
    return parseCodexCliCredential(await read(codexCliAuthPath(env), "utf8"));
  } catch {
    return null;
  }
}

export function isCodexCliCredentialLive(credential: CodexCliCredential, now = Date.now()): boolean {
  return credential.expiresAt === undefined || credential.expiresAt > now + 30_000;
}

function jsonLine(value: unknown): string {
  return `${JSON.stringify(value)}\n`;
}

function errorDetail(value: unknown): string {
  if (typeof value !== "object" || value === null) return "Codex did not refresh the session.";
  const record = value as Record<string, unknown>;
  const nested = typeof record.error === "object" && record.error !== null ? record.error as Record<string, unknown> : undefined;
  const code = typeof nested?.code === "string" ? nested.code : undefined;
  if (code === "invalid_refresh_token") return "The Codex refresh token is invalid. Sign in again with Codex CLI.";
  return "Codex could not refresh the session. Sign in again with Codex CLI.";
}

/**
 * Asks the installed Codex process to refresh its own credentials. This keeps
 * refresh-token handling inside Codex and avoids duplicating its auth storage.
 * The process is only started after the user has selected a Codex session.
 */
export async function refreshCodexCliSession(
  env: NodeJS.ProcessEnv = process.env,
  home = homedir(),
): Promise<CodexCliRefreshResult> {
  const executable = await findCodexCliExecutable(env, home);
  if (!executable) return { refreshed: false, detail: "The Codex CLI was not found." };
  return new Promise(resolvePromise => {
    let child: ChildProcess | undefined;
    let buffer = "";
    let finished = false;
    let timer: NodeJS.Timeout | undefined;
    const finish = (result: CodexCliRefreshResult): void => {
      if (finished) return;
      finished = true;
      if (timer) clearTimeout(timer);
      child?.stdout?.removeAllListeners();
      child?.stderr?.removeAllListeners();
      child?.removeAllListeners();
      try { child?.kill(); } catch { /* already exited */ }
      resolvePromise(result);
    };
    const send = (message: unknown): void => {
      try { child?.stdin?.write(jsonLine(message)); } catch { finish({ refreshed: false, detail: "Codex could not refresh the session." }); }
    };
    try {
      child = spawn(executable, ["app-server", "--stdio"], { stdio: ["pipe", "pipe", "ignore"], windowsHide: true });
    } catch {
      finish({ refreshed: false, detail: "Codex could not start its local app-server." });
      return;
    }
    timer = setTimeout(() => finish({ refreshed: false, detail: "Codex refresh timed out." }), CODEX_APP_SERVER_TIMEOUT_MS);
    child.once("error", () => finish({ refreshed: false, detail: "Codex could not start its local app-server." }));
    child.once("exit", () => {
      if (!finished) finish({ refreshed: false, detail: "Codex closed before refreshing the session." });
    });
    child.stdout?.on("data", chunk => {
      buffer += String(chunk);
      const lines = buffer.split(/\r?\n/u);
      buffer = lines.pop() ?? "";
      for (const line of lines) {
        if (!line.trim()) continue;
        let message: unknown;
        try { message = JSON.parse(line) as unknown; } catch { continue; }
        if (typeof message !== "object" || message === null) continue;
        const record = message as Record<string, unknown>;
        if (record.id === 1) {
          send({ jsonrpc: "2.0", id: 2, method: "account/read", params: { refreshToken: true } });
        } else if (record.id === 2) {
          const result = record.result;
          if (typeof result === "object" && result !== null && typeof (result as Record<string, unknown>).account === "object" && (result as Record<string, unknown>).account !== null) finish({ refreshed: true });
          else finish({ refreshed: false, detail: errorDetail(record) });
        }
      }
    });
    send({ jsonrpc: "2.0", id: 1, method: "initialize", params: { clientInfo: { name: "openwordcode", version: "0.1.0" }, capabilities: { experimentalApi: true } } });
  });
}
