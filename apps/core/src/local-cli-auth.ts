import { spawn, type ChildProcess } from "node:child_process";
import { access, readFile, stat } from "node:fs/promises";
import { homedir } from "node:os";
import { delimiter, isAbsolute, join, resolve } from "node:path";

/**
 * Local CLI sessions are deliberately separate from OpenWordCode OAuth.
 * We only read the credential stores owned by the official CLI after the user
 * explicitly chooses "Connect existing session". Refreshed tokens stay in
 * memory and are never copied into OpenWordCode's credential store.
 */
export type LocalCliProviderId = "anthropic" | "kimi" | "google-antigravity";
export type LocalCliSource = "claude-cli" | "kimi-cli" | "antigravity-cli";

export interface LocalCliCredential {
  accessToken: string;
  refreshToken?: string;
  expiresAt?: number;
  accountId?: string;
  email?: string;
  projectId?: string;
  apiBaseUrl?: string;
}

export interface LocalCliLoginLaunch {
  providerId: LocalCliProviderId;
  executable: string;
  command: string;
}

const CLAUDE_CLIENT_ID = Buffer.from("OWQxYzI1MGEtZTYxYi00NGQ5LTg4ZWQtNTk0NGQxOTYyZjVl", "base64").toString("utf8");
const CLAUDE_TOKEN_URL = "https://api.anthropic.com/v1/oauth/token";
const KIMI_CLIENT_ID = "17e5f671-d194-4dfb-9706-5516cb48c098";
const KIMI_TOKEN_URL = "https://auth.kimi.com/api/oauth/token";
const KIMI_CLI_VERSION = "0.14.0";
const GOOGLE_CLIENT_ID = String.fromCharCode(49,48,55,49,48,48,54,48,54,48,53,57,49,45,116,109,104,115,115,105,110,50,104,50,49,108,99,114,101,50,51,53,118,116,111,108,111,106,104,52,103,52,48,51,101,112,46,97,112,112,115,46,103,111,111,103,108,101,117,115,101,114,99,111,110,116,101,110,116,46,99,111,109);
const DEFAULT_GOOGLE_CLIENT_SECRET = String.fromCharCode(71,79,67,83,80,88,45,75,53,56,70,87,82,52,56,54,76,100,76,74,49,109,76,66,56,115,88,67,52,122,54,113,68,65,102);
const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";
const GOOGLE_CLOUD_CODE_API = "https://cloudcode-pa.googleapis.com";
const GOOGLE_CLOUD_CODE_DAILY_API = "https://daily-cloudcode-pa.googleapis.com";
const GOOGLE_API_VERSION = "v1internal";
const ANTIGRAVITY_CREDENTIAL_TARGET = "LegacyGeneric:target=gemini:antigravity";
const LOCAL_CLI_TIMEOUT_MS = 15_000;

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function recordValue(value: unknown): Record<string, unknown> | undefined {
  return typeof value === "object" && value !== null && !Array.isArray(value)
    ? value as Record<string, unknown>
    : undefined;
}

function numberValue(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim()) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }
  return undefined;
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

function parseExpiry(value: unknown): number | undefined {
  const numeric = numberValue(value);
  if (numeric !== undefined) return numeric > 10_000_000_000 ? numeric : numeric * 1_000;
  if (typeof value === "string" && value.trim()) {
    const parsed = Date.parse(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }
  return undefined;
}

function jwtExpiry(token: string): number | undefined {
  const payload = token.split(".")[1];
  if (!payload) return undefined;
  try {
    const claims = recordValue(JSON.parse(Buffer.from(payload, "base64url").toString("utf8")));
    return parseExpiry(claims?.exp);
  } catch {
    return undefined;
  }
}

function sourceFor(providerId: LocalCliProviderId): LocalCliSource {
  if (providerId === "anthropic") return "claude-cli";
  if (providerId === "kimi") return "kimi-cli";
  return "antigravity-cli";
}

export function localCliSource(providerId: string): LocalCliSource | undefined {
  return providerId === "anthropic" || providerId === "kimi" || providerId === "google-antigravity"
    ? sourceFor(providerId)
    : undefined;
}

export function localCliProviderIsSupported(providerId: string): providerId is LocalCliProviderId {
  return localCliSource(providerId) !== undefined;
}

export function localCliDisplayName(providerId: LocalCliProviderId): string {
  if (providerId === "anthropic") return "Claude Code";
  if (providerId === "kimi") return "Kimi CLI";
  return "Antigravity CLI";
}

export function claudeCliCredentialsPath(env: NodeJS.ProcessEnv = process.env, home = homedir()): string {
  const configDir = env.CLAUDE_CONFIG_DIR?.trim();
  const root = configDir ? expandPath(configDir, userHome(env, home)) : join(userHome(env, home), ".claude");
  return join(root, ".credentials.json");
}

export function kimiCliCredentialsPath(env: NodeJS.ProcessEnv = process.env, home = homedir()): string {
  const shareDir = env.KIMI_SHARE_DIR?.trim() || env.KIMI_CODE_HOME?.trim();
  const root = shareDir ? expandPath(shareDir, userHome(env, home)) : join(userHome(env, home), ".kimi");
  return join(root, "credentials", "kimi-code.json");
}

function parseTokenPayload(value: unknown): LocalCliCredential | null {
  const root = recordValue(value);
  if (!root) return null;
  const candidates = [
    root,
    recordValue(root.claudeAiOauth),
    recordValue(root.oauth),
    recordValue(root.tokens),
    recordValue(root.token),
    recordValue(root.credentials),
  ].filter((candidate): candidate is Record<string, unknown> => Boolean(candidate));
  for (const candidate of candidates) {
    const accessToken = stringValue(candidate.access_token) ?? stringValue(candidate.accessToken) ?? stringValue(candidate.token);
    if (!accessToken) continue;
    const refreshToken = stringValue(candidate.refresh_token) ?? stringValue(candidate.refreshToken);
    const expiresAt = parseExpiry(candidate.expires_at ?? candidate.expiresAt ?? candidate.expiry_date ?? candidate.expiry ?? candidate.expires);
    const accountId = stringValue(candidate.account_id) ?? stringValue(candidate.accountId) ?? stringValue(candidate.sub);
    const email = stringValue(candidate.email) ?? stringValue(candidate.email_address) ?? stringValue(candidate.emailAddress);
    const projectId = stringValue(candidate.project_id) ?? stringValue(candidate.projectId) ?? stringValue(candidate.cloudaicompanionProject);
    const apiBaseUrl = stringValue(candidate.api_base_url) ?? stringValue(candidate.apiBaseUrl);
    return {
      accessToken,
      ...(refreshToken ? { refreshToken } : {}),
      ...(expiresAt ?? jwtExpiry(accessToken) ? { expiresAt: expiresAt ?? jwtExpiry(accessToken) } : {}),
      ...(accountId ? { accountId } : {}),
      ...(email ? { email: email.toLowerCase() } : {}),
      ...(projectId ? { projectId } : {}),
      ...(apiBaseUrl ? { apiBaseUrl } : {}),
    };
  }
  return null;
}

export function parseClaudeCliCredential(raw: string): LocalCliCredential | null {
  try { return parseTokenPayload(JSON.parse(raw)); } catch { return null; }
}

export function parseKimiCliCredential(raw: string): LocalCliCredential | null {
  try { return parseTokenPayload(JSON.parse(raw)); } catch { return null; }
}

export function parseAntigravityCliCredential(raw: string): LocalCliCredential | null {
  try { return parseTokenPayload(JSON.parse(raw)); } catch { return null; }
}

export function isLocalCliCredentialLive(credential: LocalCliCredential, now = Date.now()): boolean {
  return credential.expiresAt === undefined || credential.expiresAt > now + 30_000;
}

async function isFile(path: string): Promise<boolean> {
  try { return (await stat(path)).isFile(); } catch { return false; }
}

async function accessibleFile(path: string): Promise<string | null> {
  if (!(await isFile(path))) return null;
  try { await access(path); return path; } catch { return null; }
}

function configuredCliPath(providerId: LocalCliProviderId, env: NodeJS.ProcessEnv, home: string): string | undefined {
  const variable = providerId === "anthropic" ? "CLAUDE_CLI_PATH" : providerId === "kimi" ? "KIMI_CLI_PATH" : "ANTIGRAVITY_CLI_PATH";
  const configured = env[variable]?.trim();
  if (!configured) return undefined;
  return configured.includes("/") || configured.includes("\\") || isAbsolute(configured) ? expandPath(configured, home) : undefined;
}

function cliNames(providerId: LocalCliProviderId): string[] {
  const base = providerId === "anthropic" ? "claude" : providerId === "kimi" ? "kimi" : "agy";
  if (process.platform !== "win32") return [base];
  return [`${base}.exe`, `${base}.cmd`, `${base}.bat`, `${base}.ps1`, base];
}

export async function findLocalCliExecutable(
  providerId: LocalCliProviderId,
  env: NodeJS.ProcessEnv = process.env,
  home = homedir(),
): Promise<string | null> {
  const configured = configuredCliPath(providerId, env, userHome(env, home));
  if (configured) return accessibleFile(configured);
  const pathEntries = (env.PATH ?? "").split(delimiter).map(value => value.trim()).filter(Boolean);
  for (const directory of pathEntries) {
    for (const name of cliNames(providerId)) {
      const candidate = await accessibleFile(join(directory, name));
      if (candidate) return candidate;
    }
  }
  return null;
}

function invocation(executable: string, args: string[]): { executable: string; args: string[] } {
  if (process.platform === "win32" && executable.toLowerCase().endsWith(".ps1")) {
    return { executable: "powershell.exe", args: ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", executable, ...args] };
  }
  if (process.platform === "win32" && /\.(?:cmd|bat)$/iu.test(executable)) {
    const command = [`"${executable.replaceAll('"', '\\\"')}"`, ...args.map(value => /\s/u.test(value) ? `"${value.replaceAll('"', '\\\"')}"` : value)].join(" ");
    return { executable: process.env.ComSpec?.trim() || "cmd.exe", args: ["/d", "/s", "/c", command] };
  }
  return { executable, args };
}

function displayCommand(executable: string, args: string[]): string {
  const quote = (value: string): string => /\s/u.test(value) ? `"${value.replaceAll('"', '\\\"')}"` : value;
  return [quote(executable), ...args.map(quote)].join(" ");
}

function spawnDetached(executable: string, args: string[]): Promise<void> {
  const command = invocation(executable, args);
  return new Promise((resolvePromise, rejectPromise) => {
    let settled = false;
    let child: ChildProcess;
    try {
      child = spawn(command.executable, command.args, { detached: true, stdio: "ignore", windowsHide: false });
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

function spawnVisibleWindows(executable: string, args: string[], env: NodeJS.ProcessEnv, title: string): Promise<void> {
  const command = invocation(executable, args);
  const quote = (value: string): string => `"${value.replaceAll('"', '\\\"')}"`;
  const line = `start "${title.replaceAll('"', "")}" ${quote(command.executable)} ${command.args.map(quote).join(" ")}`;
  return spawnDetached(env.ComSpec?.trim() || "cmd.exe", ["/d", "/c", line]);
}

/** Starts the official CLI's own login flow after an explicit user action. */
export async function launchLocalCliLogin(
  providerId: LocalCliProviderId,
  env: NodeJS.ProcessEnv = process.env,
  home = homedir(),
): Promise<LocalCliLoginLaunch> {
  const executable = await findLocalCliExecutable(providerId, env, home);
  if (!executable) {
    const name = providerId === "anthropic" ? "Claude Code" : providerId === "kimi" ? "Kimi CLI" : "Antigravity CLI";
    const variable = providerId === "anthropic" ? "CLAUDE_CLI_PATH" : providerId === "kimi" ? "KIMI_CLI_PATH" : "ANTIGRAVITY_CLI_PATH";
    throw new Error(`${name} was not found. Install it, then try again, or set ${variable} to its executable.`);
  }
  const args = providerId === "google-antigravity" ? [] : ["login"];
  if (process.platform === "win32") await spawnVisibleWindows(executable, args, env, `${localCliDisplayName(providerId)} sign-in`);
  else await spawnDetached(executable, args);
  return { providerId, executable, command: displayCommand(invocation(executable, args).executable, invocation(executable, args).args) };
}

async function readFileCredential(path: string, parser: (raw: string) => LocalCliCredential | null, reader = readFile): Promise<LocalCliCredential | null> {
  try { return parser(await reader(path, "utf8")); } catch { return null; }
}

interface CapturedProcess {
  stdout: string;
  code: number;
}

function captureProcess(executable: string, args: string[], timeoutMs = LOCAL_CLI_TIMEOUT_MS): Promise<CapturedProcess> {
  return new Promise(resolvePromise => {
    const child = spawn(executable, args, { stdio: ["ignore", "pipe", "ignore"], windowsHide: true });
    let stdout = "";
    let settled = false;
    const finish = (code: number): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolvePromise({ stdout, code });
    };
    const timer = setTimeout(() => { try { child.kill(); } catch { /* already exited */ } finish(-1); }, timeoutMs);
    child.stdout?.on("data", chunk => {
      stdout += String(chunk);
      if (stdout.length > 2_000_000) { try { child.kill(); } catch { /* already exited */ } finish(-1); }
    });
    child.once("error", () => finish(-1));
    child.once("exit", code => finish(code ?? -1));
  });
}

const WINDOWS_CREDENTIAL_SCRIPT = `
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OpenWordCodeCredentialReader {
  [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
  public struct CREDENTIAL {
    public uint Flags; public uint Type; public IntPtr TargetName; public IntPtr Comment;
    public long LastWritten; public uint CredentialBlobSize; public IntPtr CredentialBlob;
    public uint Persist; public uint AttributeCount; public IntPtr Attributes;
    public IntPtr TargetAlias; public IntPtr UserName;
  }
  [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  public static extern bool CredRead(string target, uint type, uint flags, out IntPtr credential);
  [DllImport("advapi32.dll", SetLastError = true)]
  public static extern bool CredFree(IntPtr credential);
}
'@
$ptr = [IntPtr]::Zero
if (-not [OpenWordCodeCredentialReader]::CredRead('__TARGET__', 1, 0, [ref]$ptr)) { exit 2 }
try {
  $credential = [Runtime.InteropServices.Marshal]::PtrToStructure($ptr, [type][OpenWordCodeCredentialReader+CREDENTIAL])
  if ($credential.CredentialBlobSize -gt 0) {
    $bytes = New-Object byte[] ([int]$credential.CredentialBlobSize)
    [Runtime.InteropServices.Marshal]::Copy($credential.CredentialBlob, $bytes, 0, $bytes.Length)
    [Console]::OpenStandardOutput().Write($bytes, 0, $bytes.Length)
  }
} finally { [OpenWordCodeCredentialReader]::CredFree($ptr) | Out-Null }
`;

async function readWindowsCredentialTarget(target: string): Promise<string | null> {
  const script = WINDOWS_CREDENTIAL_SCRIPT.replace("__TARGET__", target.replaceAll("'", "''"));
  const encoded = Buffer.from(script, "utf16le").toString("base64");
  const result = await captureProcess("powershell.exe", ["-NoProfile", "-NonInteractive", "-EncodedCommand", encoded]);
  return result.code === 0 && result.stdout.trim() ? result.stdout : null;
}

async function readAntigravitySecret(env: NodeJS.ProcessEnv, home: string): Promise<string | null> {
  const configured = env.ANTIGRAVITY_CREDENTIALS_PATH?.trim();
  if (configured) {
    try { return await readFile(expandPath(configured, userHome(env, home)), "utf8"); } catch { return null; }
  }
  if (process.platform === "win32") return readWindowsCredentialTarget(ANTIGRAVITY_CREDENTIAL_TARGET);
  if (process.platform === "darwin") {
    const result = await captureProcess("security", ["find-generic-password", "-s", "gemini:antigravity", "-w"]);
    return result.code === 0 && result.stdout.trim() ? result.stdout.trim() : null;
  }
  const result = await captureProcess("secret-tool", ["lookup", "service", "gemini:antigravity"]);
  return result.code === 0 && result.stdout.trim() ? result.stdout.trim() : null;
}

export async function readLocalCliCredential(
  providerId: LocalCliProviderId,
  env: NodeJS.ProcessEnv = process.env,
  home = homedir(),
): Promise<LocalCliCredential | null> {
  if (providerId === "anthropic") return readFileCredential(claudeCliCredentialsPath(env, home), parseClaudeCliCredential);
  if (providerId === "kimi") return readFileCredential(kimiCliCredentialsPath(env, home), parseKimiCliCredential);
  const raw = await readAntigravitySecret(env, userHome(env, home));
  return raw ? parseAntigravityCliCredential(raw) : null;
}

function tokenResponse(value: unknown, previous: LocalCliCredential): LocalCliCredential {
  const payload = recordValue(value);
  const accessToken = payload ? stringValue(payload.access_token) ?? stringValue(payload.accessToken) : undefined;
  if (!accessToken) throw new Error("The local CLI did not return a refreshed access token.");
  const refreshToken = payload ? stringValue(payload.refresh_token) ?? stringValue(payload.refreshToken) : undefined;
  const expiresIn = payload ? numberValue(payload.expires_in) : undefined;
  return {
    ...previous,
    accessToken,
    ...(refreshToken ?? previous.refreshToken ? { refreshToken: refreshToken ?? previous.refreshToken } : {}),
    ...(expiresIn !== undefined ? { expiresAt: Date.now() + expiresIn * 1_000 - 120_000 } : {}),
  };
}

async function postToken(fetchImpl: typeof fetch, url: string, values: Record<string, string>, headers: Record<string, string> = {}): Promise<Record<string, unknown>> {
  const response = await fetchImpl(url, {
    method: "POST",
    headers: { Accept: "application/json", "Content-Type": "application/x-www-form-urlencoded", ...headers },
    body: new URLSearchParams(values).toString(),
    signal: AbortSignal.timeout(LOCAL_CLI_TIMEOUT_MS),
  });
  const payload = recordValue(await response.json().catch(() => ({}))) ?? {};
  if (!response.ok) throw new Error(stringValue(payload.error_description) ?? stringValue(payload.error) ?? `Local CLI token refresh failed (${response.status})`);
  return payload;
}

async function discoverAntigravityProject(accessToken: string, fetchImpl: typeof fetch): Promise<string | undefined> {
  const headers = { Authorization: `Bearer ${accessToken}`, Accept: "*/*", "Content-Type": "application/json", "User-Agent": "antigravity/ide/2.5.5 (aidev_client; os_type=windows; arch=amd64)" };
  const extract = (value: unknown): string | undefined => {
    const root = recordValue(value);
    if (!root) return undefined;
    for (const key of ["cloudaicompanionProject", "projectId", "project"]) {
      const candidate = root[key];
      if (typeof candidate === "string" && candidate.trim()) return candidate.trim();
      const nested = recordValue(candidate);
      const id = stringValue(nested?.id);
      if (id) return id;
    }
    return undefined;
  };
  const first = await fetchImpl(`${GOOGLE_CLOUD_CODE_API}/${GOOGLE_API_VERSION}:loadCodeAssist`, { method: "POST", headers, body: JSON.stringify({ metadata: { ideType: "ANTIGRAVITY" } }), signal: AbortSignal.timeout(LOCAL_CLI_TIMEOUT_MS) });
  const firstPayload = await first.json().catch(() => ({}));
  const found = extract(firstPayload);
  if (found) return found;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const response = await fetchImpl(`${GOOGLE_CLOUD_CODE_DAILY_API}/${GOOGLE_API_VERSION}:onboardUser`, {
      method: "POST",
      headers: { ...headers, "x-goog-api-client": "google-api-nodejs-client/10.3.0" },
      body: JSON.stringify({ tier_id: "free-tier", metadata: { ide_type: "ANTIGRAVITY", ide_name: "antigravity", ide_version: "2.5.5" } }),
      signal: AbortSignal.timeout(LOCAL_CLI_TIMEOUT_MS),
    });
    const payload = await response.json().catch(() => ({}));
    const project = extract(recordValue(payload)?.response ?? payload);
    if (project && response.ok) return project;
  }
  return undefined;
}

/** Refreshes a CLI-owned session without writing back to the CLI's files. */
export async function refreshLocalCliCredential(
  providerId: LocalCliProviderId,
  previous: LocalCliCredential,
  env: NodeJS.ProcessEnv = process.env,
  fetchImpl: typeof fetch = fetch,
): Promise<LocalCliCredential> {
  if (!previous.refreshToken) throw new Error(`${localCliDisplayName(providerId)} session expired; sign in again with the CLI.`);
  if (providerId === "anthropic") {
    const payload = await postToken(fetchImpl, CLAUDE_TOKEN_URL, { grant_type: "refresh_token", client_id: CLAUDE_CLIENT_ID, refresh_token: previous.refreshToken });
    return tokenResponse(payload, previous);
  }
  if (providerId === "kimi") {
    const payload = await postToken(fetchImpl, KIMI_TOKEN_URL, { grant_type: "refresh_token", client_id: KIMI_CLIENT_ID, refresh_token: previous.refreshToken }, {
      "User-Agent": `KimiCLI/${KIMI_CLI_VERSION}`,
      "X-Msh-Platform": "kimi_cli",
      "X-Msh-Version": KIMI_CLI_VERSION,
      "X-Msh-Device-Id": "openwordcode-local-cli",
    });
    return tokenResponse(payload, previous);
  }
  const clientSecret = env.GOOGLE_ANTIGRAVITY_CLIENT_SECRET?.trim() || DEFAULT_GOOGLE_CLIENT_SECRET;
  const payload = await postToken(fetchImpl, GOOGLE_TOKEN_URL, {
    grant_type: "refresh_token",
    client_id: env.GOOGLE_ANTIGRAVITY_CLIENT_ID?.trim() || GOOGLE_CLIENT_ID,
    refresh_token: previous.refreshToken,
    client_secret: clientSecret,
  });
  return { ...tokenResponse(payload, previous), projectId: previous.projectId };
}

export async function ensureAntigravityProject(credential: LocalCliCredential, fetchImpl: typeof fetch = fetch): Promise<LocalCliCredential> {
  if (credential.projectId) return credential;
  const projectId = await discoverAntigravityProject(credential.accessToken, fetchImpl);
  return projectId ? { ...credential, projectId } : credential;
}
