import { spawn } from "node:child_process";
import { isAbsolute, relative, resolve } from "node:path";

const MAX_COMMAND_LENGTH = 600;
const MAX_OUTPUT_LENGTH = 20_000;
const DEFAULT_TIMEOUT_MS = 12_000;

const BLOCKED_PATTERNS = [
  /[\r\n]/u,
  /[;&|<>`]/u,
  /\$\(/u,
  /\$\{/u,
  /(?:remove-item|del(?:ete)?|erase|format(?:-volume)?|diskpart|clear-(?:content|item)|set-content|add-content|out-file|new-item|copy-item|move-item|rename-item)/iu,
  /(?:invoke-expression|\biex\b|invoke-webrequest|invoke-restmethod|curl(?:\.exe)?|wget(?:\.exe)?|bitsadmin|certutil|start-process|start-job|runas|set-executionpolicy|encodedcommand|frombase64string)/iu,
  /(?:shutdown|restart-computer|stop-computer|sc\.exe|net\s+(?:user|localgroup)|takeown|icacls|reg(?:\.exe)?\s)/iu,
];

function inside(root: string, candidate: string): boolean {
  const relativePath = relative(root, candidate);
  return relativePath === "" || (!relativePath.startsWith("..") && !isAbsolute(relativePath));
}

export function resolveWorkspaceDirectory(workspaceRoot: string, requested?: string): string {
  const root = resolve(workspaceRoot);
  const candidate = resolve(root, requested?.trim() || ".");
  if (!inside(root, candidate)) throw new Error("The console working directory must stay inside the OpenWordCode workspace.");
  return candidate;
}

export function validateConsoleCommand(command: string): { safe: boolean; reason: string } {
  const value = command.trim();
  if (!value) return { safe: false, reason: "The console command is empty." };
  if (value.length > MAX_COMMAND_LENGTH) return { safe: false, reason: `Commands are limited to ${MAX_COMMAND_LENGTH} characters.` };
  if (BLOCKED_PATTERNS.some(pattern => pattern.test(value))) return { safe: false, reason: "This command can modify the machine, execute downloaded code, or escape the safe read-only console policy." };
  if (/(?:^|\s)(?:[A-Za-z]:[\\/]|\\\\|\.\.[\\/]|\/)/u.test(value)) return { safe: false, reason: "The safe console cannot address absolute paths or paths outside the workspace." };

  const first = /^([A-Za-z][A-Za-z0-9_.-]*)/u.exec(value)?.[1]?.toLocaleLowerCase();
  if (!first) return { safe: false, reason: "The command must start with an approved read-only program." };
  const allowed = new Set(["get-childitem", "get-content", "get-location", "get-date", "get-process", "get-command", "dir", "ls", "type", "cat", "pwd", "where", "whoami", "ver", "select-string", "rg", "git", "npm", "node"]);
  if (!allowed.has(first)) return { safe: false, reason: `${first} is not enabled in the safe console.` };
  if (first === "git" && !/^git\s+(?:status|diff|log|show|branch(?:\s+--show-current)?)(?:\s|$)/iu.test(value)) return { safe: false, reason: "Only git status, diff, log, show, and branch --show-current are enabled." };
  if (first === "npm" && !/^npm\s+(?:(?:run\s+)?(?:test|typecheck|doctor|build)|--version|version)(?:\s|$)/iu.test(value)) return { safe: false, reason: "Only npm test, typecheck, doctor, build, and version checks are enabled." };
  if (first === "node" && !/^node\s+(?:--version|-v)(?:\s|$)/iu.test(value)) return { safe: false, reason: "Node execution is limited to version checks." };
  return { safe: true, reason: "Read-only workspace command." };
}

export interface ConsoleExecutionRequest {
  command: string;
  workspaceRoot: string;
  workingDirectory?: string;
  signal?: AbortSignal;
  timeoutMs?: number;
}

export interface ConsoleExecutionResult {
  ok: boolean;
  output: string;
  workingDirectory: string;
  failureReason?: string;
}

export async function executeConsoleCommand(request: ConsoleExecutionRequest): Promise<ConsoleExecutionResult> {
  const command = request.command.trim();
  const validation = validateConsoleCommand(command);
  const workingDirectory = resolveWorkspaceDirectory(request.workspaceRoot, request.workingDirectory);
  if (!validation.safe) return { ok: false, output: "", workingDirectory, failureReason: validation.reason };
  if (process.platform !== "win32") return { ok: false, output: "", workingDirectory, failureReason: "The Windows console is only available on Windows." };

  return new Promise(resolveResult => {
    const child = spawn("powershell.exe", ["-NoLogo", "-NoProfile", "-NonInteractive", "-Command", command], { cwd: workingDirectory, windowsHide: true, shell: false, stdio: ["ignore", "pipe", "pipe"] });
    let output = "";
    let settled = false;
    const timeout = setTimeout(() => {
      child.kill();
      finish({ ok: false, output, workingDirectory, failureReason: "The console command timed out." });
    }, request.timeoutMs ?? DEFAULT_TIMEOUT_MS);
    const onAbort = () => {
      child.kill();
      finish({ ok: false, output, workingDirectory, failureReason: "The console command was cancelled." });
    };
    if (request.signal?.aborted) onAbort();
    else request.signal?.addEventListener("abort", onAbort, { once: true });
    const append = (chunk: Buffer): void => {
      if (output.length < MAX_OUTPUT_LENGTH) output += chunk.toString("utf8").slice(0, MAX_OUTPUT_LENGTH - output.length);
    };
    child.stdout.on("data", append);
    child.stderr.on("data", append);
    child.on("error", error => finish({ ok: false, output, workingDirectory, failureReason: error.message }));
    child.on("close", code => finish({ ok: code === 0, output: output.trimEnd(), workingDirectory, ...(code === 0 ? {} : { failureReason: `The command exited with code ${code ?? "unknown"}.` }) }));
    function finish(result: ConsoleExecutionResult): void {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      request.signal?.removeEventListener("abort", onAbort);
      resolveResult(result);
    }
  });
}
