import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";
import { chmodSync, existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

export interface CredentialStore {
  get(id: string): Promise<string | null>;
  set(id: string, secret: string, options?: CredentialWriteOptions): Promise<void>;
  remove(id: string): Promise<void>;
  kind: "encrypted-file" | "memory";
}

export interface CredentialWriteOptions {
  /** Runs at the final synchronous persistence boundary. Throw to discard a stale write. */
  assertBeforePersist?: () => void;
}

interface EncryptedEntry {
  iv: string;
  tag: string;
  ciphertext: string;
}

interface CredentialFile {
  version: 1;
  entries: Record<string, EncryptedEntry>;
}

/**
 * Local fallback credential store. The config never contains the API key itself;
 * this file is encrypted at rest and permission-hardened where the platform allows.
 * Native Credential Manager/Keychain implementations can replace this class without
 * changing the provider or UI contracts.
 */
export class EncryptedFileCredentialStore implements CredentialStore {
  readonly kind = "encrypted-file" as const;
  private readonly keyPath: string;
  private readonly dataPath: string;

  constructor(directory: string) {
    this.keyPath = join(directory, "credential-store.key");
    this.dataPath = join(directory, "credentials.enc.json");
    mkdirSync(directory, { recursive: true, mode: 0o700 });
    try { chmodSync(directory, 0o700); } catch { /* Windows ACL hardening belongs in a native store. */ }
  }

  async get(id: string): Promise<string | null> {
    const entry = this.readFile().entries[id];
    if (!entry) return null;
    try {
      const decipher = createDecipheriv("aes-256-gcm", this.key(), Buffer.from(entry.iv, "base64url"));
      decipher.setAuthTag(Buffer.from(entry.tag, "base64url"));
      return Buffer.concat([decipher.update(Buffer.from(entry.ciphertext, "base64url")), decipher.final()]).toString("utf8");
    } catch {
      throw new Error("credential store could not decrypt the configured secret");
    }
  }

  async set(id: string, secret: string, options: CredentialWriteOptions = {}): Promise<void> {
    if (!id || !secret || /[\r\n]/.test(secret)) throw new Error("credential must be a non-empty single-line secret");
    const iv = randomBytes(12);
    const cipher = createCipheriv("aes-256-gcm", this.key(), iv);
    const ciphertext = Buffer.concat([cipher.update(secret, "utf8"), cipher.final()]);
    const file = this.readFile();
    file.entries[id] = {
      iv: iv.toString("base64url"),
      tag: cipher.getAuthTag().toString("base64url"),
      ciphertext: ciphertext.toString("base64url"),
    };
    options.assertBeforePersist?.();
    this.writeFile(file);
  }

  async remove(id: string): Promise<void> {
    const file = this.readFile();
    delete file.entries[id];
    this.writeFile(file);
  }

  private key(): Buffer {
    if (!existsSync(this.keyPath)) {
      writeFileSync(this.keyPath, randomBytes(32), { mode: 0o600, flag: "wx" });
      this.harden(this.keyPath);
    }
    const key = readFileSync(this.keyPath);
    if (key.length !== 32) throw new Error("credential store key is invalid");
    return key;
  }

  private readFile(): CredentialFile {
    if (!existsSync(this.dataPath)) return { version: 1, entries: {} };
    try {
      const parsed = JSON.parse(readFileSync(this.dataPath, "utf8")) as Partial<CredentialFile>;
      if (parsed.version !== 1 || !parsed.entries || typeof parsed.entries !== "object") throw new Error("invalid credential file");
      return { version: 1, entries: parsed.entries as Record<string, EncryptedEntry> };
    } catch (error) {
      throw new Error(`credential store is unreadable: ${error instanceof Error ? error.message : "invalid file"}`);
    }
  }

  private writeFile(file: CredentialFile): void {
    const temporary = `${this.dataPath}.${process.pid}.tmp`;
    writeFileSync(temporary, JSON.stringify(file, null, 2), { mode: 0o600 });
    this.harden(temporary);
    renameSync(temporary, this.dataPath);
    this.harden(this.dataPath);
  }

  private harden(path: string): void {
    try { chmodSync(path, 0o600); } catch { /* Windows ACL hardening belongs in a native store. */ }
  }
}

export class MemoryCredentialStore implements CredentialStore {
  readonly kind = "memory" as const;
  private readonly values = new Map<string, string>();
  async get(id: string): Promise<string | null> { return this.values.get(id) ?? null; }
  async set(id: string, secret: string, options: CredentialWriteOptions = {}): Promise<void> {
    options.assertBeforePersist?.();
    this.values.set(id, secret);
  }
  async remove(id: string): Promise<void> { this.values.delete(id); }
}

export function resolveEnvironmentReference(value: string, env: NodeJS.ProcessEnv = process.env): string | null {
  const match = /^\$\{([A-Za-z_][A-Za-z0-9_]*)\}$/.exec(value.trim());
  const name = match?.[1];
  return name ? env[name]?.trim() || null : null;
}

export function credentialDirectory(dataDirectory: string): string {
  return join(dirname(dataDirectory), `${dataDirectory.split(/[\\/]/).pop() ?? "openwordcode"}-credentials`);
}
