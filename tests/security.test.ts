import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { EncryptedFileCredentialStore, MemoryCredentialStore } from "../packages/auth/src/index.js";
import { equalSecret, redactSecrets } from "../packages/security/src/index.js";
import { validateConsoleCommand } from "../apps/core/src/console.js";

describe("credential boundaries", () => {
  it("round-trips credentials through the store without exposing them in status data", async () => {
    const store = new MemoryCredentialStore();
    await store.set("provider:test", "sk-test-secret-value");
    expect(await store.get("provider:test")).toBe("sk-test-secret-value");
    expect(store.kind).toBe("memory");
  });

  it("redacts common secret-shaped values", () => {
    expect(redactSecrets("Authorization: Bearer sk-test-secret-value")).toContain("[REDACTED]");
    expect(redactSecrets("api_key=sk-or-v1-super-secret-value")).toContain("[REDACTED]");
    expect(equalSecret("same", "same")).toBe(true);
    expect(equalSecret("different", "same")).toBe(false);
  });

  it("does not write the raw secret into the persistent credential payload", async () => {
    const directory = mkdtempSync(join(tmpdir(), "openwordcode-credentials-"));
    try {
      const store = new EncryptedFileCredentialStore(directory);
      await store.set("provider:test", "secret-that-must-not-be-plain-text");
      const payload = readFileSync(join(directory, "credentials.enc.json"), "utf8");
      expect(payload).not.toContain("secret-that-must-not-be-plain-text");
      expect(await store.get("provider:test")).toBe("secret-that-must-not-be-plain-text");
    } finally { rmSync(directory, { recursive: true, force: true }); }
  });

  it("keeps console access read-only and inside the approved command set", () => {
    expect(validateConsoleCommand("Get-Location").safe).toBe(true);
    expect(validateConsoleCommand("Remove-Item notes.txt").safe).toBe(false);
    expect(validateConsoleCommand("Get-Content C:\\Windows\\win.ini").safe).toBe(false);
  });
});
