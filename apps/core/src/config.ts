import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { z } from "zod";
import { type AgentMode, type ProviderAuthInfo, type ProviderConfig, type ProviderId } from "../../../packages/shared/src/index.js";
import { defaultProviderConfigs } from "../../../packages/providers/src/index.js";

const authSchema = z.object({
  method: z.enum(["api-key", "environment", "none", "existing-session", "oauth"]),
  credentialRef: z.string().max(200).optional(),
  oauthCredentialRef: z.string().max(200).optional(),
  oauthProvider: z.string().regex(/^[a-z0-9][a-z0-9._-]{1,60}$/).optional(),
  envVar: z.string().regex(/^[A-Za-z_][A-Za-z0-9_]*$/).optional(),
});

const providerSchema = z.object({
  id: z.string().min(1).max(80),
  displayName: z.string().min(1).max(120),
  kind: z.enum(["openai-compatible", "openwordcode-bridge", "openai-codex", "anthropic", "gemini", "google-antigravity", "demo"]),
  baseUrl: z.string().url(),
  enabled: z.boolean(),
  local: z.boolean(),
  auth: authSchema,
  defaultModel: z.string().max(200).optional(),
  privacyNote: z.string().max(500),
  internal: z.boolean().optional(),
});

const persistedModeSchema = z.enum(["manual", "auto", "skip", "ask", "suggest", "apply-after-approval"]);

function normalizeMode(value: z.infer<typeof persistedModeSchema>): AgentMode {
  if (value === "auto") return "auto";
  if (value === "skip") return "skip";
  return "manual";
}

const settingsSchema = z.object({
  version: z.literal(1),
  selectedProviderId: z.string().min(1),
  selectedModelId: z.string().optional(),
  mode: persistedModeSchema,
  theme: z.enum(["light", "dark", "system"]),
  providers: z.record(providerSchema),
  allowedOrigins: z.array(z.string().url()).max(20),
}).transform(value => ({ ...value, mode: normalizeMode(value.mode) }));

export type Settings = z.infer<typeof settingsSchema>;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function migratePersistedSettings(value: unknown): { value: unknown; changed: boolean } {
  if (!isRecord(value) || !isRecord(value.providers)) return { value, changed: false };
  const providers: Record<string, unknown> = { ...value.providers };
  const legacy = providers["legacy-bridge"];
  const hasLegacy = legacy !== undefined;
  const hasBridge = providers["openwordcode-bridge"] !== undefined;
  let changed = false;
  if (hasLegacy && !hasBridge) {
    const old = isRecord(legacy) ? legacy : {};
    providers["openwordcode-bridge"] = {
      ...old,
      id: "openwordcode-bridge",
      displayName: "OpenWordCode Bridge",
      kind: "openwordcode-bridge",
      baseUrl: "http://127.0.0.1:10101/v1",
      enabled: true,
      local: true,
      auth: { method: "oauth", credentialRef: "provider:openwordcode-account" },
      defaultModel: typeof old.defaultModel === "string" && old.defaultModel ? old.defaultModel : "gpt-5.6-luna",
      privacyNote: "Uses the local OpenWordCode Bridge. Its configured provider/account determines where document content is sent.",
    };
    changed = true;
  }
  if (hasLegacy) {
    delete providers["legacy-bridge"];
    changed = true;
  }
  if (isRecord(providers.demo) && providers.demo.internal !== true) {
    providers.demo = { ...providers.demo, internal: true };
    changed = true;
  }
  // Ollama and LM Studio were removed from the product. Drop them from older
  // persisted configs so they do not reappear in the account list.
  for (const removed of ["ollama", "lm-studio"]) {
    if (providers[removed] !== undefined) {
      delete providers[removed];
      changed = true;
    }
  }
  const rawSelectedProviderId = typeof value.selectedProviderId === "string" ? value.selectedProviderId : "";
  let selectedProviderId = rawSelectedProviderId === "legacy-bridge" || rawSelectedProviderId === "ollama" || rawSelectedProviderId === "lm-studio" || rawSelectedProviderId === "demo"
    ? "openwordcode-bridge"
    : rawSelectedProviderId;
  let selectedModelId = typeof value.selectedModelId === "string" ? value.selectedModelId : undefined;
  if (selectedProviderId !== rawSelectedProviderId) {
    selectedModelId = "gpt-5.6-luna";
    changed = true;
  }

  // The public add-in is account-first. Move API-only selections to the local
  // Bridge so a fresh install never lands on a provider with no public sign-in
  // path. Existing OAuth-capable providers remain selectable.
  const selectedProviderValue: unknown = providers[selectedProviderId];
  const selectedAuth = isRecord(selectedProviderValue) && isRecord(selectedProviderValue.auth) ? selectedProviderValue.auth : undefined;
  const selectedProviderIsPubliclyConnectable = selectedProviderId === "openwordcode-bridge"
    || (isRecord(selectedProviderValue) && (selectedProviderValue.local === true || selectedAuth?.method === "none" || typeof selectedAuth?.oauthProvider === "string"));
  if (!selectedProviderIsPubliclyConnectable) {
    selectedProviderId = "openwordcode-bridge";
    selectedModelId = "gpt-5.6-luna";
    changed = true;
  }
  return { value: { ...value, providers, selectedProviderId, selectedModelId }, changed };
}

export function dataDirectory(env: NodeJS.ProcessEnv = process.env): string {
  const explicit = env.OPENWORDCODE_DATA_DIR?.trim();
  if (explicit) return explicit;
  const base = process.platform === "win32"
    ? env.APPDATA?.trim() || env.USERPROFILE?.trim() || homedir()
    : env.XDG_CONFIG_HOME?.trim() || env.HOME?.trim() || homedir();
  return join(base, "OpenWordCode");
}

export function configPath(env: NodeJS.ProcessEnv = process.env): string { return join(dataDirectory(env), "config.json"); }

export function credentialPath(env: NodeJS.ProcessEnv = process.env): string { return join(dataDirectory(env), "credentials"); }

export function defaultSettings(): Settings {
  const providers = Object.fromEntries(defaultProviderConfigs().map(provider => [provider.id, provider]));
  return {
    version: 1,
    selectedProviderId: "openwordcode-bridge",
    selectedModelId: providers["openwordcode-bridge"]?.defaultModel,
    mode: "manual",
    theme: "system",
    providers,
    allowedOrigins: ["https://localhost:3000", "https://127.0.0.1:3000", "http://localhost:5173", "http://127.0.0.1:5173", "http://localhost:10200", "http://127.0.0.1:10200"],
  };
}

export function loadSettings(env: NodeJS.ProcessEnv = process.env): Settings {
  const defaults = defaultSettings();
  mkdirSync(dataDirectory(env), { recursive: true, mode: 0o700 });
  try {
    const parsed = JSON.parse(readFileSync(configPath(env), "utf8")) as unknown;
    const migration = migratePersistedSettings(parsed);
    const validated = settingsSchema.safeParse(migration.value);
    if (!validated.success) return defaults;
    const providers: Record<string, ProviderConfig> = { ...defaults.providers, ...validated.data.providers };
    let changed = migration.changed;
    // Existing installations keep their provider objects from the persisted
    // config. Add newly shipped OAuth metadata without changing the user's
    // selected auth method, API key reference, endpoint, or model choice.
    for (const [id, defaultProvider] of Object.entries(defaults.providers)) {
      const existing = providers[id];
      if (!existing) continue;
      let next = existing;
      if (defaultProvider.auth.oauthProvider && !existing.auth.oauthProvider) {
        next = {
          ...next,
          auth: {
            ...next.auth,
            oauthProvider: defaultProvider.auth.oauthProvider,
            oauthCredentialRef: next.auth.oauthCredentialRef ?? defaultProvider.auth.oauthCredentialRef,
          },
        };
      }
      // Kimi OAuth is the Kimi Code account surface, whose API endpoint is
      // separate from the Moonshot API-key endpoint used by older configs.
      if (id === "kimi" && existing.baseUrl === "https://api.moonshot.ai/v1") {
        next = { ...next, baseUrl: defaultProvider.baseUrl, displayName: defaultProvider.displayName, defaultModel: defaultProvider.defaultModel };
      }
      if (next !== existing) {
        providers[id] = next;
        changed = true;
      }
    }
    const settings = { ...defaults, ...validated.data, providers };
    if (changed) saveSettings(settings, env);
    return settings;
  } catch {
    return defaults;
  }
}

export function saveSettings(settings: Settings, env: NodeJS.ProcessEnv = process.env): void {
  const validated = settingsSchema.parse(settings);
  const path = configPath(env);
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 });
  const temporary = `${path}.${process.pid}.tmp`;
  writeFileSync(temporary, `${JSON.stringify(validated, null, 2)}\n`, { encoding: "utf8", mode: 0o600 });
  renameSync(temporary, path);
}

export function updateProvider(settings: Settings, provider: ProviderConfig): Settings {
  return { ...settings, providers: { ...settings.providers, [provider.id]: provider } };
}

export function providerConfig(settings: Settings, id: ProviderId): ProviderConfig {
  const provider = settings.providers[id];
  if (!provider) throw new Error(`Unknown provider: ${id}`);
  return provider;
}

export function publicProviderConfig(provider: ProviderConfig, auth: ProviderAuthInfo): ProviderConfig & { auth: ProviderAuthInfo } {
  return { ...provider, auth };
}
