import { type CredentialStore } from "../../../packages/auth/src/index.js";
import { type AuthMethod, type ProviderAuthInfo, type ProviderConfig } from "../../../packages/shared/src/index.js";
import { type ProviderOAuthCredential } from "../../../packages/providers/src/index.js";
import { type ChatGPTOAuthStatus, ChatGPTOAuthManager, OPENWORDCODE_ACCOUNT_CREDENTIAL_REF } from "./chatgpt-auth.js";
import { oauthProviderIsSupported, ProviderOAuthManager, type OAuthStatus } from "./provider-oauth.js";

/**
 * Keeps provider configuration separate from secrets. OAuth access/refresh
 * tokens are owned by ProviderOAuthManager; this class only selects the
 * authentication method used by a provider runtime.
 */
export class AuthManager {
  constructor(
    private readonly store: CredentialStore,
    private readonly env: NodeJS.ProcessEnv = process.env,
    private readonly chatGptAuth?: ChatGPTOAuthManager,
    private readonly providerOAuth?: ProviderOAuthManager,
  ) {}

  async status(provider: ProviderConfig): Promise<ProviderAuthInfo> {
    const stored = Boolean(provider.auth.credentialRef && await this.store.get(provider.auth.credentialRef));
    const envConfigured = Boolean(provider.auth.envVar && this.env[provider.auth.envVar]?.trim());
    const availableMethods = this.availableMethods(provider);

    if (provider.auth.method === "none") {
      return {
        status: provider.kind === "demo" ? "connected" : "not-configured",
        method: provider.auth.method,
        detail: provider.kind === "demo" ? "Offline demo ready" : "No cloud credential is required; test the endpoint to verify it is running.",
        availableMethods,
        credentialConfigured: stored,
        environmentConfigured: envConfigured,
      };
    }
    if (provider.auth.method === "environment") {
      return {
        status: envConfigured ? "connected" : "login-required",
        method: provider.auth.method,
        detail: envConfigured ? `Using ${provider.auth.envVar} from the Core environment` : `Set ${provider.auth.envVar} or choose another sign-in method`,
        availableMethods,
        credentialConfigured: stored,
        environmentConfigured: envConfigured,
      };
    }
    if (provider.auth.method === "api-key") {
      return {
        status: stored ? "connected" : "login-required",
        method: provider.auth.method,
        detail: stored ? "Legacy provider credential is configured" : "This provider needs a supported account connection or local runtime",
        availableMethods,
        credentialConfigured: stored,
        environmentConfigured: envConfigured,
      };
    }
    if (provider.auth.method === "oauth") {
      if (isOpenWordCodeAccountProvider(provider) && this.chatGptAuth) {
        const oauth = await this.chatGptAuth.status(provider.auth.credentialRef ?? OPENWORDCODE_ACCOUNT_CREDENTIAL_REF);
        return oauthAuthInfo(oauth, availableMethods);
      }
      if (this.providerOAuth && provider.auth.oauthProvider) {
        return oauthAuthInfo(await this.providerOAuth.status(provider), availableMethods);
      }
      return {
        status: "unsupported",
        method: provider.auth.method,
        detail: "This provider does not have an enabled OAuth adapter in OpenWordCode.",
        availableMethods,
        credentialConfigured: false,
        environmentConfigured: envConfigured,
      };
    }
    return {
      status: "unsupported",
      method: provider.auth.method,
      detail: "This authentication mode is not supported by this provider.",
      availableMethods,
      credentialConfigured: stored,
      environmentConfigured: envConfigured,
    };
  }

  async configureApiKey(provider: ProviderConfig, key: string, _label?: string): Promise<ProviderConfig> {
    const credentialRef = provider.auth.credentialRef ?? `provider:${provider.id}`;
    await this.store.set(credentialRef, key.trim());
    if (provider.auth.oauthProvider) await this.providerOAuth?.clearLocalCli(provider.auth.oauthProvider);
    return {
      ...provider,
      auth: {
        ...provider.auth,
        method: "api-key",
        credentialRef,
      },
    };
  }

  async disconnect(provider: ProviderConfig): Promise<ProviderConfig> {
    if (provider.auth.method === "oauth") {
      if (isOpenWordCodeAccountProvider(provider) && this.chatGptAuth) {
        await this.chatGptAuth.disconnect(provider.auth.credentialRef ?? OPENWORDCODE_ACCOUNT_CREDENTIAL_REF);
      } else if (this.providerOAuth && provider.auth.oauthProvider) {
        await this.providerOAuth.disconnect(provider);
      }
    } else if (provider.auth.method === "api-key" && provider.auth.credentialRef) {
      await this.store.remove(provider.auth.credentialRef);
    }

    const envVar = provider.auth.envVar ?? defaultEnvironmentVariable(provider.id);
    if (isOpenWordCodeAccountProvider(provider)) {
      return { ...provider, auth: { ...provider.auth, method: "oauth", credentialRef: provider.auth.credentialRef ?? OPENWORDCODE_ACCOUNT_CREDENTIAL_REF } };
    }

    // OAuth-only providers remain on their OAuth method after disconnect so
    // the sign-in card is immediately available again. Providers with an API
    // key/environment fallback switch back to that fallback.
    if (provider.auth.oauthProvider && !provider.auth.envVar && !provider.auth.credentialRef && oauthProviderIsSupported(provider.auth.oauthProvider)) {
      return { ...provider, auth: { ...provider.auth, method: "oauth", oauthCredentialRef: provider.auth.oauthCredentialRef ?? `oauth:${provider.auth.oauthProvider}` } };
    }
    if (envVar) return { ...provider, auth: { ...provider.auth, method: "environment", envVar, credentialRef: undefined } };
    if (provider.local) return { ...provider, auth: { ...provider.auth, method: "none", credentialRef: undefined } };
    return { ...provider, auth: { ...provider.auth, method: "api-key", credentialRef: provider.auth.credentialRef ?? `provider:${provider.id}` } };
  }

  setMethod(provider: ProviderConfig, method: AuthMethod, envVar?: string): ProviderConfig {
    if (!this.availableMethods(provider).includes(method)) throw new Error(`${provider.displayName} does not support ${method}`);
    if (method === "environment") {
      return { ...provider, auth: { ...provider.auth, method, envVar: envVar?.trim() || provider.auth.envVar || defaultEnvironmentVariable(provider.id), credentialRef: undefined } };
    }
    if (method === "api-key") {
      return { ...provider, auth: { ...provider.auth, method, credentialRef: provider.auth.credentialRef ?? `provider:${provider.id}` } };
    }
    if (method === "none") return { ...provider, auth: { ...provider.auth, method, credentialRef: undefined } };
    if (method === "oauth") {
      if (isOpenWordCodeAccountProvider(provider)) {
        return { ...provider, auth: { ...provider.auth, method, credentialRef: provider.auth.credentialRef ?? OPENWORDCODE_ACCOUNT_CREDENTIAL_REF } };
      }
      if (!provider.auth.oauthProvider || !oauthProviderIsSupported(provider.auth.oauthProvider)) throw new Error(`${provider.displayName} OAuth is not enabled in this build`);
      return { ...provider, auth: { ...provider.auth, method, oauthCredentialRef: provider.auth.oauthCredentialRef ?? `oauth:${provider.auth.oauthProvider}`, credentialRef: undefined } };
    }
    return { ...provider, auth: { ...provider.auth, method } };
  }

  /** Switches a provider to the OAuth credential after a completed flow. */
  activateOAuth(provider: ProviderConfig): ProviderConfig {
    if (isOpenWordCodeAccountProvider(provider)) {
      return { ...provider, auth: { ...provider.auth, method: "oauth", credentialRef: provider.auth.credentialRef ?? OPENWORDCODE_ACCOUNT_CREDENTIAL_REF } };
    }
    if (!provider.auth.oauthProvider || !oauthProviderIsSupported(provider.auth.oauthProvider)) throw new Error(`${provider.displayName} OAuth is not enabled in this build`);
    return { ...provider, auth: { ...provider.auth, method: "oauth", oauthCredentialRef: provider.auth.oauthCredentialRef ?? `oauth:${provider.auth.oauthProvider}`, credentialRef: undefined } };
  }

  async resolveSecret(provider: ProviderConfig): Promise<string | null> {
    if (provider.auth.method === "environment") return provider.auth.envVar ? this.env[provider.auth.envVar]?.trim() || null : null;
    if (provider.auth.method === "api-key" && provider.auth.credentialRef) return this.store.get(provider.auth.credentialRef);
    return null;
  }

  async resolveOAuth(provider: ProviderConfig): Promise<ProviderOAuthCredential | null> {
    if (provider.auth.method !== "oauth") return null;
    if (isOpenWordCodeAccountProvider(provider)) return this.chatGptAuth?.resolve(provider.auth.credentialRef ?? OPENWORDCODE_ACCOUNT_CREDENTIAL_REF) ?? null;
    if (!this.providerOAuth) return null;
    return this.providerOAuth.resolve(provider);
  }

  availableMethods(provider: ProviderConfig): AuthMethod[] {
    if (isOpenWordCodeAccountProvider(provider)) return ["oauth"];
    if (provider.kind === "demo") return ["none"];
    if (provider.local) return ["none"];
    return provider.auth.oauthProvider && oauthProviderIsSupported(provider.auth.oauthProvider) ? ["oauth"] : [];
  }
}

function isOpenWordCodeAccountProvider(provider: ProviderConfig): boolean {
  return provider.kind === "openwordcode-bridge" || provider.id === "openwordcode-account";
}

function oauthAuthInfo(oauth: ChatGPTOAuthStatus | OAuthStatus, availableMethods: AuthMethod[]): ProviderAuthInfo {
  return {
    status: oauth.status,
    method: "oauth",
    detail: oauth.detail,
    availableMethods,
    credentialConfigured: oauth.credentialConfigured,
    environmentConfigured: false,
    ...(oauth.source ? { source: oauth.source } : {}),
  };
}

function defaultEnvironmentVariable(providerId: string): string | undefined {
  const values: Record<string, string> = {
    openai: "OPENAI_API_KEY",
    anthropic: "ANTHROPIC_API_KEY",
    openrouter: "OPENROUTER_API_KEY",
    google: "GEMINI_API_KEY",
    xai: "XAI_API_KEY",
    kimi: "MOONSHOT_API_KEY",
  };
  return values[providerId];
}
