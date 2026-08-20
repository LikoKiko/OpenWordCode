import { useCallback, useEffect, useMemo, useRef, useState, type ChangeEvent, type JSX } from "react";
import {
  ArrowClockwise,
  ArrowLeft,
  ArrowUp,
  CaretRight,
  ClockCounterClockwise,
  GearSix,
  Globe,
  Paperclip,
  Plus,
  Sparkle,
  TerminalWindow,
  X,
} from "@phosphor-icons/react";
import type { AgentAction, AgentEvent, AgentMode, ChatAttachment, DocumentSnapshot, ModelInfo, ProposedChange, ProviderSummary, SkillSummary } from "../../../packages/shared/src/index.js";
import { comparableText, visualElementTargetText } from "../../../packages/shared/src/index.js";
import { createWordAdapter, isOfficeHost, waitForOfficeReady, type WordApplicationAdapter } from "../../../packages/app-word/src/index.js";
import { AppliedEditList, AttachmentList, ConsoleActionCard, DEFAULT_SKILLS, ModePicker, ModelPicker, parseSkillFile, RecentTasksDrawer, SearchList, SkillsDrawer, ThinkingTrace, type AppliedEdit, type AttachmentPreview, type ModelEffort, type RecentTask, type SearchItem } from "./components";
import {
  approveChange,
  approveConsoleAction,
  completeChange,
  disconnectChatGPT,
  getBridgeStatus,
  getChatGPTLoginStatus,
  getHealth,
  getModels,
  getOAuthLoginStatus,
  getProviders,
  getSettings,
  initializeCore,
  rejectConsoleAction,
  saveSettings,
  startOAuthLogin,
  startChatGPTLogin,
  startCodexCliLogin,
  startLocalCliLogin,
  useLocalCliSession,
  useCodexCliSession,
  streamAgent,
  cancelOAuthLogin,
  completeOAuthLogin,
  disconnectOAuth,
} from "./api";
import type { OAuthLoginStart } from "./api";
import { dismissNotifications, notifyError, notifySuccess } from "./notifications";

type Tab = "chat" | "settings";
type UiMessage = { id: string; role: "user" | "assistant"; content: string; toolActivity?: string; attachments?: AttachmentPreview[]; actions?: AgentAction[]; edits?: AppliedEdit[] };
type TaskSession = { id: string; title: string; createdAt: string; updatedAt: string; messages: UiMessage[] };

const MAX_ATTACHMENTS = 4;
const MAX_ATTACHMENT_BYTES = 6_000_000;
const MAX_TOTAL_ATTACHMENT_BYTES = 12_000_000;
const TASK_HISTORY_STORAGE_KEY = "openwordcode.task-history.v1";
const SKILLS_STORAGE_KEY = "openwordcode.skills.v1";
const MAX_TASK_HISTORY = 20;
const SUPPORTED_IMAGE_TYPES = new Set(["image/gif", "image/jpeg", "image/png", "image/webp"]);
function loadStoredSkills(): SkillSummary[] {
  if (typeof window === "undefined") return DEFAULT_SKILLS;
  try {
    const raw = window.localStorage.getItem(SKILLS_STORAGE_KEY);
    if (!raw) return DEFAULT_SKILLS;
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed) || !parsed.length) return DEFAULT_SKILLS;
    const custom = parsed.filter(item => isRecord(item) && !item.isDefault) as SkillSummary[];
    const defaults = DEFAULT_SKILLS.map(def => {
      const saved = parsed.find(item => isRecord(item) && item.id === def.id) as SkillSummary | undefined;
      return saved ? { ...def, enabled: saved.enabled !== false } : def;
    });
    return [...defaults, ...custom];
  } catch {
    return DEFAULT_SKILLS;
  }
}

function saveStoredSkills(skills: SkillSummary[]): void {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(SKILLS_STORAGE_KEY, JSON.stringify(skills));
  } catch {
    // ignore
  }
}

function fileMimeType(file: File): ChatAttachment["mimeType"] | null {
  if (file.type === "application/pdf" || SUPPORTED_IMAGE_TYPES.has(file.type)) return file.type as ChatAttachment["mimeType"];
  if (file.name.toLocaleLowerCase().endsWith(".pdf")) return "application/pdf";
  return null;
}

function readDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => typeof reader.result === "string" ? resolve(reader.result) : reject(new Error(`Could not read ${file.name}`));
    reader.onerror = () => reject(reader.error ?? new Error(`Could not read ${file.name}`));
    reader.readAsDataURL(file);
  });
}

function uid(): string {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function taskTitle(messages: UiMessage[]): string {
  const firstUser = messages.find(message => message.role === "user");
  const normalized = firstUser?.content.replace(/\s+/g, " ").trim() ?? "";
  if (!normalized) return "New task";
  return normalized.length > 52 ? `${normalized.slice(0, 52).trimEnd()}…` : normalized;
}

function storedMessages(messages: UiMessage[]): UiMessage[] {
  return messages.map(message => ({
    id: message.id,
    role: message.role,
    content: message.content,
    ...(message.toolActivity ? { toolActivity: message.toolActivity } : {}),
    ...(message.attachments?.length ? {
      attachments: message.attachments.map(({ id, name, mimeType, size }) => ({ id, name, mimeType, size })),
    } : {}),
    ...(message.actions?.length ? { actions: message.actions } : {}),
    ...(message.edits?.length ? { edits: message.edits } : {}),
  }));
}

function parseStoredAttachment(value: unknown): AttachmentPreview | null {
  if (!isRecord(value) || typeof value.id !== "string" || typeof value.name !== "string" || typeof value.size !== "number") return null;
  if (value.mimeType !== "application/pdf" && value.mimeType !== "image/gif" && value.mimeType !== "image/jpeg" && value.mimeType !== "image/png" && value.mimeType !== "image/webp") return null;
  return { id: value.id, name: value.name, mimeType: value.mimeType, size: value.size };
}

function parseStoredMessage(value: unknown): UiMessage | null {
  if (!isRecord(value) || typeof value.id !== "string" || (value.role !== "user" && value.role !== "assistant") || typeof value.content !== "string") return null;
  // Older task history stored the pre-rebase failure as assistant text. It is
  // not a real response and should not keep resurfacing after the add-in is
  // upgraded. Current failures are reported by the live error banner instead.
  const content = value.content
    .replace(/\s*I couldn’t apply this edit in Word:\s*The target paragraph changed after this suggestion was created\.?/giu, "")
    .replace(/\s*The target paragraph changed after this suggestion was created\.?/giu, "")
    .trim();
  if (value.role === "assistant" && !content) return null;
  const attachments = Array.isArray(value.attachments) ? value.attachments.map(parseStoredAttachment).filter((attachment): attachment is AttachmentPreview => Boolean(attachment)) : [];
  const edits = Array.isArray(value.edits) ? value.edits.flatMap(item => {
    if (!isRecord(item) || typeof item.id !== "string" || typeof item.description !== "string") return [];
    return [{ id: item.id, description: item.description, ...(typeof item.before === "string" ? { before: item.before } : {}), ...(typeof item.after === "string" ? { after: item.after } : {}) }];
  }) : [];
  return {
    id: value.id,
    role: value.role,
    content,
    ...(typeof value.toolActivity === "string" ? { toolActivity: value.toolActivity } : {}),
    ...(attachments.length ? { attachments } : {}),
    ...(Array.isArray(value.actions) ? { actions: value.actions.flatMap(item => isRecord(item) && item.type === "console" && typeof item.id === "string" && typeof item.command === "string" && typeof item.workingDirectory === "string" && typeof item.reason === "string" && typeof item.status === "string" && typeof item.createdAt === "string" ? [item as unknown as AgentAction] : []) } : {}),
    ...(edits.length ? { edits } : {}),
  };
}

function loadTaskSessions(): TaskSession[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(TASK_HISTORY_STORAGE_KEY);
    if (!raw) return [];
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.flatMap(value => {
      if (!isRecord(value) || typeof value.id !== "string" || typeof value.title !== "string") return [];
      const messages = Array.isArray(value.messages) ? value.messages.map(parseStoredMessage).filter((message): message is UiMessage => Boolean(message)) : [];
      if (!messages.length) return [];
      const now = new Date().toISOString();
      return [{
        id: value.id,
        title: value.title || taskTitle(messages),
        createdAt: typeof value.createdAt === "string" ? value.createdAt : now,
        updatedAt: typeof value.updatedAt === "string" ? value.updatedAt : now,
        messages,
      }];
    }).sort((left, right) => right.updatedAt.localeCompare(left.updatedAt)).slice(0, MAX_TASK_HISTORY);
  } catch {
    return [];
  }
}

function upsertTask(sessions: TaskSession[], id: string, messages: UiMessage[]): TaskSession[] {
  if (!messages.length) return sessions;
  const existing = sessions.find(session => session.id === id);
  const now = new Date().toISOString();
  const next: TaskSession = {
    id,
    title: taskTitle(messages),
    createdAt: existing?.createdAt ?? now,
    updatedAt: now,
    messages: storedMessages(messages),
  };
  return [next, ...sessions.filter(session => session.id !== id)].slice(0, MAX_TASK_HISTORY);
}

function ChatGPTAccountCard({ provider, busy, codexLoginStarted, onLogin, onStartCodexCliLogin, onUseCodexCli, onDisconnect }: { provider: ProviderSummary; busy: boolean; codexLoginStarted?: boolean; onLogin: () => void; onStartCodexCliLogin: () => void; onUseCodexCli: () => void; onDisconnect: () => void }): JSX.Element {
  const connected = provider.auth.status === "connected";
  const isCodexCli = provider.auth.source === "codex-cli";
  return <div className="account-detail">
    <p className="account-detail-text">{connected ? provider.auth.detail : isCodexCli ? "Connected via your local Codex CLI session." : "Sign in with your ChatGPT account in your browser or connect your local Codex CLI session."}</p>
    <div className="account-actions">{connected ? <button className="text-button" type="button" onClick={onDisconnect}>Disconnect</button> : <><button className="apply-button" type="button" onClick={onLogin} disabled={busy}>{busy ? "Opening browser…" : "Sign in"}</button><button className="text-button" type="button" onClick={onUseCodexCli} disabled={busy}>{codexLoginStarted ? "Connect after login" : "Use Codex CLI"}</button>{onStartCodexCliLogin ? <button className="text-button" type="button" onClick={onStartCodexCliLogin} disabled={busy}>{codexLoginStarted ? "Waiting for login…" : "CLI login"}</button> : null}</>}</div>
  </div>;
}

function localCliForProvider(providerId: string): { source: "claude-cli" | "kimi-cli" | "antigravity-cli"; label: string } | null {
  if (providerId === "anthropic") return { source: "claude-cli", label: "Claude Code" };
  if (providerId === "kimi") return { source: "kimi-cli", label: "Kimi CLI" };
  if (providerId === "google-antigravity") return { source: "antigravity-cli", label: "Antigravity CLI" };
  return null;
}

function accountSubline(provider: ProviderSummary): string {
  if (provider.auth.status === "connected") return "Connected";
  if (provider.auth.status === "unsupported" && localCliForProvider(provider.id)) return "CLI connection available";
  if (provider.auth.status === "expired") return "Session expired";
  return provider.auth.status === "login-required" ? "Not connected" : provider.auth.status.replace(/-/g, " ");
}

function OAuthAccountCard({ provider, flow, busy, localCliLoginStarted, manualCode, onManualCodeChange, onCompleteManualCode, onLogin, onOpenBrowser, onCancel, onStartLocalCliLogin, onUseLocalCli, onDisconnect }: { provider: ProviderSummary; flow: OAuthLoginStart | null; busy: boolean; localCliLoginStarted?: boolean; manualCode?: string; onManualCodeChange?: (value: string) => void; onCompleteManualCode?: () => void; onLogin: () => void; onOpenBrowser: () => void; onCancel: () => void; onStartLocalCliLogin?: () => void; onUseLocalCli?: () => void; onDisconnect: () => void }): JSX.Element {
  const flowForProvider = flow?.providerId === provider.id ? flow : null;
  const localCli = localCliForProvider(provider.id);
  const connected = provider.auth.method === "oauth" && provider.auth.status === "connected";
  const unsupported = provider.auth.status === "unsupported" && provider.auth.method === "oauth";
  const usingLocalCli = localCli !== null && provider.auth.source === localCli.source;
  return <div className="account-detail">
    <p className="account-detail-text">{connected ? provider.auth.detail : usingLocalCli ? `Using ${localCli?.label} session` : "Sign in opens your browser. OpenWordCode stores credentials in your local encrypted store."}</p>
    {flowForProvider?.userCode ? <div className="oauth-code"><span>Code</span><code>{flowForProvider.userCode}</code></div> : null}
    {flowForProvider?.providerId === "xai" ? <div className="oauth-manual-code"><input value={manualCode ?? ""} onChange={event => onManualCodeChange?.(event.target.value)} placeholder="Paste the xAI code" autoComplete="off" spellCheck={false} /><button className="apply-button" type="button" onClick={onCompleteManualCode} disabled={!manualCode?.trim() || busy}>{busy ? "Checking…" : "Finish"}</button></div> : null}
    {flowForProvider?.detail ? <p className="account-detail-note">{flowForProvider.detail}</p> : null}
    <div className="account-actions">
      {connected ? <button className="text-button" type="button" onClick={onDisconnect}>Disconnect</button> : flowForProvider ? <>{flowForProvider.authorizeUrl || flowForProvider.verificationUrl ? <button className="apply-button" type="button" onClick={onOpenBrowser}>Open sign-in</button> : null}<button className="text-button" type="button" onClick={onCancel}>Cancel</button></> : <><button className="apply-button" type="button" onClick={onLogin} disabled={busy || unsupported}>{busy ? "Opening browser…" : "Sign in"}</button>{localCli && onUseLocalCli ? <button className="text-button" type="button" onClick={onUseLocalCli} disabled={busy}>{localCliLoginStarted ? "Connect after login" : `Use ${localCli.label}`}</button> : null}{localCli && onStartLocalCliLogin ? <button className="text-button" type="button" onClick={onStartLocalCliLogin} disabled={busy}>{localCliLoginStarted ? `Waiting for ${localCli.label}…` : `Launch CLI`}</button> : null}</>}
    </div>
  </div>;
}

export default function App(): JSX.Element {
  const [adapter, setAdapter] = useState<WordApplicationAdapter>(() => createWordAdapter());
  const [wordReady, setWordReady] = useState(() => !isOfficeHost());
  const [, setSnapshot] = useState<DocumentSnapshot | null>(null);
  const [coreOnline, setCoreOnline] = useState(false);
  const [coreVersion, setCoreVersion] = useState("—");
  const [providers, setProviders] = useState<ProviderSummary[]>([]);
  const [models, setModels] = useState<ModelInfo[]>([]);
  const [selectedProviderId, setSelectedProviderId] = useState("openwordcode-bridge");
  const [selectedModelId, setSelectedModelId] = useState("");
  const [modelEffort, setModelEffort] = useState<ModelEffort>("high");
  const [mode, setMode] = useState<AgentMode>("manual");
  const [theme, setTheme] = useState<"light" | "dark" | "system">("system");
  const [tab, setTab] = useState<Tab>("chat");
  const [prompt, setPrompt] = useState("");
  const [taskSessions, setTaskSessions] = useState<TaskSession[]>(loadTaskSessions);
  const [activeTaskId, setActiveTaskId] = useState(() => taskSessions[0]?.id ?? uid());
  const [recentTasksOpen, setRecentTasksOpen] = useState(false);
  const [skills, setSkills] = useState<SkillSummary[]>(loadStoredSkills);
  const [skillsDrawerOpen, setSkillsDrawerOpen] = useState(false);
  const [messages, setMessages] = useState<UiMessage[]>(() => taskSessions[0]?.messages ?? []);
  const [busy, setBusy] = useState(false);
  const [loadingModels, setLoadingModels] = useState(false);
  const setError = useCallback((message: string): void => {
    if (message) notifyError(message);
    else dismissNotifications();
  }, []);
  const [bridge, setBridge] = useState<{ available: boolean; endpoint: string; models: number; detail: string } | null>(null);
  const [actingAction, setActingAction] = useState<string | null>(null);
  const [abortController, setAbortController] = useState<AbortController | null>(null);
  const [composerToolsOpen, setComposerToolsOpen] = useState(false);
  const [attachments, setAttachments] = useState<ChatAttachment[]>([]);
  const [webSearchEnabled, setWebSearchEnabled] = useState(false);
  const [consoleEnabled, setConsoleEnabled] = useState(false);
  const [uploadingFiles, setUploadingFiles] = useState(false);
  const [chatGptFlowId, setChatGptFlowId] = useState<string | null>(null);
  const [chatGptLoginBusy, setChatGptLoginBusy] = useState(false);
  const [codexCliLoginStarted, setCodexCliLoginStarted] = useState(false);
  const [oauthFlowId, setOAuthFlowId] = useState<string | null>(null);
  const [oauthLoginInfo, setOAuthLoginInfo] = useState<OAuthLoginStart | null>(null);
  const [oauthManualCode, setOAuthManualCode] = useState("");
  const [oauthLoginBusy, setOAuthLoginBusy] = useState(false);
  const [localCliLoginProviderId, setLocalCliLoginProviderId] = useState<string | null>(null);
  const [approvalNoticeDismissed, setApprovalNoticeDismissed] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const autoAppliedChanges = useRef(new Set<string>());
  const applyQueue = useRef(Promise.resolve());

  const selectedProvider = useMemo(() => providers.find(provider => provider.id === selectedProviderId), [providers, selectedProviderId]);
  const visibleProviders = useMemo(() => providers.filter(provider => !provider.internal && provider.id !== "demo" && (provider.id === "openwordcode-bridge" || provider.auth.availableMethods.includes("oauth") || provider.auth.status === "connected")), [providers]);
  const connectedProviderCount = useMemo(() => visibleProviders.filter(provider => provider.auth.status === "connected").length, [visibleProviders]);
  const activeSkills = useMemo(() => skills.filter(skill => skill.enabled !== false), [skills]);
  const selectedModel = useMemo(() => models.find(model => model.id === selectedModelId), [models, selectedModelId]);
  const modelLabel = selectedModel?.name || selectedModelId || (loadingModels ? "Loading model…" : "Choose model");

  const refreshSnapshot = useCallback(async (): Promise<DocumentSnapshot> => {
    const next = await adapter.readSnapshot();
    setSnapshot(next);
    return next;
  }, [adapter]);

  const loadModels = useCallback(async (providerId: string, preferredModel?: string, fallbackModel?: string): Promise<void> => {
    setLoadingModels(true);
    try {
      const next = await getModels(providerId);
      setModels(next);
      const selected = next.find(model => model.id === preferredModel)?.id ?? next[0]?.id ?? fallbackModel ?? "";
      setSelectedModelId(selected);
      if (selected) await saveSettings({ selectedProviderId: providerId, selectedModelId: selected });
    } catch (cause) {
      setModels([]);
      const detail = cause instanceof Error ? cause.message : "Model discovery failed";
      // A missing key is the expected first-run state now that users connect
      // their own provider account from the task pane. Keep the connect card
      // usable without covering it with a startup error banner.
      if (!/api key|environment variable|credential|authentication/iu.test(detail)) setError(detail);
    } finally {
      setLoadingModels(false);
    }
  }, []);

  const loadCore = useCallback(async (): Promise<void> => {
    setError("");
    try {
      const bootstrap = await initializeCore();
      const health = await getHealth();
      const [nextProviders, nextSettings] = await Promise.all([getProviders(), getSettings()]);
      setCoreOnline(health.status === "ok");
      setCoreVersion(bootstrap.version || health.version);
      setProviders(nextProviders);
      setSelectedProviderId(nextSettings.selectedProviderId);
      setMode(nextSettings.mode);
      setTheme(nextSettings.theme);
      const fallbackModel = nextProviders.find(provider => provider.id === nextSettings.selectedProviderId)?.defaultModel;
      await loadModels(nextSettings.selectedProviderId, nextSettings.selectedModelId, fallbackModel);
      setBridge(await getBridgeStatus());
    } catch (cause) {
      setCoreOnline(false);
      setError(cause instanceof Error ? cause.message : "OpenWordCode Core is offline");
    }
  }, [loadModels]);

  useEffect(() => { void loadCore(); }, [loadCore]);
  useEffect(() => {
    if (!chatGptFlowId) return;
    let disposed = false;
    let timer: number | undefined;
    const poll = async (): Promise<void> => {
      try {
        const result = await getChatGPTLoginStatus(chatGptFlowId);
        if (disposed) return;
        if (result.status === "connected") {
          setChatGptFlowId(null);
          setChatGptLoginBusy(false);
          await saveSettings({ selectedProviderId: "openwordcode-bridge", selectedModelId: "gpt-5.6-luna" });
          await loadCore();
          return;
        }
        if (result.status === "error") {
          setChatGptFlowId(null);
          setChatGptLoginBusy(false);
          setError(result.detail);
          return;
        }
        timer = window.setTimeout(() => { void poll(); }, 1_000);
      } catch (cause) {
        if (!disposed) {
          setChatGptFlowId(null);
          setChatGptLoginBusy(false);
          setError(cause instanceof Error ? cause.message : "Could not read ChatGPT sign-in status");
        }
      }
    };
    void poll();
    return () => { disposed = true; if (timer !== undefined) window.clearTimeout(timer); };
  }, [chatGptFlowId, loadCore]);
  useEffect(() => {
    if (!codexCliLoginStarted) return;
    let disposed = false;
    let timer: number | undefined;
    const poll = async (): Promise<void> => {
      try {
        const nextProviders = await getProviders();
        if (disposed) return;
        const provider = nextProviders.find(item => item.kind === "openwordcode-bridge");
        if (provider?.auth.source === "codex-cli" && provider.auth.status === "connected") {
          setCodexCliLoginStarted(false);
          setChatGptLoginBusy(false);
          await saveSettings({ selectedProviderId: "openwordcode-bridge", selectedModelId: "gpt-5.6-luna" });
          await loadCore();
          return;
        }
        timer = window.setTimeout(() => { void poll(); }, 1_500);
      } catch {
        if (!disposed) timer = window.setTimeout(() => { void poll(); }, 2_000);
      }
    };
    void poll();
    return () => { disposed = true; if (timer !== undefined) window.clearTimeout(timer); };
  }, [codexCliLoginStarted, loadCore]);
  useEffect(() => {
    if (!localCliLoginProviderId) return;
    let disposed = false;
    let timer: number | undefined;
    const poll = async (): Promise<void> => {
      try {
        const nextProviders = await getProviders();
        if (disposed) return;
        const provider = nextProviders.find(item => item.id === localCliLoginProviderId);
        const source = localCliForProvider(localCliLoginProviderId)?.source;
        if (provider && source && provider.auth.source === source && provider.auth.status === "connected") {
          setLocalCliLoginProviderId(null);
          setOAuthLoginBusy(false);
          await loadCore();
          return;
        }
        timer = window.setTimeout(() => { void poll(); }, 1_500);
      } catch {
        if (!disposed) timer = window.setTimeout(() => { void poll(); }, 2_000);
      }
    };
    void poll();
    return () => { disposed = true; if (timer !== undefined) window.clearTimeout(timer); };
  }, [localCliLoginProviderId, loadCore]);
  useEffect(() => {
    if (!oauthFlowId) return;
    let disposed = false;
    let timer: number | undefined;
    const poll = async (): Promise<void> => {
      try {
        const result = await getOAuthLoginStatus(oauthFlowId);
        if (disposed) return;
        setOAuthLoginInfo(result);
        if (result.status === "connected") {
          setOAuthFlowId(null);
          setOAuthLoginBusy(false);
          setOAuthManualCode("");
          await loadCore();
          return;
        }
        if (result.status === "error" || result.status === "cancelled") {
          setOAuthFlowId(null);
          setOAuthLoginBusy(false);
          setOAuthManualCode("");
          setError(result.detail);
          return;
        }
        timer = window.setTimeout(() => { void poll(); }, 1_000);
      } catch (cause) {
        if (!disposed) {
          setOAuthFlowId(null);
          setOAuthLoginBusy(false);
          setError(cause instanceof Error ? cause.message : "Could not read OAuth sign-in status");
        }
      }
    };
    void poll();
    return () => { disposed = true; if (timer !== undefined) window.clearTimeout(timer); };
  }, [oauthFlowId, loadCore]);
  useEffect(() => {
    let disposed = false;
    if (!isOfficeHost()) {
      setWordReady(true);
      void refreshSnapshot().catch(cause => setError(cause instanceof Error ? cause.message : "Preview document is unavailable"));
      return () => { disposed = true; };
    }
    setWordReady(false);
    void waitForOfficeReady().then(ready => {
      if (disposed) return;
      if (!ready) {
        setError("Word is still starting. Close and reopen the OpenWordCode pane if Office.js did not finish loading.");
        return;
      }
      const next = createWordAdapter();
      setAdapter(next);
      setWordReady(true);
      void next.readSnapshot().then(setSnapshot).catch(cause => setError(cause instanceof Error ? cause.message : "Word could not provide the active document"));
    });
    return () => { disposed = true; };
  }, []);
  useEffect(() => {
    const media = window.matchMedia("(prefers-color-scheme: dark)");
    const applyTheme = (): void => {
      document.documentElement.dataset.theme = theme === "dark" || (theme === "system" && media.matches) ? "dark" : "light";
    };
    applyTheme();
    if (theme !== "system") return;
    media.addEventListener("change", applyTheme);
    return () => media.removeEventListener("change", applyTheme);
  }, [theme]);
  useEffect(() => { setApprovalNoticeDismissed(false); }, [mode]);
  useEffect(() => {
    const onShortcut = (event: KeyboardEvent) => {
      if (event.ctrlKey && event.key.toLocaleLowerCase() === "u") {
        event.preventDefault();
        fileInputRef.current?.click();
      }
    };
    window.addEventListener("keydown", onShortcut);
    return () => window.removeEventListener("keydown", onShortcut);
  }, []);
  useEffect(() => {
    try {
      window.localStorage.setItem(TASK_HISTORY_STORAGE_KEY, JSON.stringify(taskSessions));
    } catch {
      // Local history is a convenience; a blocked or full storage area should not break chat.
    }
  }, [taskSessions]);
  useEffect(() => {
    if (!messages.length) return;
    const timer = window.setTimeout(() => {
      setTaskSessions(previous => upsertTask(previous, activeTaskId, messages));
    }, 350);
    return () => window.clearTimeout(timer);
  }, [activeTaskId, messages]);

  const onFilesSelected = useCallback(async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    const input = event.currentTarget;
    const selectedFiles = Array.from(input.files ?? []);
    input.value = "";
    if (!selectedFiles.length) return;
    if (attachments.length + selectedFiles.length > MAX_ATTACHMENTS) {
      setError(`You can attach up to ${MAX_ATTACHMENTS} files at a time.`);
      return;
    }
    const supported = selectedFiles.flatMap(file => {
      const mimeType = fileMimeType(file);
      if (!mimeType) {
        setError(`${file.name} is not supported. Use an image or PDF.`);
        return [];
      }
      if (file.size > MAX_ATTACHMENT_BYTES) {
        setError(`${file.name} is too large. The limit is 6 MB per file.`);
        return [];
      }
      return [{ file, mimeType }];
    });
    if (!supported.length) return;
    if (attachments.reduce((total, file) => total + file.size, 0) + supported.reduce((total, item) => total + item.file.size, 0) > MAX_TOTAL_ATTACHMENT_BYTES) {
      setError("Attachments are limited to 12 MB total.");
      return;
    }
    setUploadingFiles(true);
    try {
      const next = await Promise.all(supported.map(async ({ file, mimeType }) => ({ id: uid(), name: file.name, mimeType, size: file.size, dataUrl: await readDataUrl(file) })));
      setAttachments(previous => [...previous, ...next]);
      setError("");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not read the attachment");
    } finally {
      setUploadingFiles(false);
    }
  }, [attachments]);

  const removeAttachment = (id: string): void => setAttachments(previous => previous.filter(file => file.id !== id));

  const sendInstruction = useCallback(async (instruction = prompt): Promise<void> => {
    const currentAttachments = attachments;
    const trimmed = instruction.trim() || (currentAttachments.length ? "Review the attached files and summarize the relevant findings." : "");
    if ((!trimmed && !currentAttachments.length) || busy || !coreOnline || !selectedModelId || !wordReady) return;
    setPrompt("");
    setAttachments([]);
    setComposerToolsOpen(false);
    setError("");
    const current = await refreshSnapshot();
    const userMessage: UiMessage = { id: uid(), role: "user", content: trimmed, ...(currentAttachments.length ? { attachments: currentAttachments } : {}) };
    const assistantId = uid();
    setMessages(previous => [...previous, userMessage, { id: assistantId, role: "assistant", content: "" }]);
    const controller = new AbortController();
    setAbortController(controller);
    setBusy(true);
    try {
      await streamAgent({
        providerId: selectedProviderId,
        modelId: selectedModelId,
        effort: modelEffort,
        instruction: trimmed,
        mode,
        document: current,
        ...(messages.length ? { conversation: messages.slice(-8).map(message => ({ role: message.role, content: message.content })) } : {}),
        ...(currentAttachments.length ? { attachments: currentAttachments } : {}),
        skills: activeSkills,
        tools: { ...(webSearchEnabled ? { webSearch: true } : {}), ...(consoleEnabled ? { console: true } : {}) },
      }, (event: AgentEvent) => {
        if (event.type === "token") {
          setMessages(previous => previous.map(message => message.id === assistantId ? { ...message, content: message.content + event.delta } : message));
        }
        if (event.type === "tool") {
          setMessages(previous => previous.map(message => message.id === assistantId ? { ...message, toolActivity: event.state === "started" ? `Reading ${event.name.replace(/_/g, " ")}…` : undefined } : message));
        }
        if (event.type === "proposal") {
          if (!autoAppliedChanges.current.has(event.change.id)) {
            autoAppliedChanges.current.add(event.change.id);
            applyQueue.current = applyQueue.current.catch(() => undefined).then(() => apply(event.change, assistantId));
          }
        }
        if (event.type === "action") {
          setMessages(previous => previous.map(message => message.id === assistantId ? { ...message, actions: [...(message.actions ?? []).filter(action => action.id !== event.action.id), event.action] } : message));
        }
        if (event.type === "done") {
          setMessages(previous => previous.map(message => message.id === assistantId ? { ...message, content: event.answer, toolActivity: undefined } : message));
        }
        if (event.type === "error") setError(event.message);
      }, controller.signal);
    } catch (cause) {
      if (!controller.signal.aborted) setError(cause instanceof Error ? cause.message : "Generation failed");
    } finally {
      setBusy(false);
      setAbortController(null);
    }
  }, [attachments, busy, consoleEnabled, coreOnline, messages, mode, prompt, refreshSnapshot, selectedModelId, selectedProviderId, webSearchEnabled, wordReady]);

  async function apply(change: ProposedChange, assistantId: string): Promise<void> {
    setError("");
    try {
      const current = await refreshSnapshot();
      const currentParagraph = change.target.kind === "paragraph"
        ? current.paragraphs.find(item => item.id === change.target.id)
          ?? (change.target.paragraphIndex !== undefined ? current.paragraphs[change.target.paragraphIndex] : undefined)
          ?? current.paragraphs.find(item => item.text === (change.before ?? change.target.beforeText))
        : undefined;
      const currentVisual = change.target.kind === "visual"
        ? current.visualElements?.find(item => item.id === change.target.id)
          ?? current.visualElements?.find(item => item.kind === change.target.visualKind && item.index === change.target.visualIndex)
        : undefined;
      const expectedBefore = change.before ?? change.target.beforeText;
      const paragraphBefore = currentParagraph?.text;
      const currentBefore = change.target.kind === "selection"
        ? current.selection.text
        : change.target.kind === "document"
        ? comparableText(current.documentText) === comparableText(expectedBefore) ? expectedBefore : current.documentText
        : change.target.kind === "visual"
        ? (currentVisual ? visualElementTargetText(currentVisual) : "")
        : paragraphBefore && comparableText(paragraphBefore) === comparableText(expectedBefore)
        ? expectedBefore
        : paragraphBefore ?? "";
      const approved = await approveChange(change.id, currentBefore);
      const result = await adapter.applyChange(approved);
      await completeChange(change.id, result.success, result.message);
      if (!result.success) {
        const failure = result.message ?? "Word rejected the change";
        setError(failure);
        setMessages(previous => previous.map(message => message.id === assistantId
          ? { ...message, content: `${message.content}\n\nI couldn’t apply this edit in Word: ${failure}`.trim() }
          : message));
      }
      else {
        await refreshSnapshot();
        const edit: AppliedEdit = { id: change.id, description: change.description, before: change.before ?? currentBefore, ...(typeof change.after === "string" ? { after: change.after } : {}) };
        setMessages(previous => previous.map(message => message.id === assistantId ? { ...message, edits: [...(message.edits ?? []).filter(item => item.id !== edit.id), edit] } : message));
      }
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not apply change");
    }
  }

  const updateActionInMessages = (action: AgentAction): void => {
    setMessages(previous => previous.map(message => message.actions?.some(item => item.id === action.id)
      ? { ...message, actions: message.actions.map(item => item.id === action.id ? action : item) }
      : message));
  };

  async function runConsoleAction(action: AgentAction): Promise<void> {
    setActingAction(action.id);
    try {
      updateActionInMessages(await approveConsoleAction(action.id));
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not run the console action");
    } finally {
      setActingAction(null);
    }
  }

  async function blockConsoleAction(action: AgentAction): Promise<void> {
    try {
      updateActionInMessages(await rejectConsoleAction(action.id));
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not reject the console action");
    }
  }

  const chooseProvider = async (id: string): Promise<void> => {
    setSelectedProviderId(id);
    setSelectedModelId("");
    await saveSettings({ selectedProviderId: id });
    await loadModels(id);
  };

  const loginChatGPT = async (): Promise<void> => {
    setChatGptLoginBusy(true);
    setError("");
    try {
      const flow = await startChatGPTLogin();
      setChatGptFlowId(flow.flowId);
      const browser = window.open(flow.authorizeUrl, "_blank", "noopener,noreferrer");
      if (!browser) setError("Your browser blocked the sign-in window. Allow pop-ups for the OpenWordCode Core address and try again.");
    } catch (cause) {
      setChatGptLoginBusy(false);
      setError(cause instanceof Error ? cause.message : "Could not start ChatGPT sign-in");
    }
  };

  const beginCodexCliLogin = async (): Promise<void> => {
    setChatGptLoginBusy(true);
    setError("");
    try {
      await startCodexCliLogin();
      setCodexCliLoginStarted(true);
      await loadCore();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not start Codex CLI sign-in");
    } finally {
      setChatGptLoginBusy(false);
    }
  };

  const connectCodexCli = async (): Promise<void> => {
    setChatGptLoginBusy(true);
    setError("");
    try {
      await useCodexCliSession();
      setCodexCliLoginStarted(false);
      await saveSettings({ selectedProviderId: "openwordcode-bridge", selectedModelId: "gpt-5.6-luna" });
      await loadCore();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not use the Codex CLI session");
    } finally {
      setChatGptLoginBusy(false);
    }
  };

  const logoutChatGPT = async (): Promise<void> => {
    try {
      setCodexCliLoginStarted(false);
      await disconnectChatGPT();
      await loadCore();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not disconnect ChatGPT");
    }
  };

  const openOAuthBrowser = (): void => {
    const url = oauthLoginInfo?.authorizeUrl ?? oauthLoginInfo?.verificationUrl;
    if (!url) return;
    const browser = window.open(url, "_blank", "noopener,noreferrer");
    if (!browser) setError("Your browser blocked the sign-in window. Allow pop-ups for the OpenWordCode Core address and try again.");
  };

  const loginOAuth = async (): Promise<void> => {
    if (!selectedProvider || selectedProvider.kind === "openwordcode-bridge") return;
    setOAuthLoginBusy(true);
    setError("");
    try {
      const flow = await startOAuthLogin(selectedProvider.id);
      setOAuthLoginInfo(flow);
      setOAuthFlowId(flow.flowId);
      setOAuthManualCode("");
      const url = flow.authorizeUrl ?? flow.verificationUrl;
      if (url) {
        const browser = window.open(url, "_blank", "noopener,noreferrer");
        if (!browser) setError("Your browser blocked the sign-in window. Allow pop-ups for the OpenWordCode Core address and try again.");
      }
    } catch (cause) {
      setOAuthLoginBusy(false);
      setOAuthLoginInfo(null);
      setOAuthManualCode("");
      setError(cause instanceof Error ? cause.message : "Could not start OAuth sign-in");
    }
  };

  const completeXaiCode = async (): Promise<void> => {
    if (!oauthFlowId || !oauthManualCode.trim()) return;
    setOAuthLoginBusy(true);
    setError("");
    try {
      const result = await completeOAuthLogin(oauthFlowId, oauthManualCode.trim());
      setOAuthLoginInfo(result);
      if (result.status === "connected") {
        setOAuthFlowId(null);
        setOAuthLoginBusy(false);
        setOAuthManualCode("");
        await loadCore();
      }
    } catch (cause) {
      setOAuthLoginBusy(false);
      setError(cause instanceof Error ? cause.message : "Could not complete xAI sign-in");
    }
  };

  const beginLocalCliLogin = async (): Promise<void> => {
    if (!selectedProvider || !localCliForProvider(selectedProvider.id)) return;
    setOAuthLoginBusy(true);
    setError("");
    try {
      await startLocalCliLogin(selectedProvider.id);
      setLocalCliLoginProviderId(selectedProvider.id);
      await loadCore();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not start the local CLI sign-in");
      setOAuthLoginBusy(false);
    }
  };

  const connectLocalCli = async (): Promise<void> => {
    if (!selectedProvider || !localCliForProvider(selectedProvider.id)) return;
    setOAuthLoginBusy(true);
    setError("");
    try {
      await useLocalCliSession(selectedProvider.id);
      setLocalCliLoginProviderId(null);
      await loadCore();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not use the local CLI session");
    } finally {
      setOAuthLoginBusy(false);
    }
  };

  const cancelOAuth = async (): Promise<void> => {
    if (!oauthFlowId) return;
    try {
      await cancelOAuthLogin(oauthFlowId);
      setOAuthFlowId(null);
      setOAuthLoginInfo(null);
      setOAuthLoginBusy(false);
      setOAuthManualCode("");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not cancel OAuth sign-in");
    }
  };

  const logoutOAuth = async (): Promise<void> => {
    if (!selectedProvider || selectedProvider.kind === "openwordcode-bridge") return;
    try {
      setLocalCliLoginProviderId(null);
      setOAuthFlowId(null);
      setOAuthLoginBusy(false);
      await disconnectOAuth(selectedProvider.id);
      setOAuthLoginInfo(null);
      setOAuthManualCode("");
      await loadCore();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not disconnect OAuth account");
    }
  };

  const openChat = (): void => {
    setTab("chat");
    setError("");
  };

  const newTask = (): void => {
    if (messages.length) setTaskSessions(previous => upsertTask(previous, activeTaskId, messages));
    abortController?.abort();
    setActiveTaskId(uid());
    setMessages([]);
    setPrompt("");
    setAttachments([]);
    setComposerToolsOpen(false);
    setRecentTasksOpen(false);
    setError("");
    setTab("chat");
  };

  const openTask = (taskId: string): void => {
    const task = taskSessions.find(session => session.id === taskId);
    if (!task) return;
    if (messages.length && activeTaskId !== taskId) setTaskSessions(previous => upsertTask(previous, activeTaskId, messages));
    abortController?.abort();
    setActiveTaskId(task.id);
    setMessages(task.messages);
    setPrompt("");
    setAttachments([]);
    setComposerToolsOpen(false);
    setRecentTasksOpen(false);
    setError("");
    setTab("chat");
  };

  const deleteTask = (taskId: string): void => {
    setTaskSessions(previous => previous.filter(session => session.id !== taskId));
    if (taskId !== activeTaskId) return;
    abortController?.abort();
    setActiveTaskId(uid());
    setMessages([]);
    setPrompt("");
    setAttachments([]);
    setComposerToolsOpen(false);
    setRecentTasksOpen(false);
    setError("");
    setTab("chat");
  };

  const clearTaskHistory = (): void => {
    abortController?.abort();
    setTaskSessions([]);
    setActiveTaskId(uid());
    setMessages([]);
    setPrompt("");
    setAttachments([]);
    setComposerToolsOpen(false);
    setRecentTasksOpen(false);
    setError("");
    setTab("chat");
  };

  const recentTasks: RecentTask[] = taskSessions.map(({ id, title, updatedAt }) => ({ id, title, updatedAt }));

  const changeApprovalMode = (next: AgentMode): void => {
    setMode(next);
    setApprovalNoticeDismissed(false);
    void saveSettings({ mode: next });
  };

  const toggleSkill = (id: string): void => {
    setSkills(current => {
      const next = current.map(skill => skill.id === id ? { ...skill, enabled: skill.enabled === false } : skill);
      saveStoredSkills(next);
      return next;
    });
  };

  const uploadSkillFile = async (file: File): Promise<void> => {
    try {
      const text = await file.text();
      const parsed = parseSkillFile(text, file.name);
      const newSkill: SkillSummary = {
        id: `skill_${uid()}`,
        ...parsed,
      };
      setSkills(current => {
        const next = [...current, newSkill];
        saveStoredSkills(next);
        return next;
      });
      notifySuccess(`Added skill: ${newSkill.name}`);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not read skill file");
    }
  };

  const createSkill = (draft: Omit<SkillSummary, "id">): void => {
    const newSkill: SkillSummary = {
      id: `skill_${uid()}`,
      ...draft,
    };
    setSkills(current => {
      const next = [...current, newSkill];
      saveStoredSkills(next);
      return next;
    });
    notifySuccess(`Saved skill: ${newSkill.name}`);
  };

  const deleteSkill = (id: string): void => {
    setSkills(current => {
      const next = current.filter(skill => skill.id !== id);
      saveStoredSkills(next);
      return next;
    });
  };

  const composerTools: SearchItem[] = [
    { label: "Add files or photos", shortcut: "Ctrl+U", icon: <Paperclip size={17} />, action: () => fileInputRef.current?.click() },
    { label: "Search the web", detail: webSearchEnabled ? "Enabled for this chat" : "Use live sources for current questions", icon: <Globe size={17} />, checked: webSearchEnabled, action: () => setWebSearchEnabled(current => !current) },
    { label: "Use Windows console", detail: consoleEnabled ? "Safe workspace commands enabled" : "Let the AI inspect the workspace with safe commands", icon: <TerminalWindow size={17} />, checked: consoleEnabled, action: () => setConsoleEnabled(current => !current) },
    {
      label: "AI Skills",
      detail: activeSkills.length ? `${activeSkills.length} active prompt ${activeSkills.length === 1 ? "recipe" : "recipes"}` : "Prompt recipes",
      icon: <Sparkle size={17} weight="fill" />,
      action: () => setSkillsDrawerOpen(true),
    },
  ];

  return (
    <main className="shell">
      <header className="topbar">
        <button className="brand" onClick={openChat} aria-label="OpenWordCode chat">
          <span className="brand-mark"><img className="brand-logo" src="/openwordcode-logo.png" alt="" aria-hidden="true" /></span>
          <span className="brand-copy"><strong>OpenWordCode</strong></span>
        </button>
        <div className="header-actions">
          <span className={`core-status ${coreOnline ? "online" : "offline"}`} role="status" aria-label={coreOnline ? "Core connected" : "Core offline"} title={coreOnline ? `Core ${coreVersion}` : "Start OpenWordCode Core"}>
            <span className="status-dot" />
          </span>
          <button className={`icon-button ${recentTasksOpen ? "active" : ""}`} onClick={() => setRecentTasksOpen(true)} aria-label="Chat history" title="Chat history">
            <ClockCounterClockwise size={16} />
            {recentTasks.length ? <span className="count-badge">{recentTasks.length}</span> : null}
          </button>
          <button className="new-task-button" onClick={newTask} aria-label="New task" title="New task">
            <Plus size={13} weight="bold" /><span>New</span>
          </button>
          <button className={`icon-button ${tab === "settings" ? "active" : ""}`} onClick={() => setTab("settings")} aria-label="Settings" title="Settings">
            <GearSix size={16} />
          </button>
        </div>
      </header>

      {recentTasksOpen ? <RecentTasksDrawer tasks={recentTasks} activeTaskId={activeTaskId} onClose={() => setRecentTasksOpen(false)} onSelect={openTask} onNewTask={newTask} onDelete={deleteTask} onClearAll={clearTaskHistory} /> : null}
      {skillsDrawerOpen ? <SkillsDrawer skills={skills} onToggleSkill={toggleSkill} onUploadSkill={uploadSkillFile} onCreateSkill={createSkill} onDeleteSkill={deleteSkill} onClose={() => setSkillsDrawerOpen(false)} /> : null}

      {tab === "chat" ? (
        <section className="chat-panel">
          <div className="chat-scroll">
            {messages.length === 0 ? <div className="empty-chat" aria-hidden="true" /> : (
              <div className="messages">
                {messages.map(message => (
                  <article className={`message ${message.role}`} key={message.id}>
                    {message.role === "user" ? <div className="user-message" dir="auto"><span>{message.content}</span>{message.attachments ? <AttachmentList attachments={message.attachments} label="Sent attachments" /> : null}</div> : <div className="assistant-message"><div className="assistant-label"><Sparkle size={13} weight="fill" /> OpenWordCode</div><div className="assistant-content" dir="auto">{busy && message.id === messages[messages.length - 1]?.id ? <ThinkingTrace active activity={message.toolActivity} /> : null}{message.content || null}{message.edits ? <AppliedEditList edits={message.edits} /> : null}{message.actions?.map(action => <ConsoleActionCard key={action.id} action={action} acting={actingAction === action.id} onApprove={runConsoleAction} onReject={blockConsoleAction} />)}</div></div>}
                  </article>
                ))}
              </div>
            )}

          </div>

          {mode === "skip" && !approvalNoticeDismissed ? <div className="approval-notice"><div><strong>Skip all approvals is on.</strong><p>OpenWordCode will not pause, even for unsafe document actions. You can change this from the approval control below.</p></div><button type="button" onClick={() => setApprovalNoticeDismissed(true)} aria-label="Dismiss approval notice" title="Dismiss"><X size={18} /></button></div> : null}

          <div className="composer-wrap">
            {attachments.length ? <AttachmentList attachments={attachments} onRemove={removeAttachment} /> : null}
            <textarea dir="auto" value={prompt} onChange={event => setPrompt(event.target.value)} onKeyDown={event => { if (event.key === "Enter" && !event.shiftKey) { event.preventDefault(); void sendInstruction(); } }} placeholder="Ask OpenWordCode…" disabled={!coreOnline || busy || !wordReady} rows={2} />
            <div className="composer-footer">
              <div className="composer-tools">
                <button className={`composer-add ${composerToolsOpen ? "active" : ""}`} onClick={() => setComposerToolsOpen(current => !current)} aria-label="Open chat tools" aria-expanded={composerToolsOpen} title="Chat tools"><Plus size={15} weight="bold" /></button>
                {composerToolsOpen ? <SearchList items={composerTools} onClose={() => setComposerToolsOpen(false)} /> : null}
              </div>
              <input ref={fileInputRef} className="file-input" type="file" accept="image/png,image/jpeg,image/webp,image/gif,application/pdf,.pdf" multiple onChange={event => void onFilesSelected(event)} />
              <ModePicker mode={mode} onChange={changeApprovalMode} disabled={!coreOnline || busy || !wordReady} />
              <ModelPicker models={models} selectedModelId={selectedModelId} label={modelLabel} loading={loadingModels} effort={modelEffort} onEffortChange={setModelEffort} onChange={modelId => { setSelectedModelId(modelId); void saveSettings({ selectedModelId: modelId }); }} disabled={!coreOnline || loadingModels || busy || !wordReady} />
              {busy ? <button className="send-button stop" onClick={() => abortController?.abort()} aria-label="Stop response" title="Stop response"><X size={14} weight="bold" /></button> : <button className="send-button" onClick={() => void sendInstruction()} disabled={(!prompt.trim() && !attachments.length) || !selectedModelId || !coreOnline || uploadingFiles || !wordReady} aria-label="Send message" title="Send message"><ArrowUp size={14} weight="bold" /></button>}
            </div>
          </div>
          <p className="composer-disclaimer">OpenWordCode can make mistakes. Check important information.</p>
        </section>
      ) : null}

      {tab === "settings" ? (
        <section className="secondary-view settings-view">
          <div className="settings-nav">
            <button className="back-button" onClick={openChat}><ArrowLeft size={15} /> Chat</button>
            <span className="settings-title">Settings</span>
            <button className="settings-refresh" onClick={() => void loadCore()} aria-label="Refresh" title="Refresh"><ArrowClockwise size={15} /></button>
          </div>

          <div className="settings-group">
            <div className="settings-label"><span>Accounts</span><small>{connectedProviderCount} connected</small></div>
            <div className="settings-list">
              {visibleProviders.map(provider => {
                const active = provider.id === selectedProviderId;
                return <div className={`settings-row-wrap ${active ? "expanded" : ""}`} key={provider.id}>
                  <button className={`settings-row ${active ? "selected" : ""}`} onClick={() => void chooseProvider(provider.id)} aria-expanded={active}>
                    <span className={`row-dot ${provider.auth.status}`} />
                    <span className="row-name">{provider.displayName}</span>
                    <span className="row-state">{accountSubline(provider)}</span>
                    <CaretRight size={13} className={active ? "row-caret open" : "row-caret"} />
                  </button>
                  {active && selectedProvider ? <div className="settings-row-detail">
                    {selectedProvider.kind === "openwordcode-bridge" ? <ChatGPTAccountCard provider={selectedProvider} busy={chatGptLoginBusy} codexLoginStarted={codexCliLoginStarted} onLogin={() => void loginChatGPT()} onStartCodexCliLogin={() => void beginCodexCliLogin()} onUseCodexCli={() => void connectCodexCli()} onDisconnect={() => void logoutChatGPT()} /> : selectedProvider.auth.availableMethods.includes("oauth") ? <OAuthAccountCard provider={selectedProvider} flow={oauthLoginInfo} busy={oauthLoginBusy} localCliLoginStarted={localCliLoginProviderId === selectedProvider.id} manualCode={oauthManualCode} onManualCodeChange={setOAuthManualCode} onCompleteManualCode={() => void completeXaiCode()} onLogin={() => void loginOAuth()} onOpenBrowser={openOAuthBrowser} onCancel={() => void cancelOAuth()} onStartLocalCliLogin={() => void beginLocalCliLogin()} onUseLocalCli={() => void connectLocalCli()} onDisconnect={() => void logoutOAuth()} /> : <div className="account-detail"><p className="account-detail-text">{selectedProvider.auth.detail}</p></div>}
                    {selectedProviderId === "openwordcode-bridge" ? <div className="bridge-meta"><span>{bridge?.available ? `${bridge.models} models · ${bridge.endpoint}` : bridge?.detail ?? "Detecting…"}</span><button className="text-button" onClick={() => void getBridgeStatus().then(setBridge)}>Refresh</button></div> : null}
                  </div> : null}
                </div>;
              })}
            </div>
          </div>

          <div className="settings-group">
            <div className="settings-label"><span>Appearance</span></div>
            <div className="segmented">{(["system", "light", "dark"] as const).map(value => <button className={theme === value ? "active" : ""} key={value} onClick={() => { setTheme(value); void saveSettings({ theme: value }); }}>{value[0]!.toUpperCase() + value.slice(1)}</button>)}</div>
          </div>

          <div className="about-row"><span>OpenWordCode</span><span>Core {coreVersion}</span></div>
        </section>
      ) : null}
    </main>
  );
}
