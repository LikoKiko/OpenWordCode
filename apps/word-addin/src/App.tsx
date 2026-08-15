import { useCallback, useEffect, useMemo, useRef, useState, type ChangeEvent, type JSX } from "react";
import {
  ArrowClockwise,
  ArrowLeft,
  ArrowUp,
  Check,
  ClockCounterClockwise,
  FileText,
  GearSix,
  Globe,
  Info,
  Key,
  Paperclip,
  Plus,
  Sparkle,
  TerminalWindow,
  X,
} from "@phosphor-icons/react";
import type { AgentAction, AgentEvent, AgentMode, ChatAttachment, DocumentSnapshot, ModelInfo, ProposedChange, ProviderSummary } from "../../../packages/shared/src/index.js";
import { comparableText, visualElementTargetText } from "../../../packages/shared/src/index.js";
import { createWordAdapter, isOfficeHost, waitForOfficeReady, type WordApplicationAdapter } from "../../../packages/app-word/src/index.js";
import { AppliedEditList, AttachmentList, ConsoleActionCard, ModePicker, ModelPicker, RecentTasksDrawer, SearchList, ThinkingTrace, type AppliedEdit, type AttachmentPreview, type ModelEffort, type RecentTask, type SearchItem } from "./components";
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
  saveApiKey,
  saveSettings,
  startOAuthLogin,
  startChatGPTLogin,
  streamAgent,
  cancelOAuthLogin,
  disconnectOAuth,
} from "./api";
import type { OAuthLoginStart } from "./api";

type Tab = "chat" | "settings";
type UiMessage = { id: string; role: "user" | "assistant"; content: string; toolActivity?: string; attachments?: AttachmentPreview[]; actions?: AgentAction[]; edits?: AppliedEdit[] };
type TaskSession = { id: string; title: string; createdAt: string; updatedAt: string; messages: UiMessage[] };

const MAX_ATTACHMENTS = 4;
const MAX_ATTACHMENT_BYTES = 6_000_000;
const MAX_TOTAL_ATTACHMENT_BYTES = 12_000_000;
const TASK_HISTORY_STORAGE_KEY = "openwordcode.task-history.v1";
const MAX_TASK_HISTORY = 20;
const SUPPORTED_IMAGE_TYPES = new Set(["image/gif", "image/jpeg", "image/png", "image/webp"]);
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

function authLabel(provider: ProviderSummary): string {
  if (provider.auth.status === "connected") return provider.auth.method === "environment" ? "Environment" : provider.auth.method === "api-key" ? "API key" : provider.auth.method === "oauth" ? provider.kind === "openwordcode-bridge" ? "OpenWordCode account" : "OAuth account" : "Connected";
  if (provider.auth.status === "not-configured" && provider.local) return "Local";
  if (provider.auth.status === "login-required" && provider.auth.method === "api-key") return "Add API key";
  if (provider.auth.status === "login-required" && provider.auth.method === "oauth") return "Sign in";
  if (provider.auth.status === "unsupported" && provider.auth.method === "oauth") return "Setup required";
  return provider.auth.status.replace(/-/g, " ");
}

function ChatGPTAccountCard({ provider, busy, compact, onLogin, onDisconnect }: { provider: ProviderSummary; busy: boolean; compact?: boolean; onLogin: () => void; onDisconnect: () => void }): JSX.Element {
  const connected = provider.auth.status === "connected";
  const unsupported = provider.auth.status === "unsupported";
  return <article className={`account-card${compact ? " compact" : ""}`}>
    <div className="account-card-heading"><span className="account-card-icon"><Sparkle size={16} weight="fill" /></span><div><strong>OpenWordCode Bridge</strong><small>{connected ? "Account connected" : "Account sign-in"}</small></div><span className={`auth-badge ${provider.auth.status}`}>{authLabel(provider)}</span></div>
    <h2>{connected ? "Your OpenWordCode account is ready" : "Connect your OpenWordCode account"}</h2>
    <p>{connected ? provider.auth.detail : "Sign in in your browser. OpenWordCode never asks for your account password inside Word."}</p>
    {unsupported ? <p className="account-card-warning">This build needs an OpenWordCode-owned OAuth client registration before account sign-in can be enabled.</p> : null}
    <div className="account-card-actions">{connected ? <button className="text-button" type="button" onClick={onDisconnect}>Disconnect</button> : <button className="apply-button" type="button" onClick={onLogin} disabled={busy || unsupported}>{busy ? "Opening browser…" : "Sign in"}</button>}</div>
  </article>;
}

function OAuthAccountCard({ provider, flow, busy, compact, onLogin, onOpenBrowser, onCancel, onDisconnect }: { provider: ProviderSummary; flow: OAuthLoginStart | null; busy: boolean; compact?: boolean; onLogin: () => void; onOpenBrowser: () => void; onCancel: () => void; onDisconnect: () => void }): JSX.Element {
  const flowForProvider = flow?.providerId === provider.id ? flow : null;
  const accountName = provider.id === "anthropic" ? "Claude" : provider.displayName;
  const connected = provider.auth.method === "oauth" && provider.auth.status === "connected";
  const unsupported = provider.auth.status === "unsupported" && provider.auth.method === "oauth";
  return <article className={`account-card oauth-card${compact ? " compact" : ""}`}>
    <div className="account-card-heading"><span className="account-card-icon"><Globe size={16} /></span><div><strong>{accountName}</strong><small>{connected ? "Account connected" : "Account sign-in"}</small></div><span className={`auth-badge ${connected ? "connected" : unsupported ? "unsupported" : "login-required"}`}>{connected ? "Ready" : unsupported ? "Unavailable" : "Sign in"}</span></div>
    <h2>{connected ? `${accountName} account is ready` : `Connect your ${accountName} account`}</h2>
    <p>{connected ? provider.auth.detail : "Sign in in your browser. OpenWordCode never asks for your account password inside Word; the token stays in the local Core credential store."}</p>
    {unsupported ? <p className="account-card-warning">This provider needs a provider-specific transport that is not enabled in this build.</p> : null}
    {flowForProvider?.userCode ? <div className="oauth-code"><small>Verification code</small><code>{flowForProvider.userCode}</code></div> : null}
    {flowForProvider?.detail ? <p className="oauth-flow-detail">{flowForProvider.detail}</p> : null}
    <div className="account-card-actions">
      {connected ? <button className="text-button" type="button" onClick={onDisconnect}>Disconnect</button> : flowForProvider ? <><button className="text-button" type="button" onClick={onCancel}>Cancel</button>{flowForProvider.authorizeUrl || flowForProvider.verificationUrl ? <button className="apply-button" type="button" onClick={onOpenBrowser}>Open sign-in</button> : null}</> : <button className="apply-button" type="button" onClick={onLogin} disabled={busy || unsupported}>{busy ? "Preparing sign-in…" : "Sign in"}</button>}
    </div>
  </article>;
}

function ApiKeyCard({ provider, value, busy, compact, onChange, onConnect }: { provider: ProviderSummary; value: string; busy: boolean; compact?: boolean; onChange: (value: string) => void; onConnect: () => void }): JSX.Element {
  const connected = provider.auth.status === "connected";
  return <article className={`account-card api-key-card${compact ? " compact" : ""}`}>
    <div className="account-card-heading"><span className="account-card-icon"><Key size={16} /></span><div><strong>{provider.displayName}</strong><small>{connected ? "API key connected" : "API key setup"}</small></div><span className={`auth-badge ${provider.auth.status}`}>{connected ? "Ready" : "Required"}</span></div>
    <h2>{connected ? "API access is ready" : `Connect ${provider.displayName}`}</h2>
    <p>{connected ? "Replace the key below whenever you need to use a different account." : "Paste your API key to connect. It is sent only to OpenWordCode Core and stored in its local encrypted credential store."}</p>
    <div className="key-form"><input type="password" value={value} onChange={event => onChange(event.target.value)} placeholder={connected ? "Paste a replacement API key" : "Paste API key"} autoComplete="off" /><button className="apply-button" type="button" onClick={onConnect} disabled={!value.trim() || busy}>{busy ? "Connecting…" : connected ? "Replace key" : "Connect"}</button></div>
  </article>;
}

export default function App(): JSX.Element {
  const [adapter, setAdapter] = useState<WordApplicationAdapter>(() => createWordAdapter());
  const [wordReady, setWordReady] = useState(() => !isOfficeHost());
  const [snapshot, setSnapshot] = useState<DocumentSnapshot | null>(null);
  const [coreOnline, setCoreOnline] = useState(false);
  const [coreVersion, setCoreVersion] = useState("—");
  const [providers, setProviders] = useState<ProviderSummary[]>([]);
  const [models, setModels] = useState<ModelInfo[]>([]);
  const [selectedProviderId, setSelectedProviderId] = useState("openai");
  const [selectedModelId, setSelectedModelId] = useState("");
  const [modelEffort, setModelEffort] = useState<ModelEffort>("high");
  const [mode, setMode] = useState<AgentMode>("manual");
  const [theme, setTheme] = useState<"light" | "dark" | "system">("system");
  const [tab, setTab] = useState<Tab>("chat");
  const [prompt, setPrompt] = useState("");
  const [taskSessions, setTaskSessions] = useState<TaskSession[]>(loadTaskSessions);
  const [activeTaskId, setActiveTaskId] = useState(() => taskSessions[0]?.id ?? uid());
  const [recentTasksOpen, setRecentTasksOpen] = useState(false);
  const [messages, setMessages] = useState<UiMessage[]>(() => taskSessions[0]?.messages ?? []);
  const [busy, setBusy] = useState(false);
  const [loadingModels, setLoadingModels] = useState(false);
  const [error, setError] = useState("");
  const [keyInput, setKeyInput] = useState("");
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
  const [oauthFlowId, setOAuthFlowId] = useState<string | null>(null);
  const [oauthLoginInfo, setOAuthLoginInfo] = useState<OAuthLoginStart | null>(null);
  const [oauthLoginBusy, setOAuthLoginBusy] = useState(false);
  const [approvalNoticeDismissed, setApprovalNoticeDismissed] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const autoAppliedChanges = useRef(new Set<string>());
  const applyQueue = useRef(Promise.resolve());

  const selectedProvider = useMemo(() => providers.find(provider => provider.id === selectedProviderId), [providers, selectedProviderId]);
  const chatGptProvider = useMemo(() => providers.find(provider => provider.id === "openwordcode-bridge"), [providers]);
  const genericOAuthVisible = Boolean(selectedProvider && selectedProvider.kind !== "openwordcode-bridge" && selectedProvider.auth.availableMethods.includes("oauth") && (selectedProvider.auth.method === "oauth" || selectedProvider.auth.method === "environment") && selectedProvider.auth.status !== "connected");
  const selectedModel = useMemo(() => models.find(model => model.id === selectedModelId), [models, selectedModelId]);
  const modelLabel = selectedModel?.name || selectedModelId || (loadingModels ? "Loading model…" : "Choose model");
  const selectionLabel = snapshot?.selection.isTable
    ? `Table selected${snapshot.selection.tableCount && snapshot.selection.tableCount > 1 ? ` · ${snapshot.selection.tableCount} tables` : ""}`
    : snapshot?.selection.isEmpty
    ? "No text selected"
    : `${snapshot?.selection.text.length ?? 0} characters selected`;

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
          await loadCore();
          return;
        }
        if (result.status === "error" || result.status === "cancelled") {
          setOAuthFlowId(null);
          setOAuthLoginBusy(false);
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
        instruction: trimmed,
        mode,
        document: current,
        ...(messages.length ? { conversation: messages.slice(-8).map(message => ({ role: message.role, content: message.content })) } : {}),
        ...(currentAttachments.length ? { attachments: currentAttachments } : {}),
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

  const connectKey = async (): Promise<void> => {
    if (!keyInput.trim() || !selectedProvider) return;
    try {
      await saveApiKey(selectedProvider.id, keyInput.trim());
      setKeyInput("");
      setProviders(await getProviders());
      await loadModels(selectedProvider.id);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not save the API key");
    }
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

  const logoutChatGPT = async (): Promise<void> => {
    try {
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
      const url = flow.authorizeUrl ?? flow.verificationUrl;
      if (url) {
        const browser = window.open(url, "_blank", "noopener,noreferrer");
        if (!browser) setError("Your browser blocked the sign-in window. Allow pop-ups for the OpenWordCode Core address and try again.");
      }
    } catch (cause) {
      setOAuthLoginBusy(false);
      setOAuthLoginInfo(null);
      setError(cause instanceof Error ? cause.message : "Could not start OAuth sign-in");
    }
  };

  const cancelOAuth = async (): Promise<void> => {
    if (!oauthFlowId) return;
    try {
      await cancelOAuthLogin(oauthFlowId);
      setOAuthFlowId(null);
      setOAuthLoginInfo(null);
      setOAuthLoginBusy(false);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not cancel OAuth sign-in");
    }
  };

  const logoutOAuth = async (): Promise<void> => {
    if (!selectedProvider || selectedProvider.kind === "openwordcode-bridge") return;
    try {
      await disconnectOAuth(selectedProvider.id);
      setOAuthLoginInfo(null);
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

  const composerTools: SearchItem[] = [
    { label: "Add files or photos", shortcut: "Ctrl+U", icon: <Paperclip size={17} />, action: () => fileInputRef.current?.click() },
    { label: "Search the web", detail: webSearchEnabled ? "Enabled for this chat" : "Use live sources for current questions", icon: <Globe size={17} />, checked: webSearchEnabled, action: () => setWebSearchEnabled(current => !current) },
    { label: "Use Windows console", detail: consoleEnabled ? "Safe workspace commands enabled" : "Let the AI inspect the workspace with safe commands", icon: <TerminalWindow size={17} />, checked: consoleEnabled, action: () => setConsoleEnabled(current => !current) },
    {
      label: "Skills",
      icon: <Sparkle size={17} weight="fill" />,
      children: [
        { label: "Refresh document context", detail: selectionLabel, icon: <ArrowClockwise size={16} />, action: () => { void refreshSnapshot(); } },
        { label: "Review selected text", detail: "Check clarity, grammar, and risks", icon: <Sparkle size={16} weight="fill" />, action: () => { void sendInstruction("Review the selected text for clarity, grammar, inconsistencies, and important issues."); } },
        { label: "Rewrite selected text", detail: "Preserve meaning, improve the wording", icon: <FileText size={16} />, action: () => { void sendInstruction("Rewrite the selected text clearly while preserving its meaning."); } },
        { label: "Insert a table", detail: "Add a 3 × 4 table at the cursor", icon: <Plus size={16} />, action: () => { void sendInstruction("Add a 3 by 4 table at the current cursor."); } },
      ],
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
            <ClockCounterClockwise size={18} />
            {recentTasks.length ? <span className="count-badge">{recentTasks.length}</span> : null}
          </button>
          <button className="new-task-button" onClick={newTask} aria-label="New task" title="New task">
            <Plus size={15} weight="bold" /><span>New task</span>
          </button>
          <button className={`icon-button ${tab === "settings" ? "active" : ""}`} onClick={() => setTab("settings")} aria-label="Settings" title="Settings">
            <GearSix size={18} />
          </button>
        </div>
      </header>

      {error ? <div className="error-banner"><Info size={16} weight="fill" /><span>{error}</span><button onClick={() => setError("")} aria-label="Dismiss error"><X size={15} /></button></div> : null}

      {recentTasksOpen ? <RecentTasksDrawer tasks={recentTasks} activeTaskId={activeTaskId} onClose={() => setRecentTasksOpen(false)} onSelect={openTask} onNewTask={newTask} onDelete={deleteTask} onClearAll={clearTaskHistory} /> : null}

      {tab === "chat" ? (
        <section className="chat-panel">
          <div className="chat-scroll">
            {messages.length === 0 ? genericOAuthVisible ? <OAuthAccountCard provider={selectedProvider!} flow={oauthLoginInfo} busy={oauthLoginBusy} onLogin={() => void loginOAuth()} onOpenBrowser={openOAuthBrowser} onCancel={() => void cancelOAuth()} onDisconnect={() => void logoutOAuth()} /> : selectedProvider?.auth.status === "login-required" && selectedProvider.auth.availableMethods.includes("api-key") ? <ApiKeyCard provider={selectedProvider} value={keyInput} busy={loadingModels} onChange={setKeyInput} onConnect={() => void connectKey()} /> : chatGptProvider && selectedProvider?.kind === "openwordcode-bridge" ? <ChatGPTAccountCard provider={chatGptProvider} busy={chatGptLoginBusy} onLogin={() => void loginChatGPT()} onDisconnect={() => void logoutChatGPT()} /> : <div className="empty-chat" aria-hidden="true" /> : (
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
            <textarea dir="auto" value={prompt} onChange={event => setPrompt(event.target.value)} onKeyDown={event => { if (event.key === "Enter" && !event.shiftKey) { event.preventDefault(); void sendInstruction(); } }} placeholder="Write a message…" disabled={!coreOnline || busy || !wordReady} rows={3} />
            <div className="composer-footer">
              <div className="composer-tools">
                <button className={`composer-add ${composerToolsOpen ? "active" : ""}`} onClick={() => setComposerToolsOpen(current => !current)} aria-label="Open chat tools" aria-expanded={composerToolsOpen} title="Chat tools"><Plus size={19} /></button>
                {composerToolsOpen ? <SearchList items={composerTools} onClose={() => setComposerToolsOpen(false)} /> : null}
              </div>
              <input ref={fileInputRef} className="file-input" type="file" accept="image/png,image/jpeg,image/webp,image/gif,application/pdf,.pdf" multiple onChange={event => void onFilesSelected(event)} />
              <ModePicker mode={mode} onChange={changeApprovalMode} disabled={!coreOnline || busy || !wordReady} />
              <ModelPicker models={models} selectedModelId={selectedModelId} label={modelLabel} loading={loadingModels} effort={modelEffort} onEffortChange={setModelEffort} onChange={modelId => { setSelectedModelId(modelId); void saveSettings({ selectedModelId: modelId }); }} disabled={!coreOnline || loadingModels || busy || !wordReady} />
              {busy ? <button className="send-button stop" onClick={() => abortController?.abort()} aria-label="Stop response" title="Stop response"><X size={18} weight="bold" /></button> : <button className="send-button" onClick={() => void sendInstruction()} disabled={(!prompt.trim() && !attachments.length) || !selectedModelId || !coreOnline || uploadingFiles || !wordReady} aria-label="Send message" title="Send message"><ArrowUp size={18} weight="bold" /></button>}
            </div>
          </div>
          <p className="composer-disclaimer">OpenWordCode can make mistakes. Please double-check responses.</p>
        </section>
      ) : null}

      {tab === "settings" ? (
        <section className="secondary-view settings-view">
          <div className="view-heading"><button className="back-button" onClick={openChat}><ArrowLeft size={16} /> Chat</button><span className="settings-heading-label">Settings</span></div>
          <div className="view-title"><div><h1>Settings</h1><p>Control providers, authentication, and appearance.</p></div><button className="text-button" onClick={() => void loadCore()}>Refresh</button></div>

          <div className="settings-section">
            <h2>Providers</h2>
            <div className="provider-list">{providers.map(provider => <button className={`provider-row ${provider.id === selectedProviderId ? "selected" : ""}`} key={provider.id} onClick={() => void chooseProvider(provider.id)}><span className="provider-icon"><Sparkle size={15} weight="fill" /></span><span className="provider-copy"><strong>{provider.displayName}</strong><small>{provider.local ? "Local" : "Cloud"} · {authLabel(provider)}</small></span>{provider.id === selectedProviderId ? <Check size={16} weight="bold" /> : null}</button>)}</div>
          </div>

          <div className="settings-section settings-card">
            <div className="setting-row heading-row"><div><h2>Authentication</h2><p>{selectedProvider?.displayName ?? "Choose a provider"}</p></div>{selectedProvider ? <span className={`auth-badge ${selectedProvider.auth.status}`}>{authLabel(selectedProvider)}</span> : null}</div>
            {selectedProvider?.kind === "openwordcode-bridge" ? <ChatGPTAccountCard provider={selectedProvider} busy={chatGptLoginBusy} compact onLogin={() => void loginChatGPT()} onDisconnect={() => void logoutChatGPT()} /> : null}
            {selectedProvider && selectedProvider.kind !== "openwordcode-bridge" && selectedProvider.auth.availableMethods.includes("oauth") ? <OAuthAccountCard provider={selectedProvider} flow={oauthLoginInfo} busy={oauthLoginBusy} compact onLogin={() => void loginOAuth()} onOpenBrowser={openOAuthBrowser} onCancel={() => void cancelOAuth()} onDisconnect={() => void logoutOAuth()} /> : null}
            {selectedProvider?.auth.availableMethods.includes("api-key") ? <div className="key-form"><input type="password" value={keyInput} onChange={event => setKeyInput(event.target.value)} placeholder="Paste API key · stored only in Core" autoComplete="off" /><button className="apply-button" onClick={() => void connectKey()} disabled={!keyInput.trim() || loadingModels}>{loadingModels ? "Connecting…" : selectedProvider.auth.status === "connected" ? "Replace key" : "Connect"}</button></div> : null}
            <p className="setting-detail">{selectedProvider?.auth.detail ?? "Authentication status will appear here."}</p>
            {selectedProvider?.auth.availableMethods.includes("environment") ? <p className="setting-helper">Environment references are supported. Configure the provider variable in the Core process.</p> : null}
          </div>

          <div className="settings-section">
            <h2>Appearance</h2>
            <div className="segmented">{(["system", "light", "dark"] as const).map(value => <button className={theme === value ? "active" : ""} key={value} onClick={() => { setTheme(value); void saveSettings({ theme: value }); }}>{value[0]!.toUpperCase() + value.slice(1)}</button>)}</div>
          </div>

          {selectedProviderId === "openwordcode-bridge" ? <div className="bridge-card"><div><strong>OpenWordCode Bridge</strong><small>{bridge?.available ? `${bridge.models} models · ${bridge.endpoint}` : bridge?.detail ?? "Detecting local bridge…"}</small></div><button className="text-button" onClick={() => void getBridgeStatus().then(setBridge)}>Refresh</button></div> : null}
          <div className="about-row"><span>OpenWordCode</span><span>Core {coreVersion}</span></div>
        </section>
      ) : null}
    </main>
  );
}
