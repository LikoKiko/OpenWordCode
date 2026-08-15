import { useEffect, useRef, useState, type JSX, type ReactNode, type RefObject } from "react";
import {
  CaretDown,
  CaretRight,
  Check,
  CheckCircle,
  CircleNotch,
  Clock,
  FastForward,
  FileText,
  Gauge,
  Hand,
  ImageSquare,
  Info,
  Plus,
  Sparkle,
  TerminalWindow,
  Trash,
  X,
  Warning,
} from "@phosphor-icons/react";
import type { AgentAction, AgentMode, ChatAttachment, ModelInfo } from "../../../packages/shared/src/index.js";

export type AttachmentPreview = Omit<ChatAttachment, "dataUrl"> & { dataUrl?: string };
export type AppliedEdit = { id: string; description: string; before?: string; after?: string };

const LOADER_PATTERNS = {
  Drive: [0, 90, 180, 90, 180, 270, 180, 270, 360],
  Dots: [0, 90, 180, 90, 180, 270, 180, 270, 360],
  Orbit: [0, 110, 220, 770, 330, 440, 660, 550, null],
} as const;

function useElapsed(): string {
  const [ticks, setTicks] = useState(0);
  useEffect(() => {
    const timer = window.setInterval(() => setTicks(current => current + 1), 100);
    return () => window.clearInterval(timer);
  }, []);
  const seconds = ticks / 10;
  return seconds < 60 ? `${seconds.toFixed(1)}s` : `${Math.floor(seconds / 60)}m ${(seconds % 60).toFixed(1)}s`;
}

export function LoadingState({ label = "Working", variant = "Drive" }: { label?: string; variant?: keyof typeof LOADER_PATTERNS }): JSX.Element {
  const elapsed = useElapsed();
  const delays = LOADER_PATTERNS[variant] ?? LOADER_PATTERNS.Drive;
  const round = variant === "Dots";

  return (
    <div className="loading-state" role="status" aria-live="polite">
      <span className="loading-grid" aria-hidden="true">
        {delays.map((delay, index) => <i className={`loading-cell ${round ? "round" : ""} ${delay === null ? "idle" : `delay-${delay}`}`} key={index} />)}
      </span>
      <span className="loading-label">{label}</span>
      <span className="loading-time"><Clock size={12} />{elapsed}</span>
    </div>
  );
}

export function ThinkingTrace({ active, activity, steps = ["Reading document context", "Preparing a focused response"] }: { active: boolean; activity?: string; steps?: string[] }): JSX.Element {
  const [expanded, setExpanded] = useState(false);
  const label = active ? activity || "Working" : "Response ready";

  return (
    <section className={`thinking-trace ${active ? "active" : "settled"}`}>
      <button className="thinking-toggle" type="button" aria-expanded={expanded} onClick={() => setExpanded(current => !current)}>
        {active && !expanded ? <LoadingState label={label} /> : <><Sparkle size={14} weight="fill" /><span>{label}</span></>}
        <CaretDown className={expanded ? "expanded" : ""} size={13} />
      </button>
      {expanded ? (
        <div className="thinking-body">
          {active ? <LoadingState label={activity || "Working"} /> : <div className="thinking-settled"><CheckCircle size={14} weight="fill" />Completed</div>}
          <div className="thinking-steps">
            {steps.map((step, index) => <div className="thinking-step" key={step}><span className={`thinking-step-icon ${active && index === steps.length - 1 ? "current" : "done"}`}>{active && index === steps.length - 1 ? <CircleNotch className="spin" size={13} /> : <Check size={12} weight="bold" />}</span><span>{step}</span></div>)}
          </div>
        </div>
      ) : null}
    </section>
  );
}

function editPreview(value: string, max = 220): string {
  const normalized = value.replace(/\s+/g, " ").trim();
  return normalized.length > max ? `${normalized.slice(0, max).trimEnd()}…` : normalized;
}

export function AppliedEditList({ edits }: { edits: AppliedEdit[] }): JSX.Element {
  const [expanded, setExpanded] = useState(true);
  if (!edits.length) return <></>;
  const first = edits[0]!;

  return (
    <section className={`edit-history ${expanded ? "expanded" : "collapsed"}`}>
      <button className="edit-history-toggle" type="button" aria-expanded={expanded} onClick={() => setExpanded(current => !current)}>
        <span className="edit-history-heading"><CheckCircle size={14} weight="fill" /><span>Edited {edits.length}. {editPreview(first.description, 56)}</span></span>
        <CaretDown className={expanded ? "expanded" : ""} size={13} />
      </button>
      {expanded ? <div className="edit-history-list">
        {edits.map(edit => (
          <article className="edit-history-card" key={edit.id}>
            <div className="edit-history-card-heading"><span><FileText size={14} /><strong>{editPreview(edit.description, 96)}</strong></span><Check size={14} weight="bold" /></div>
            {edit.before ? <div className="edit-diff-row"><small>Before</small><div className="edit-before">{editPreview(edit.before)}</div></div> : null}
            {edit.after ? <div className="edit-diff-row"><small>After</small><div className="edit-after">{editPreview(edit.after)}</div></div> : <div className="edit-applied">Applied directly in Word</div>}
          </article>
        ))}
      </div> : null}
    </section>
  );
}

export type SearchItem = { label: string; detail?: string; shortcut?: string; icon?: ReactNode; action?: () => void; children?: SearchItem[]; checked?: boolean };

export function SearchList({ items, onClose }: { items: SearchItem[]; onClose: () => void }): JSX.Element {
  const [expanded, setExpanded] = useState<string | null>(null);

  const renderItem = (item: SearchItem, nested = false): JSX.Element => {
    const hasChildren = Boolean(item.children?.length);
    return (
      <div className={nested ? "tools-menu-nested-item" : "tools-menu-item-wrap"} key={item.label}>
        <button
          type="button"
          className={`tools-menu-item ${nested ? "nested" : ""} ${item.checked ? "checked" : ""}`}
          aria-pressed={item.checked === undefined ? undefined : item.checked}
          onClick={() => {
            if (hasChildren) {
              setExpanded(current => current === item.label ? null : item.label);
              return;
            }
            item.action?.();
            onClose();
          }}
        >
          <span className="tools-menu-icon">{item.icon || <Sparkle size={16} />}</span>
          <span className="tools-menu-copy"><strong>{item.label}</strong>{item.detail ? <small>{item.detail}</small> : null}</span>
          {item.shortcut ? <span className="tools-menu-shortcut">{item.shortcut}</span> : null}
          {item.checked ? <Check size={15} weight="bold" /> : null}
          {hasChildren ? <CaretRight className={expanded === item.label ? "expanded" : ""} size={16} /> : null}
        </button>
        {hasChildren && expanded === item.label ? <div className="tools-menu-submenu">{item.children!.map(child => renderItem(child, true))}</div> : null}
      </div>
    );
  };

  return <div className="tools-menu" role="dialog" aria-label="Chat tools">{items.map(item => renderItem(item))}</div>;
}

export function ConsoleActionCard({ action, acting, onApprove, onReject }: { action: AgentAction; acting?: boolean; onApprove: (action: AgentAction) => void; onReject: (action: AgentAction) => void }): JSX.Element {
  const pending = action.status === "pending";
  const completed = action.status === "completed";
  return (
    <article className={`console-action-card ${action.status}`}>
      <div className="console-action-heading"><span><TerminalWindow size={15} /><strong>Console action</strong></span><small>{completed ? "Completed" : pending ? "Approval required" : action.status}</small></div>
      <code dir="ltr">{action.command}</code>
      <p>{action.reason}</p>
      <small className="console-action-cwd" dir="ltr">Working directory: {action.workingDirectory}</small>
      {action.output ? <pre dir="ltr">{action.output}</pre> : null}
      {action.failureReason ? <p className="failure-note">{action.failureReason}</p> : null}
      {pending ? <div className="console-action-buttons"><button type="button" className="text-button" onClick={() => onReject(action)} disabled={acting}>Keep blocked</button><button type="button" className="apply-button" onClick={() => onApprove(action)} disabled={acting}>{acting ? "Running…" : "Run command"}</button></div> : null}
    </article>
  );
}

const APPROVAL_MODES: Array<{ value: AgentMode; label: string; detail: string }> = [
  { value: "manual", label: "Ask before every action", detail: "Review each change before it runs" },
  { value: "auto", label: "Auto-approve safe actions", detail: "Pause only when something looks unsafe" },
  { value: "skip", label: "Never pause for approval", detail: "Run every action without asking" },
];

function ApprovalModeIcon({ mode, size = 19 }: { mode: AgentMode; size?: number }): JSX.Element {
  if (mode === "manual") return <Hand size={size} />;
  if (mode === "auto") return <FastForward size={size} />;
  return <Warning size={size} />;
}

function useDismissibleMenu(): { open: boolean; setOpen: (value: boolean | ((current: boolean) => boolean)) => void; ref: RefObject<HTMLDivElement | null> } {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: PointerEvent) => {
      if (ref.current && !ref.current.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("pointerdown", onPointerDown);
    return () => document.removeEventListener("pointerdown", onPointerDown);
  }, [open]);
  return { open, setOpen, ref };
}

export function ModePicker({ mode, onChange, disabled }: { mode: AgentMode; onChange: (mode: AgentMode) => void; disabled?: boolean }): JSX.Element {
  const menu = useDismissibleMenu();
  const selected = APPROVAL_MODES.find(option => option.value === mode) ?? APPROVAL_MODES[0]!;
  return (
    <div className="picker mode-picker" ref={menu.ref}>
      <button className="picker-trigger approval-trigger" type="button" onClick={() => menu.setOpen(current => !current)} disabled={disabled} aria-haspopup="listbox" aria-expanded={menu.open} aria-label={`Approval limit: ${selected.label}`} title={selected.label}>
        <ApprovalModeIcon mode={mode} />
      </button>
      {menu.open ? (
        <div className="picker-menu approval-menu" role="listbox" aria-label="Approval limit">
          {APPROVAL_MODES.map(option => <button key={option.value} type="button" role="option" aria-selected={option.value === mode} className={`picker-option approval-option ${option.value === mode ? "selected" : ""}`} onClick={() => { onChange(option.value); menu.setOpen(false); }}><span className="approval-option-leading"><ApprovalModeIcon mode={option.value} /></span><span className="picker-option-copy"><strong>{option.label}</strong><small>{option.detail}</small></span>{option.value === mode ? <Check size={16} weight="bold" /> : null}</button>)}
        </div>
      ) : null}
    </div>
  );
}

export type ModelEffort = "low" | "medium" | "high";

const EFFORTS: Array<{ value: ModelEffort; label: string }> = [
  { value: "low", label: "Low" },
  { value: "medium", label: "Medium" },
  { value: "high", label: "High" },
];

const MODEL_DESCRIPTIONS = ["For toughest challenges", "For complex tasks", "Most efficient for everyday tasks", "Fastest for quick answers"];

export function ModelPicker({ models, selectedModelId, label, loading, disabled, effort, onEffortChange, onChange }: { models: ModelInfo[]; selectedModelId: string; label: string; loading?: boolean; disabled?: boolean; effort: ModelEffort; onEffortChange: (effort: ModelEffort) => void; onChange: (modelId: string) => void }): JSX.Element {
  const menu = useDismissibleMenu();
  const [showMoreModels, setShowMoreModels] = useState(true);
  const [effortOpen, setEffortOpen] = useState(false);
  const selected = models.find(model => model.id === selectedModelId);
  const featuredModels = models.slice(0, 4);
  const moreModels = models.slice(4);
  const effortLabel = EFFORTS.find(option => option.value === effort)?.label ?? "High";
  const selectModel = (modelId: string): void => {
    onChange(modelId);
    menu.setOpen(false);
  };
  const renderModel = (model: ModelInfo, index: number, featured: boolean): JSX.Element => {
    const name = model.name || model.id;
    const requiresCredits = model.id.toLocaleLowerCase().includes("fable");
    return (
      <button key={model.id} type="button" role="option" aria-selected={model.id === selectedModelId} className={`picker-option model-option ${featured ? "featured-model" : "more-model"} ${model.id === selectedModelId ? "selected" : ""}`} onClick={() => selectModel(model.id)}>
        <span className="picker-option-copy"><strong>{name}{requiresCredits ? <span className="credit-badge"><Info size={11} />Requires usage credits</span> : null}</strong>{featured ? <small>{MODEL_DESCRIPTIONS[index] ?? (model.capabilities.vision ? "Vision-ready model" : "Available through your provider")}</small> : null}</span>
        {model.id === selectedModelId ? <Check size={16} weight="bold" /> : null}
      </button>
    );
  };
  return (
    <div className="picker model-picker" ref={menu.ref}>
      <button className="picker-trigger model-trigger" type="button" onClick={() => menu.setOpen(current => !current)} disabled={disabled} aria-haspopup="listbox" aria-expanded={menu.open} aria-label={`Model: ${selected?.name || label}`}>
        <span className="picker-trigger-label model-trigger-label"><strong>{loading ? "Loading model…" : selected?.name || label}</strong><small>{effortLabel}</small></span>
        <CaretDown size={11} />
      </button>
      {menu.open ? (
        <div className="picker-menu model-menu" role="listbox" aria-label="Model">
          <div className="model-options">
            {featuredModels.length ? featuredModels.map((model, index) => renderModel(model, index, true)) : <div className="picker-empty">No models available</div>}
            {models.length > 4 ? <>
              <div className="picker-divider" />
              <button type="button" className="effort-row" onClick={() => setEffortOpen(current => !current)} aria-expanded={effortOpen}>
                <span><strong>Effort</strong><small>{effortLabel}</small></span><CaretRight className={effortOpen ? "expanded" : ""} size={17} />
              </button>
              {effortOpen ? <div className="effort-options" role="group" aria-label="Effort"><span className="effort-label"><Gauge size={14} /> Choose effort</span>{EFFORTS.map(option => <button type="button" key={option.value} className={option.value === effort ? "selected" : ""} onClick={() => { onEffortChange(option.value); setEffortOpen(false); }}>{option.label}{option.value === effort ? <Check size={14} weight="bold" /> : null}</button>)}</div> : null}
              <div className="picker-divider" />
              <button type="button" className="more-models-toggle" onClick={() => setShowMoreModels(current => !current)} aria-expanded={showMoreModels}><span>More models</span><CaretDown className={showMoreModels ? "expanded" : ""} size={15} /></button>
              {showMoreModels ? moreModels.map(model => renderModel(model, 0, false)) : null}
            </> : null}
          </div>
        </div>
      ) : null}
    </div>
  );
}

export function AttachmentList({ attachments, onRemove, label = "Attached files" }: { attachments: AttachmentPreview[]; onRemove?: (id: string) => void; label?: string }): JSX.Element | null {
  if (!attachments.length) return null;
  return <div className="attachment-list" aria-label={label}>{attachments.map(file => <div className="attachment-chip" key={file.id}><span className="attachment-icon">{file.mimeType === "application/pdf" ? <FileText size={13} /> : <ImageSquare size={13} />}</span><span className="attachment-copy"><strong>{file.name}</strong><small>{formatBytes(file.size)}</small></span>{onRemove ? <button type="button" className="attachment-remove" onClick={() => onRemove(file.id)} aria-label={`Remove ${file.name}`} title={`Remove ${file.name}`}><X size={12} /></button> : null}</div>)}</div>;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export type RecentTask = { id: string; title: string; updatedAt: string };

function taskTimeLabel(updatedAt: string): string {
  const timestamp = new Date(updatedAt).getTime();
  if (!Number.isFinite(timestamp)) return "";
  const seconds = Math.max(0, Math.floor((Date.now() - timestamp) / 1000));
  if (seconds < 60) return "Just now";
  if (seconds < 3600) return `${Math.floor(seconds / 60)}m ago`;
  if (seconds < 86400) return `${Math.floor(seconds / 3600)}h ago`;
  return new Date(updatedAt).toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

export function RecentTasksDrawer({ tasks, activeTaskId, onClose, onSelect, onNewTask, onDelete, onClearAll }: { tasks: RecentTask[]; activeTaskId: string; onClose: () => void; onSelect: (id: string) => void; onNewTask: () => void; onDelete: (id: string) => void; onClearAll: () => void }): JSX.Element {
  const [confirmClear, setConfirmClear] = useState(false);

  return (
    <div className="task-drawer-layer">
      <button type="button" className="task-drawer-backdrop" aria-label="Close chat history" onClick={onClose} />
      <section className="task-drawer" role="dialog" aria-modal="true" aria-labelledby="chat-history-title">
        <header className="task-drawer-header">
          <div className="task-drawer-heading">
            <span className="task-drawer-heading-icon"><img src="/chat-history-icon.png" alt="" aria-hidden="true" /></span>
            <div>
              <h2 id="chat-history-title">Chat history</h2>
              <p>Your conversations on this device</p>
            </div>
          </div>
          <button type="button" className="icon-button task-drawer-close" onClick={onClose} aria-label="Close chat history" title="Close"><X size={17} /></button>
        </header>

        <div className="task-drawer-toolbar"><span>Conversations</span><span className="task-drawer-count">{tasks.length}</span></div>

        <div className="task-drawer-list">
          {tasks.length ? tasks.map(task => (
            <div key={task.id} className={`task-row-shell ${task.id === activeTaskId ? "active" : ""}`}>
              <button type="button" className="task-row" onClick={() => onSelect(task.id)}>
                <span className="task-row-dot" />
                <span className="task-row-copy"><strong>{task.title}</strong><small>{taskTimeLabel(task.updatedAt)}</small></span>
                {task.id === activeTaskId ? <span className="task-row-current">Open</span> : null}
              </button>
              <button type="button" className="task-row-delete" onClick={() => onDelete(task.id)} aria-label={`Delete task ${task.title}`} title="Delete task">
                <Trash size={14} />
              </button>
            </div>
          )) : (
            <div className="task-empty">
              <span className="task-empty-icon"><img src="/chat-history-icon.png" alt="" aria-hidden="true" /></span>
              <strong>No chat history yet</strong>
              <span>Start a conversation and it will appear here.</span>
            </div>
          )}
        </div>

        <footer className="task-drawer-footer">
          <button type="button" className="new-task-button task-drawer-new" onClick={onNewTask}><Plus size={15} weight="bold" /><span>New chat</span></button>
          <div className="task-drawer-footer-meta">
            {tasks.length ? (confirmClear ? <div className="task-clear-confirm" role="group" aria-label="Confirm clear history"><span>Clear all?</span><button type="button" className="task-clear-cancel" onClick={() => setConfirmClear(false)}>Cancel</button><button type="button" className="task-clear-confirm-button" onClick={() => { setConfirmClear(false); onClearAll(); }}>Clear</button></div> : <button type="button" className="task-clear-button" onClick={() => setConfirmClear(true)}>Clear all</button>) : null}
            <span>{tasks.length ? `${tasks.length} saved ${tasks.length === 1 ? "task" : "tasks"}` : "Local history"}</span>
          </div>
        </footer>
      </section>
    </div>
  );
}
