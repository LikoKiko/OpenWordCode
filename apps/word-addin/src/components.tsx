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
  Hand,
  ImageSquare,
  Plus,
  Sparkle,
  TerminalWindow,
  Trash,
  UploadSimple,
  X,
  Warning,
} from "@phosphor-icons/react";
import type { AgentAction, AgentMode, ChatAttachment, ModelInfo, SkillSummary } from "../../../packages/shared/src/index.js";

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
          <span className="tools-menu-icon">{item.icon || <Sparkle size={15} />}</span>
          <span className="tools-menu-copy"><strong>{item.label}</strong>{item.detail ? <small>{item.detail}</small> : null}</span>
          {item.shortcut ? <span className="tools-menu-shortcut">{item.shortcut}</span> : null}
          {item.checked ? <Check size={13} weight="bold" /> : null}
          {hasChildren ? <CaretRight className={expanded === item.label ? "expanded" : ""} size={13} /> : null}
        </button>
        {hasChildren && expanded === item.label ? <div className="tools-menu-submenu">{item.children!.map(child => renderItem(child, true))}</div> : null}
      </div>
    );
  };

  return (
    <div className="tools-menu" role="dialog" aria-label="Chat tools">
      <div className="picker-header">Actions & tools</div>
      {items.map(item => renderItem(item))}
    </div>
  );
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
  { value: "manual", label: "Ask before each action", detail: "Review changes before applying" },
  { value: "auto", label: "Auto-approve safe actions", detail: "Pause only on unsafe actions" },
  { value: "skip", label: "Skip all approvals", detail: "Apply every action directly" },
];

function ApprovalModeIcon({ mode, size = 15 }: { mode: AgentMode; size?: number }): JSX.Element {
  if (mode === "manual") return <Hand size={size} weight="regular" />;
  if (mode === "auto") return <FastForward size={size} weight="regular" />;
  return <Warning size={size} weight="regular" />;
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
      <button className="picker-trigger approval-trigger" type="button" onClick={() => menu.setOpen(current => !current)} disabled={disabled} aria-haspopup="listbox" aria-expanded={menu.open} aria-label={`Approval mode: ${selected.label}`} title={selected.label}>
        <ApprovalModeIcon mode={mode} size={15} />
      </button>
      {menu.open ? (
        <div className="picker-menu approval-menu" role="listbox" aria-label="Approval mode">
          <div className="picker-header">Approval mode</div>
          {APPROVAL_MODES.map(option => <button key={option.value} type="button" role="option" aria-selected={option.value === mode} className={`picker-option approval-option ${option.value === mode ? "selected" : ""}`} onClick={() => { onChange(option.value); menu.setOpen(false); }}><span className="approval-option-leading"><ApprovalModeIcon mode={option.value} size={15} /></span><span className="picker-option-copy"><strong>{option.label}</strong><small>{option.detail}</small></span>{option.value === mode ? <Check size={14} weight="bold" /> : null}</button>)}
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
        <span className="picker-option-copy"><strong>{name}{requiresCredits ? <span className="credit-badge">Credits</span> : null}</strong>{featured ? <small>{MODEL_DESCRIPTIONS[index] ?? (model.capabilities.vision ? "Vision-ready" : "Provider model")}</small> : null}</span>
        {model.id === selectedModelId ? <Check size={14} weight="bold" /> : null}
      </button>
    );
  };
  return (
    <div className="picker model-picker" ref={menu.ref}>
      <button className="picker-trigger model-trigger" type="button" onClick={() => menu.setOpen(current => !current)} disabled={disabled} aria-haspopup="listbox" aria-expanded={menu.open} aria-label={`Model: ${selected?.name || label}`}>
        <span className="picker-trigger-label model-trigger-label">
          <span className="model-name">{loading ? "Loading…" : selected?.name || label}</span>
          <span className="model-effort-tag">{effortLabel}</span>
        </span>
        <CaretDown size={11} className="picker-caret" />
      </button>
      {menu.open ? (
        <div className="picker-menu model-menu" role="listbox" aria-label="Model">
          <div className="model-options">
            <div className="picker-header">Models</div>
            {featuredModels.length ? featuredModels.map((model, index) => renderModel(model, index, true)) : <div className="picker-empty">No models found</div>}
            
            <div className="picker-divider" />
            <div className="effort-section">
              <span className="effort-title">Thinking effort</span>
              <div className="effort-segmented">
                {EFFORTS.map(option => (
                  <button
                    type="button"
                    key={option.value}
                    className={`effort-btn ${option.value === effort ? "active" : ""}`}
                    onClick={() => onEffortChange(option.value)}
                  >
                    {option.label}
                  </button>
                ))}
              </div>
            </div>

            {moreModels.length > 0 ? (
              <>
                <div className="picker-divider" />
                <button type="button" className="more-models-toggle" onClick={() => setShowMoreModels(current => !current)} aria-expanded={showMoreModels}>
                  <span>Other models ({moreModels.length})</span>
                  <CaretDown className={showMoreModels ? "expanded" : ""} size={13} />
                </button>
                {showMoreModels ? moreModels.map(model => renderModel(model, 0, false)) : null}
              </>
            ) : null}
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

export const DEFAULT_SKILLS: SkillSummary[] = [
  {
    id: "docx-precision",
    name: "DOCX Precision & Formatting",
    description: "Word typography, table styling, heading hierarchy, and structural layout.",
    instructions: "When proposing or applying Word document edits, enforce clean formatting: use proper heading styles (Heading 1, 2, 3), clean line spacing, styled tables with highlighted headers, and clear visual element anchoring. Always choose precise Word.Range operations.",
    isDefault: true,
    enabled: true,
  },
  {
    id: "contract-redlining",
    name: "Legal & Contract Redlining",
    description: "Clause analysis, defined term consistency, risk flagging, and formal redlines.",
    instructions: "Analyze legal documents and agreements with rigorous precision. Verify that defined terms are capitalized consistently. Flag ambiguous, unilateral, or excessive liability clauses. When editing clauses, preserve section numbering, cross-references, and provide clear legal rationales for each tracked modification.",
    isDefault: true,
    enabled: false,
  },
  {
    id: "executive-briefing",
    name: "Executive Summary & Briefing",
    description: "High-impact briefings with key takeaways, risks, and decisive action items.",
    instructions: "Distill document content for senior leadership: lead with a concise 2-sentence executive takeaway, followed by structured bullet points categorized under Key Findings, Critical Risks, and Immediate Action Items with clear next steps.",
    isDefault: true,
    enabled: false,
  },
  {
    id: "academic-technical",
    name: "Technical & Report Polish",
    description: "Professional tone, rigorous terminology, structured tables, and clarity.",
    instructions: "Ensure technical precision and professional tone: eliminate colloquialisms, strengthen argument flow, format analytical data into clean tables with headers, and structure complex points into numbered recommendations.",
    isDefault: true,
    enabled: false,
  },
];

export function parseSkillFile(content: string, filename: string): Omit<SkillSummary, "id"> {
  const trimmed = content.trim();
  if (trimmed.startsWith("---")) {
    const secondDelimiter = trimmed.indexOf("---", 3);
    if (secondDelimiter > 0) {
      const meta = trimmed.slice(3, secondDelimiter);
      const body = trimmed.slice(secondDelimiter + 3).trim();
      const nameLine = meta.split("\n").find(line => line.trim().startsWith("name:"));
      const descLine = meta.split("\n").find(line => line.trim().startsWith("description:"));
      const name = nameLine ? nameLine.replace("name:", "").replace(/["']/g, "").trim() : filename.replace(/\.[^/.]+$/, "");
      const description = descLine ? descLine.replace("description:", "").replace(/["']/g, "").trim() : "Custom uploaded AI skill";
      return { name: name || "Custom skill", description, instructions: body || trimmed, enabled: true };
    }
  }
  if (trimmed.startsWith("#")) {
    const firstLineEnd = trimmed.indexOf("\n");
    const name = (firstLineEnd > 0 ? trimmed.slice(0, firstLineEnd) : trimmed).replace(/^#+\s*/, "").trim();
    const rest = firstLineEnd > 0 ? trimmed.slice(firstLineEnd).trim() : trimmed;
    return { name: name || "Custom skill", description: "Custom AI skill", instructions: rest || trimmed, enabled: true };
  }
  const fallbackName = filename.replace(/\.[^/.]+$/, "").replace(/[-_]+/g, " ");
  return {
    name: fallbackName ? fallbackName.charAt(0).toUpperCase() + fallbackName.slice(1) : "Custom skill",
    description: "Custom uploaded AI skill prompt",
    instructions: trimmed,
    enabled: true,
  };
}

export function SkillsDrawer({
  skills,
  onToggleSkill,
  onUploadSkill,
  onCreateSkill,
  onDeleteSkill,
  onClose,
}: {
  skills: SkillSummary[];
  onToggleSkill: (id: string) => void;
  onUploadSkill: (file: File) => Promise<void>;
  onCreateSkill: (skill: Omit<SkillSummary, "id">) => void;
  onDeleteSkill: (id: string) => void;
  onClose: () => void;
}): JSX.Element {
  const [isCreating, setIsCreating] = useState(false);
  const [expandedSkillId, setExpandedSkillId] = useState<string | null>(null);
  const [newName, setNewName] = useState("");
  const [newDesc, setNewDesc] = useState("");
  const [newInstructions, setNewInstructions] = useState("");
  const uploadInputRef = useRef<HTMLInputElement | null>(null);

  const activeCount = skills.filter(s => s.enabled !== false).length;

  const handleCreateSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newName.trim() || !newInstructions.trim()) return;
    onCreateSkill({
      name: newName.trim(),
      description: newDesc.trim() || "Custom AI skill recipe",
      instructions: newInstructions.trim(),
      enabled: true,
    });
    setNewName("");
    setNewDesc("");
    setNewInstructions("");
    setIsCreating(false);
  };

  return (
    <div className="task-drawer-layer">
      <button type="button" className="task-drawer-backdrop" aria-label="Close skills manager" onClick={onClose} />
      <section className="task-drawer skills-drawer" role="dialog" aria-modal="true" aria-labelledby="skills-title">
        <header className="task-drawer-header">
          <div className="task-drawer-heading">
            <span className="skills-heading-icon"><Sparkle size={20} weight="fill" /></span>
            <div>
              <h2 id="skills-title">AI Skills</h2>
              <p>Pre-packaged prompt recipes loaded into the AI</p>
            </div>
          </div>
          <button type="button" className="icon-button task-drawer-close" onClick={onClose} aria-label="Close skills" title="Close"><X size={17} /></button>
        </header>

        <div className="skills-toolbar">
          <div className="skills-status">
            <span className="skills-active-dot" />
            <strong>{activeCount} active {activeCount === 1 ? "skill" : "skills"}</strong>
          </div>
          <div className="skills-toolbar-actions">
            <input
              ref={uploadInputRef}
              type="file"
              accept=".md,.txt,.skill,.json,text/plain,text/markdown"
              className="file-input"
              onChange={async (e) => {
                const file = e.target.files?.[0];
                if (file) {
                  await onUploadSkill(file);
                  e.target.value = "";
                }
              }}
            />
            <button type="button" className="skills-action-btn" onClick={() => uploadInputRef.current?.click()} title="Upload SKILL.md or prompt file">
              <UploadSimple size={13} weight="bold" /><span>Upload</span>
            </button>
            <button type="button" className={`skills-action-btn ${isCreating ? "active" : ""}`} onClick={() => setIsCreating(curr => !curr)} title="Create custom skill">
              <Plus size={13} weight="bold" /><span>New</span>
            </button>
          </div>
        </div>

        {isCreating ? (
          <form className="skill-create-form" onSubmit={handleCreateSubmit}>
            <div className="skill-create-header">
              <strong>Create custom skill</strong>
              <button type="button" className="icon-button" onClick={() => setIsCreating(false)}><X size={14} /></button>
            </div>
            <input
              type="text"
              placeholder="Skill name (e.g. Legal Redlining)"
              value={newName}
              onChange={e => setNewName(e.target.value)}
              required
              autoFocus
            />
            <input
              type="text"
              placeholder="Short description / purpose"
              value={newDesc}
              onChange={e => setNewDesc(e.target.value)}
            />
            <textarea
              placeholder="Detailed prompt instructions / recipe for the AI..."
              value={newInstructions}
              onChange={e => setNewInstructions(e.target.value)}
              rows={4}
              required
            />
            <div className="skill-create-actions">
              <button type="button" className="text-button" onClick={() => setIsCreating(false)}>Cancel</button>
              <button type="submit" className="apply-button" disabled={!newName.trim() || !newInstructions.trim()}>Save skill</button>
            </div>
          </form>
        ) : null}

        <div className="task-drawer-list skills-list">
          {skills.map(skill => {
            const isActive = skill.enabled !== false;
            const isExpanded = expandedSkillId === skill.id;
            return (
              <article className={`skill-card ${isActive ? "active" : "inactive"}`} key={skill.id}>
                <div className="skill-card-main">
                  <button
                    type="button"
                    className={`skill-toggle-switch ${isActive ? "on" : "off"}`}
                    onClick={() => onToggleSkill(skill.id)}
                    aria-pressed={isActive}
                    title={isActive ? "Disable skill" : "Enable skill"}
                  >
                    <span className="skill-switch-knob" />
                  </button>
                  <div className="skill-card-content" onClick={() => setExpandedSkillId(curr => curr === skill.id ? null : skill.id)}>
                    <div className="skill-card-title-row">
                      <strong>{skill.name}</strong>
                      {skill.isDefault ? <span className="skill-badge-default">Built-in</span> : <span className="skill-badge-custom">Custom</span>}
                    </div>
                    <p className="skill-card-desc">{skill.description}</p>
                  </div>
                  {!skill.isDefault ? (
                    <button
                      type="button"
                      className="task-row-delete skill-delete-btn"
                      onClick={() => onDeleteSkill(skill.id)}
                      aria-label={`Delete skill ${skill.name}`}
                      title="Delete skill"
                    >
                      <Trash size={13} />
                    </button>
                  ) : null}
                </div>
                {isExpanded ? (
                  <div className="skill-instructions-preview">
                    <span className="skill-instructions-label">Prompt Instructions:</span>
                    <pre>{skill.instructions}</pre>
                  </div>
                ) : null}
              </article>
            );
          })}
        </div>

        <footer className="task-drawer-footer">
          <button type="button" className="text-button" onClick={onClose}>Done</button>
          <span>Active skills are injected into the agent prompt automatically.</span>
        </footer>
      </section>
    </div>
  );
}
