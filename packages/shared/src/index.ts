export const OPENWORDCODE_VERSION = "0.1.0";

export type AuthMethod =
  | "api-key"
  | "environment"
  | "none"
  | "existing-session"
  | "oauth";

export type AuthStatus =
  | "not-configured"
  | "detecting"
  | "login-required"
  | "connected"
  | "expired"
  | "refreshing"
  | "error"
  | "unsupported";

export type ProviderKind =
  | "openai-compatible"
  | "openwordcode-bridge"
  | "openai-codex"
  | "anthropic"
  | "gemini"
  | "google-antigravity"
  | "demo";

export type ProviderId =
  | "openai"
  | "anthropic"
  | "openrouter"
  | "google"
  | "xai"
  | "kimi"
  | "openwordcode-bridge"
  | "demo"
  | (string & {});

export interface ProviderAuthConfig {
  method: AuthMethod;
  credentialRef?: string;
  /** Encrypted reference used for an OAuth refresh/access-token pair. */
  oauthCredentialRef?: string;
  /** Provider-specific OAuth flow id, for example `anthropic` or `google-antigravity`. */
  oauthProvider?: string;
  envVar?: string;
}

export interface ProviderConfig {
  id: ProviderId;
  displayName: string;
  kind: ProviderKind;
  baseUrl: string;
  enabled: boolean;
  local: boolean;
  auth: ProviderAuthConfig;
  defaultModel?: string;
  privacyNote: string;
  /** Internal routing providers are never shown in the task-pane provider list. */
  internal?: boolean;
}

export interface ProviderAuthInfo {
  status: AuthStatus;
  method: AuthMethod;
  detail: string;
  availableMethods: AuthMethod[];
  credentialConfigured: boolean;
  environmentConfigured: boolean;
  source?: "openwordcode-account" | "codex-cli" | "oauth" | "claude-cli" | "kimi-cli" | "antigravity-cli";
}

export interface ProviderSummary extends ProviderConfig {
  auth: ProviderAuthInfo;
  modelCount?: number;
  lastError?: string;
}

export interface ModelInfo {
  id: string;
  providerId: ProviderId;
  name: string;
  contextWindow?: number;
  capabilities: {
    streaming: boolean;
    tools: boolean;
    reasoning?: boolean;
    vision?: boolean;
  };
}

export type ChatRole = "system" | "user" | "assistant" | "tool";

export interface ChatMessage {
  role: ChatRole;
  content: string;
  toolCallId?: string;
  toolCalls?: ToolCall[];
}

export interface ChatAttachment {
  id: string;
  name: string;
  mimeType: "application/pdf" | "image/gif" | "image/jpeg" | "image/png" | "image/webp";
  size: number;
  dataUrl: string;
}

export interface SkillSummary {
  id: string;
  name: string;
  description: string;
  instructions: string;
  isDefault?: boolean;
  enabled?: boolean;
  author?: string;
  version?: string;
}

export interface ToolCall {
  id: string;
  name: string;
  arguments: string;
}

export interface ChatToolDefinition {
  type: "function";
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

export interface DocumentParagraph {
  id: string;
  index: number;
  text: string;
  style?: string;
}

export type DocumentVisualElementKind = "inlinePicture" | "shape";

/**
 * A visual element discovered inside the active Word document. Inline
 * pictures have a stable document-order/paragraph location; floating shapes
 * can additionally expose point-based X/Y geometry on newer desktop hosts.
 */
export interface DocumentVisualElement {
  id: string;
  kind: DocumentVisualElementKind;
  index: number;
  shapeId?: number;
  shapeType?: string;
  name?: string;
  altTextTitle?: string;
  altTextDescription?: string;
  hyperlink?: string;
  imageFormat?: string;
  mimeType?: Exclude<ChatAttachment["mimeType"], "application/pdf">;
  width?: number;
  height?: number;
  /** Floating-shape horizontal position in Word points. */
  x?: number;
  /** Floating-shape vertical position in Word points. */
  y?: number;
  rotation?: number;
  relativeHorizontalPosition?: string;
  relativeVerticalPosition?: string;
  wrapType?: string;
  anchorParagraphIndex?: number;
  anchorText?: string;
  rangeStart?: number;
  rangeEnd?: number;
  /** Present only for a bounded set of decodable embedded pictures. */
  dataUrl?: string;
  size?: number;
  contentAvailable: boolean;
  contentOmittedReason?: string;
}

export interface DocumentTarget {
  kind: "selection" | "paragraph" | "document" | "visual";
  id: string;
  beforeText: string;
  beforeFingerprint: string;
  /** Stable body paragraph position; the text fingerprint can change when Word normalizes markup. */
  paragraphIndex?: number;
  visualKind?: DocumentVisualElementKind;
  visualIndex?: number;
  shapeId?: number;
}

export interface DocumentSnapshot {
  documentId: string;
  title?: string;
  selection: {
    text: string;
    isEmpty: boolean;
    /** Word can report an empty text range while the user has selected a whole table. */
    isTable?: boolean;
    tableCount?: number;
    rangeStart?: number;
    rangeEnd?: number;
    selectedVisualElementIds?: string[];
    target: DocumentTarget;
  };
  documentText: string;
  paragraphs: DocumentParagraph[];
  visualElements?: DocumentVisualElement[];
  visualContentTruncated?: boolean;
  outline: Array<{ id: string; text: string; level: number; index: number }>;
  capabilities: {
    canRead: boolean;
    canWrite: boolean;
    canComment: boolean;
    canFormat: boolean;
    host?: string;
    platform?: string;
  };
  truncated?: boolean;
}

/**
 * Names from the Word.Range surface. The adapter keeps these names in one
 * allowlist so the agent can request a capability without receiving an
 * Office.js object or arbitrary executable code.
 */
export const WORD_RANGE_PROPERTIES = [
  "bold", "boldBidirectional", "bookmarks", "borders", "case", "characterWidth", "combineCharacters", "conflicts", "contentControls", "context", "disableCharacterSpaceGrid", "editors", "emphasisMark", "end", "endnotes", "fields", "fitTextWidth", "font", "footnotes", "frames", "grammarChecked", "hasNoProofing", "highlightColorIndex", "horizontalInVertical", "hyperlink", "hyperlinks", "id", "inlinePictures", "isEmpty", "isEndOfRowMark", "isTextVisibleOnScreen", "italic", "italicBidirectional", "kana", "languageDetected", "languageId", "languageIdFarEast", "languageIdOther", "listFormat", "lists", "pages", "paragraphs", "parentBody", "parentContentControl", "parentContentControlOrNullObject", "parentTable", "parentTableCell", "parentTableCellOrNullObject", "parentTableOrNullObject", "revisions", "sections", "shading", "shapes", "showAll", "spellingChecked", "start", "storyLength", "storyType", "style", "styleBuiltIn", "tableColumns", "tables", "text", "twoLinesInOne", "underline",
] as const;

export const WORD_RANGE_METHODS = [
  "clear", "compareLocationWith", "delete", "detectLanguage", "expandTo", "expandToOrNullObject", "getBookmarks", "getComments", "getContentControls", "getHtml", "getHyperlinkRanges", "getNextTextRange", "getNextTextRangeOrNullObject", "getOoxml", "getRange", "getReviewedText", "getTextRanges", "getTrackedChanges", "highlight", "insertBookmark", "insertBreak", "insertCanvas", "insertComment", "insertContentControl", "insertEndnote", "insertField", "insertFileFromBase64", "insertFootnote", "insertGeometricShape", "insertHtml", "insertInlinePictureFromBase64", "insertOoxml", "insertParagraph", "insertPictureFromBase64", "insertTable", "insertText", "insertTextBox", "intersectWith", "intersectWithOrNullObject", "load", "removeHighlight", "search", "select", "set", "split", "toJSON", "track", "untrack",
] as const;

/**
 * Safe structured operations implemented by the Word adapter for a selected
 * Word.Table. They are deliberately separate from the public Word.Range API:
 * the model can request the operation, but it is never forwarded to an
 * Office.js object with an arbitrary method name.
 */
export const WORD_TABLE_OPERATIONS = ["fillTable"] as const;

/** Structured visual operations implemented by the Word adapter. */
export const WORD_VISUAL_OPERATIONS = [
  "cropImage",
  "editImage",
  "removeBackground",
  "resizeImage",
  "deleteImage",
  "setImageAltText",
  "moveShape",
  "rotateShape",
  "resizeShape",
  "setShapeWrap",
] as const;

export const WORD_RANGE_EVENTS = ["onCommentAdded", "onCommentChanged", "onCommentDeselected", "onCommentSelected"] as const;

export type WordRangeProperty = typeof WORD_RANGE_PROPERTIES[number];
export type WordRangeMethod = typeof WORD_RANGE_METHODS[number];
export type WordTableOperation = typeof WORD_TABLE_OPERATIONS[number];
export type WordVisualOperation = typeof WORD_VISUAL_OPERATIONS[number];
export type WordRangeOperationName = WordRangeProperty | WordRangeMethod | WordTableOperation | WordVisualOperation;

export interface WordRangeOperation {
  name: WordRangeOperationName | (string & {});
  args?: unknown[];
  /** Used when assigning a single Range property. */
  value?: unknown;
  /** Adapter routing hint for structured selections; never forwarded to Office.js. */
  scope?: "range" | "table" | "image" | "shape";
}

const WORD_RANGE_READ_ONLY_METHODS = new Set<WordRangeMethod>([
  "compareLocationWith", "detectLanguage", "expandTo", "expandToOrNullObject", "getBookmarks", "getComments", "getContentControls", "getHtml", "getHyperlinkRanges", "getNextTextRange", "getNextTextRangeOrNullObject", "getOoxml", "getRange", "getReviewedText", "getTextRanges", "getTrackedChanges", "intersectWith", "intersectWithOrNullObject", "load", "search", "split", "toJSON", "track", "untrack",
]);

const WORD_RANGE_READ_ONLY_PROPERTIES = new Set<WordRangeProperty>([
  "bookmarks", "borders", "conflicts", "contentControls", "context", "editors", "end", "endnotes", "fields", "footnotes", "frames", "hyperlinks", "id", "inlinePictures", "isEmpty", "isEndOfRowMark", "isTextVisibleOnScreen", "listFormat", "lists", "pages", "paragraphs", "parentBody", "parentContentControl", "parentContentControlOrNullObject", "parentTable", "parentTableCell", "parentTableCellOrNullObject", "parentTableOrNullObject", "revisions", "sections", "shading", "shapes", "start", "storyLength", "storyType", "tableColumns", "tables", "text",
]);

export function isKnownWordRangeOperation(name: string): name is WordRangeOperationName {
  return (WORD_RANGE_PROPERTIES as readonly string[]).includes(name)
    || (WORD_RANGE_METHODS as readonly string[]).includes(name)
    || (WORD_TABLE_OPERATIONS as readonly string[]).includes(name)
    || (WORD_VISUAL_OPERATIONS as readonly string[]).includes(name);
}

export function isWordRangeProperty(name: string): name is WordRangeProperty {
  return (WORD_RANGE_PROPERTIES as readonly string[]).includes(name);
}

export function isWordRangeMethod(name: string): name is WordRangeMethod {
  return (WORD_RANGE_METHODS as readonly string[]).includes(name);
}

export function isWordRangeMutation(name: string): boolean {
  if ((WORD_TABLE_OPERATIONS as readonly string[]).includes(name)) return true;
  if ((WORD_VISUAL_OPERATIONS as readonly string[]).includes(name)) return true;
  if (isWordRangeMethod(name)) return !WORD_RANGE_READ_ONLY_METHODS.has(name);
  if (isWordRangeProperty(name)) return !WORD_RANGE_READ_ONLY_PROPERTIES.has(name);
  return false;
}

export type ProposedChangeType = "replace_text" | "insert_text" | "format" | "comment" | "range_operation";
export type ProposedChangeStatus = "pending" | "approved" | "rejected" | "applied" | "failed";

export interface ProposedChange {
  id: string;
  type: ProposedChangeType;
  target: DocumentTarget;
  description: string;
  before?: string;
  after?: string;
  status: ProposedChangeStatus;
  createdAt: string;
  operation?: WordRangeOperation;
  failureReason?: string;
}

export type AgentMode = "manual" | "auto" | "skip";

export interface AgentRequest {
  providerId: ProviderId;
  modelId: string;
  instruction: string;
  mode: AgentMode;
  document: DocumentSnapshot;
  attachments?: ChatAttachment[];
  conversation?: ChatMessage[];
  skills?: SkillSummary[];
  tools?: {
    webSearch?: boolean;
    console?: boolean;
  };
}

export type AgentActionStatus = "pending" | "approved" | "rejected" | "completed" | "failed";

export interface AgentAction {
  id: string;
  type: "console";
  command: string;
  workingDirectory: string;
  reason: string;
  status: AgentActionStatus;
  createdAt: string;
  output?: string;
  failureReason?: string;
}

export type AgentEvent =
  | { type: "status"; message: string }
  | { type: "token"; delta: string }
  | { type: "tool"; name: string; state: "started" | "completed"; detail?: string }
  | { type: "proposal"; change: ProposedChange }
  | { type: "action"; action: AgentAction }
  | { type: "done"; answer: string; changes: ProposedChange[]; actions?: AgentAction[]; truncated?: boolean }
  | { type: "error"; code: string; message: string };

/** Small, deterministic fingerprint used by both the core and the Word adapter. */
export function textFingerprint(value: string): string {
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

/**
 * Normalizes Word's serialization-only paragraph characters for target
 * comparison. This is deliberately not used as the displayed document text.
 */
export function comparableText(value: string): string {
  return value
    .normalize("NFKC")
    .replace(/\r\n?/gu, "\n")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u2028\u2029]/gu, "")
    .replace(/[\u00a0\u00ad\u200b-\u200f\u202a-\u202e\u2060\u2066-\u2069\ufeff]/gu, "")
    .replace(/[ \t]+/gu, " ")
    .replace(/ *\n */gu, "\n")
    .trim();
}

/** Stable, compact visual fingerprint text. Pixel data is intentionally excluded. */
export function visualElementTargetText(element: DocumentVisualElement): string {
  return JSON.stringify({
    id: element.id,
    kind: element.kind,
    index: element.index,
    shapeId: element.shapeId,
    shapeType: element.shapeType,
    name: element.name,
    altTextTitle: element.altTextTitle,
    altTextDescription: element.altTextDescription,
    hyperlink: element.hyperlink,
    imageFormat: element.imageFormat,
    width: element.width,
    height: element.height,
    x: element.x,
    y: element.y,
    rotation: element.rotation,
    relativeHorizontalPosition: element.relativeHorizontalPosition,
    relativeVerticalPosition: element.relativeVerticalPosition,
    wrapType: element.wrapType,
    anchorParagraphIndex: element.anchorParagraphIndex,
    anchorText: element.anchorText,
    rangeStart: element.rangeStart,
    rangeEnd: element.rangeEnd,
  });
}

export function capText(value: string, maxChars: number): { value: string; truncated: boolean } {
  if (value.length <= maxChars) return { value, truncated: false };
  return { value: `${value.slice(0, maxChars)}\n[… truncated by OpenWordCode]`, truncated: true };
}
