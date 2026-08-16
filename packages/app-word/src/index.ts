import {
  capText,
  comparableText,
  isKnownWordRangeOperation,
  isWordRangeMethod,
  isWordRangeMutation,
  isWordRangeProperty,
  textFingerprint,
  visualElementTargetText,
  WORD_VISUAL_OPERATIONS,
  type DocumentParagraph,
  type DocumentSnapshot,
  type DocumentVisualElement,
  type ProposedChange,
  type WordRangeOperation,
} from "../../shared/src/index.js";

export interface WordApplicationAdapter {
  readonly kind: "office" | "memory";
  readSnapshot(): Promise<DocumentSnapshot>;
  applyChange(change: ProposedChange): Promise<{ success: boolean; message?: string }>;
  diagnostics(): Record<string, string | boolean | undefined>;
}

interface OfficeRangeLike {
  text: string;
  style?: string;
  start?: number;
  end?: number;
  load(properties: string): void;
  tables?: OfficeTableCollectionLike;
  inlinePictures?: OfficeInlinePictureCollectionLike;
  shapes?: OfficeShapeCollectionLike;
  parentTableOrNullObject?: OfficeTableLike;
  insertText(text: string, mode: string): void;
  insertHtml?(html: string, mode: string): unknown;
  insertTable?(rowCount: number, columnCount: number, mode: string, values?: string[][]): unknown;
  insertInlinePictureFromBase64?(base64: string, mode: string): OfficeInlinePictureLike;
}

interface OfficeClientResultLike<T> { value: T }

interface OfficeInlinePictureLike {
  width: number;
  height: number;
  altTextTitle?: string;
  altTextDescription?: string;
  hyperlink?: string;
  lockAspectRatio?: boolean;
  imageFormat?: string;
  paragraph?: OfficeRangeLike;
  load(properties: string): void;
  getBase64ImageSrc(): OfficeClientResultLike<string>;
  getRange?(location?: string): OfficeRangeLike;
  insertInlinePictureFromBase64?(base64: string, mode: string): OfficeInlinePictureLike;
  delete?(): void;
}

interface OfficeInlinePictureCollectionLike {
  items: OfficeInlinePictureLike[];
  load(properties: string): void;
}

interface OfficeShapeTextWrapLike {
  type?: string;
  load?(properties: string): void;
  set?(properties: Record<string, unknown>): void;
}

interface OfficeShapeLike {
  id: number;
  type?: string;
  name?: string;
  altTextDescription?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  rotation?: number;
  lockAspectRatio?: boolean;
  relativeHorizontalPosition?: string;
  relativeVerticalPosition?: string;
  textWrap?: OfficeShapeTextWrapLike;
  load(properties: string): void;
  delete?(): void;
  moveHorizontally?(delta: number): void;
  moveVertically?(delta: number): void;
  scaleHeight?(factor: number, relativeToOriginalSize?: boolean, scale?: string): void;
  scaleWidth?(factor: number, relativeToOriginalSize?: boolean, scale?: string): void;
  set?(properties: Record<string, unknown>): void;
}

interface OfficeShapeCollectionLike {
  items: OfficeShapeLike[];
  load(properties: string): void;
}

interface OfficeTableLike {
  isNullObject?: boolean;
  style?: string;
  styleBuiltIn?: string;
  rows?: OfficeTableRowCollectionLike;
  load(properties: string): void;
  delete?(): void;
}

interface OfficeTableCollectionLike {
  items: OfficeTableLike[];
  load(properties: string): void;
}

interface OfficeTableRowCollectionLike {
  items: OfficeTableRowLike[];
  load(properties: string): void;
}

interface OfficeTableRowLike {
  cells?: OfficeTableCellCollectionLike;
}

interface OfficeTableCellCollectionLike {
  items: OfficeTableCellLike[];
  load(properties: string): void;
}

interface OfficeTableCellLike {
  body?: OfficeTableCellBodyLike;
}

interface OfficeTableCellBodyLike {
  insertText(text: string, mode: string): unknown;
}

interface OfficeParagraphCollectionLike {
  items: OfficeRangeLike[];
  load(properties: string): void;
}

interface OfficeBodyLike extends OfficeRangeLike {
  paragraphs: OfficeParagraphCollectionLike;
  inlinePictures?: OfficeInlinePictureCollectionLike;
  shapes?: OfficeShapeCollectionLike;
}

interface OfficeDocumentLike {
  getSelection(): OfficeRangeLike;
  body: OfficeBodyLike;
  load?(properties: string): void;
  changeTrackingMode?: string;
}

interface OfficeContextLike {
  document: OfficeDocumentLike;
  sync(): Promise<void>;
}

interface OfficeApiLike {
  run<T>(callback: (context: OfficeContextLike) => Promise<T>): Promise<T>;
}

interface OfficeGlobals {
  Office?: {
    host?: string;
    platform?: string;
    context?: { requirements?: { isSetSupported?: (name: string, version: string) => boolean } };
    onReady?: (callback?: (info: { host?: string; platform?: string }) => void) => Promise<unknown>;
  };
  Word?: OfficeApiLike;
}

function officeGlobals(): OfficeGlobals {
  return globalThis as unknown as OfficeGlobals;
}

function paragraphId(index: number, text: string): string {
  return `paragraph-${index}-${textFingerprint(text)}`;
}

function tableValuesForChange(change: ProposedChange): string[][] | null {
  if (change.type !== "insert_text" || typeof change.after !== "string") return null;
  const rows = change.after.split(/\r?\n/).map(row => row.split("\t"));
  const columnCount = rows[0]?.length ?? 0;
  if (rows.length < 2 || columnCount < 2 || rows.some(row => row.length !== columnCount)) return null;
  return rows;
}

function supportsWordApi(version: string): boolean {
  const requirements = officeGlobals().Office?.context?.requirements;
  return requirements?.isSetSupported?.("WordApi", version) === true;
}

function supportsRequirementSet(name: string, version: string): boolean {
  return officeGlobals().Office?.context?.requirements?.isSetSupported?.(name, version) === true;
}

function validRangePosition(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0;
}

const MAX_VISUAL_ELEMENTS = 40;
const MAX_VISUAL_CONTENT_ITEMS = 4;
const MAX_VISUAL_ITEM_BYTES = 6 * 1024 * 1024;
const MAX_VISUAL_TOTAL_BYTES = 12 * 1024 * 1024;

type SupportedImageMime = "image/gif" | "image/jpeg" | "image/png" | "image/webp";

function imageMimeType(format: string | undefined, base64: string): SupportedImageMime | undefined {
  const dataMime = /^data:(image\/(?:gif|jpeg|png|webp));base64,/iu.exec(base64)?.[1]?.toLowerCase();
  if (dataMime === "image/gif" || dataMime === "image/jpeg" || dataMime === "image/png" || dataMime === "image/webp") return dataMime;
  const encoded = bareBase64(base64).replace(/\s/gu, "");
  if (encoded.startsWith("iVBORw0KGgo")) return "image/png";
  if (encoded.startsWith("/9j/")) return "image/jpeg";
  if (encoded.startsWith("R0lGOD")) return "image/gif";
  if (encoded.startsWith("UklGR")) return "image/webp";
  switch ((format ?? "").toLowerCase()) {
    case "gif": return "image/gif";
    case "jpeg":
    case "jpg": return "image/jpeg";
    case "png": return "image/png";
    case "webp": return "image/webp";
    default: return undefined;
  }
}

function bareBase64(value: string): string {
  const comma = value.indexOf(",");
  return value.startsWith("data:") && comma >= 0 ? value.slice(comma + 1) : value;
}

function base64ByteSize(value: string): number {
  const data = bareBase64(value).replace(/\s/gu, "");
  const padding = data.endsWith("==") ? 2 : data.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor(data.length * 3 / 4) - padding);
}

function anchorParagraphIndex(paragraphs: DocumentParagraph[], anchorText: string | undefined, fromIndex: number): number | undefined {
  if (!anchorText) return undefined;
  const normalized = comparableText(anchorText);
  const after = paragraphs.find(item => item.index >= fromIndex && comparableText(item.text) === normalized);
  return (after ?? paragraphs.find(item => comparableText(item.text) === normalized))?.index;
}

async function readOfficeVisualElements(context: OfficeContextLike, body: OfficeBodyLike, paragraphs: DocumentParagraph[]): Promise<{ elements: DocumentVisualElement[]; truncated: boolean }> {
  const elements: DocumentVisualElement[] = [];
  let contentBytes = 0;
  let contentItems = 0;
  let truncated = false;

  try {
    const collection = body.inlinePictures;
    if (collection) {
      collection.load("items");
      await context.sync();
      const pictures = collection.items.slice(0, MAX_VISUAL_ELEMENTS);
      truncated ||= collection.items.length > pictures.length;
      truncated ||= pictures.length > MAX_VISUAL_CONTENT_ITEMS;
      const formatSupported = supportsRequirementSet("WordApiDesktop", "1.1");
      const paragraphSupported = supportsWordApi("1.2");
      const rangeSupported = supportsWordApi("1.3");
      const ranges: Array<OfficeRangeLike | undefined> = [];
      for (const picture of pictures) {
        picture.load(`width,height,altTextTitle,altTextDescription,hyperlink,lockAspectRatio${formatSupported ? ",imageFormat" : ""}`);
        if (paragraphSupported && picture.paragraph) picture.paragraph.load("text");
        const range = rangeSupported && typeof picture.getRange === "function" ? picture.getRange() : undefined;
        range?.load("start,end");
        ranges.push(range);
      }
      await context.sync();

      const imageResults = new Map<number, OfficeClientResultLike<string>>();
      for (let index = 0; index < pictures.length && contentItems < MAX_VISUAL_CONTENT_ITEMS; index += 1) {
        try {
          imageResults.set(index, pictures[index]!.getBase64ImageSrc());
          contentItems += 1;
        } catch {
          // Metadata remains useful even when this host cannot return pixels.
        }
      }
      if (imageResults.size) await context.sync();

      let paragraphCursor = 0;
      for (let index = 0; index < pictures.length; index += 1) {
        const picture = pictures[index]!;
        const imageFormat = formatSupported ? picture.imageFormat : undefined;
        const anchorText = paragraphSupported ? picture.paragraph?.text : undefined;
        const paragraphIndex = anchorParagraphIndex(paragraphs, anchorText, paragraphCursor);
        if (paragraphIndex !== undefined) paragraphCursor = paragraphIndex;
        const base64 = imageResults.get(index)?.value;
        const mimeType = base64 ? imageMimeType(imageFormat, base64) : undefined;
        const size = base64 ? base64ByteSize(base64) : undefined;
        let dataUrl: string | undefined;
        let contentOmittedReason: string | undefined;
        if (!base64) contentOmittedReason = index >= MAX_VISUAL_CONTENT_ITEMS ? "Only the first four embedded pictures are sent as vision context." : "Word did not expose this picture's pixel data.";
        else if (!mimeType) contentOmittedReason = `The ${imageFormat ?? "unknown"} image format is metadata-only.`;
        else if ((size ?? 0) > MAX_VISUAL_ITEM_BYTES) contentOmittedReason = "The embedded picture exceeds the 6 MB vision limit.";
        else if (contentBytes + (size ?? 0) > MAX_VISUAL_TOTAL_BYTES) {
          contentOmittedReason = "The document reached the 12 MB embedded-image vision limit.";
          truncated = true;
        } else {
          dataUrl = base64.startsWith("data:") ? base64 : `data:${mimeType};base64,${base64}`;
          contentBytes += size ?? 0;
        }
        const range = ranges[index];
        elements.push({
          id: `inline-picture-${index}`,
          kind: "inlinePicture",
          index,
          width: picture.width,
          height: picture.height,
          ...(picture.altTextTitle ? { altTextTitle: picture.altTextTitle } : {}),
          ...(picture.altTextDescription ? { altTextDescription: picture.altTextDescription } : {}),
          ...(picture.hyperlink ? { hyperlink: picture.hyperlink } : {}),
          ...(imageFormat ? { imageFormat } : {}),
          ...(mimeType ? { mimeType } : {}),
          ...(anchorText ? { anchorText: capText(anchorText, 500).value } : {}),
          ...(paragraphIndex !== undefined ? { anchorParagraphIndex: paragraphIndex } : {}),
          ...(validRangePosition(range?.start) ? { rangeStart: range.start } : {}),
          ...(validRangePosition(range?.end) ? { rangeEnd: range.end } : {}),
          ...(dataUrl ? { dataUrl, size } : {}),
          contentAvailable: Boolean(dataUrl),
          ...(contentOmittedReason ? { contentOmittedReason } : {}),
        });
      }
    }
  } catch {
    // Inline picture discovery must never make the entire document unreadable.
  }

  try {
    const collection = body.shapes;
    if (collection && supportsRequirementSet("WordApiDesktop", "1.2")) {
      collection.load("items/id,items/type,items/name,items/altTextDescription,items/left,items/top,items/width,items/height,items/rotation,items/relativeHorizontalPosition,items/relativeVerticalPosition");
      await context.sync();
      const remaining = Math.max(0, MAX_VISUAL_ELEMENTS - elements.length);
      const shapes = collection.items.slice(0, remaining);
      truncated ||= collection.items.length > shapes.length;
      for (const shape of shapes) shape.textWrap?.load?.("type");
      if (shapes.some(shape => shape.textWrap?.load)) await context.sync();
      for (let index = 0; index < shapes.length; index += 1) {
        const shape = shapes[index]!;
        elements.push({
          id: `shape-${shape.id}`,
          kind: "shape",
          index,
          shapeId: shape.id,
          ...(shape.type ? { shapeType: String(shape.type) } : {}),
          ...(shape.name ? { name: shape.name } : {}),
          ...(shape.altTextDescription ? { altTextDescription: shape.altTextDescription } : {}),
          ...(typeof shape.left === "number" ? { x: shape.left } : {}),
          ...(typeof shape.top === "number" ? { y: shape.top } : {}),
          ...(typeof shape.width === "number" ? { width: shape.width } : {}),
          ...(typeof shape.height === "number" ? { height: shape.height } : {}),
          ...(typeof shape.rotation === "number" ? { rotation: shape.rotation } : {}),
          ...(shape.relativeHorizontalPosition ? { relativeHorizontalPosition: String(shape.relativeHorizontalPosition) } : {}),
          ...(shape.relativeVerticalPosition ? { relativeVerticalPosition: String(shape.relativeVerticalPosition) } : {}),
          ...(shape.textWrap?.type ? { wrapType: String(shape.textWrap.type) } : {}),
          contentAvailable: false,
          contentOmittedReason: "Office.js exposes this floating shape's geometry but not its rendered pixels.",
        });
      }
    }
  } catch {
    // Floating shapes are a newer WordApiDesktop surface; keep old Word usable.
  }

  return { elements, truncated };
}

interface ImageCropDescriptor {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
  unit?: "percent" | "pixels";
}

interface ImageEditDescriptor {
  crop?: ImageCropDescriptor;
  removeBackground?: boolean;
  backgroundTolerance?: number;
  rotate?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
  brightness?: number;
  contrast?: number;
  saturation?: number;
  grayscale?: number | boolean;
}

function recordValue(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
}

function operationDescriptor(operation: WordRangeOperation): Record<string, unknown> {
  return recordValue(operation.args?.[0] ?? operation.value);
}

function finiteNumber(value: unknown): number | undefined {
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}

function boundedNumber(value: unknown, minimum: number, maximum: number, fallback: number): number {
  const number = finiteNumber(value);
  return number === undefined ? fallback : Math.min(maximum, Math.max(minimum, number));
}

function removeBackgroundPixels(drawing: CanvasRenderingContext2D, width: number, height: number, tolerance: number): void {
  const pixelCount = width * height;
  if (pixelCount > 8_000_000) throw new Error("Background removal is limited to images up to 8 million pixels.");
  const imageData = drawing.getImageData(0, 0, width, height);
  const pixels = imageData.data;
  const step = Math.max(1, Math.floor((width + height) / 80));
  let red = 0;
  let green = 0;
  let blue = 0;
  let samples = 0;
  const addSample = (x: number, y: number): void => {
    const offset = (y * width + x) * 4;
    if ((pixels[offset + 3] ?? 0) < 10) return;
    red += pixels[offset] ?? 0;
    green += pixels[offset + 1] ?? 0;
    blue += pixels[offset + 2] ?? 0;
    samples += 1;
  };
  for (let x = 0; x < width; x += step) {
    addSample(x, 0);
    addSample(x, height - 1);
  }
  for (let y = 0; y < height; y += step) {
    addSample(0, y);
    addSample(width - 1, y);
  }
  if (!samples) return;
  const background = { red: red / samples, green: green / samples, blue: blue / samples };
  const threshold = tolerance * tolerance;
  const visited = new Uint8Array(pixelCount);
  const queue = new Uint32Array(pixelCount);
  let head = 0;
  let tail = 0;
  const enqueue = (x: number, y: number): void => {
    if (x < 0 || y < 0 || x >= width || y >= height) return;
    const index = y * width + x;
    if (visited[index]) return;
    visited[index] = 1;
    queue[tail] = index;
    tail += 1;
  };
  for (let x = 0; x < width; x += 1) {
    enqueue(x, 0);
    enqueue(x, height - 1);
  }
  for (let y = 1; y < height - 1; y += 1) {
    enqueue(0, y);
    enqueue(width - 1, y);
  }
  while (head < tail) {
    const index = queue[head++] ?? 0;
    const offset = index * 4;
    const difference = ((pixels[offset] ?? 0) - background.red) ** 2 + ((pixels[offset + 1] ?? 0) - background.green) ** 2 + ((pixels[offset + 2] ?? 0) - background.blue) ** 2;
    if (difference > threshold || (pixels[offset + 3] ?? 0) < 10) continue;
    pixels[offset + 3] = 0;
    const x = index % width;
    const y = Math.floor(index / width);
    enqueue(x - 1, y);
    enqueue(x + 1, y);
    enqueue(x, y - 1);
    enqueue(x, y + 1);
  }
  drawing.putImageData(imageData, 0, 0);
}

export function imageCropBounds(width: number, height: number, crop: ImageCropDescriptor = {}): { x: number; y: number; width: number; height: number } {
  const percent = crop.unit !== "pixels";
  const convertX = (value: unknown) => percent ? width * boundedNumber(value, 0, 99, 0) / 100 : boundedNumber(value, 0, width - 1, 0);
  const convertY = (value: unknown) => percent ? height * boundedNumber(value, 0, 99, 0) / 100 : boundedNumber(value, 0, height - 1, 0);
  const left = convertX(crop.left);
  const right = convertX(crop.right);
  const top = convertY(crop.top);
  const bottom = convertY(crop.bottom);
  if (left + right >= width || top + bottom >= height) throw new Error("Crop margins remove the entire image.");
  return { x: left, y: top, width: width - left - right, height: height - top - bottom };
}

function loadBrowserImage(dataUrl: string): Promise<HTMLImageElement> {
  if (typeof Image === "undefined") return Promise.reject(new Error("Image editing is available only inside the Word add-in."));
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("The embedded image format could not be decoded by the Word add-in."));
    image.src = dataUrl;
  });
}

async function transformImageBase64(base64: string, mimeType: SupportedImageMime, edit: ImageEditDescriptor): Promise<string> {
  const sourceUrl = base64.startsWith("data:") ? base64 : `data:${mimeType};base64,${base64}`;
  const image = await loadBrowserImage(sourceUrl);
  const sourceWidth = image.naturalWidth || image.width;
  const sourceHeight = image.naturalHeight || image.height;
  if (!sourceWidth || !sourceHeight) throw new Error("Word returned an image without pixel dimensions.");
  const crop = imageCropBounds(sourceWidth, sourceHeight, edit.crop);
  const rotation = boundedNumber(edit.rotate, -3600, 3600, 0) * Math.PI / 180;
  const cos = Math.abs(Math.cos(rotation));
  const sin = Math.abs(Math.sin(rotation));
  const outputWidth = Math.max(1, Math.ceil(crop.width * cos + crop.height * sin));
  const outputHeight = Math.max(1, Math.ceil(crop.width * sin + crop.height * cos));
  if (outputWidth * outputHeight > 40_000_000) throw new Error("The transformed image is too large to process safely inside Word.");
  const canvas = document.createElement("canvas");
  canvas.width = outputWidth;
  canvas.height = outputHeight;
  const drawing = canvas.getContext("2d");
  if (!drawing) throw new Error("The Word add-in could not start its image editor.");
  const grayscale = typeof edit.grayscale === "boolean" ? (edit.grayscale ? 100 : 0) : boundedNumber(edit.grayscale, 0, 100, 0);
  drawing.filter = [
    `brightness(${boundedNumber(edit.brightness, 0, 300, 100)}%)`,
    `contrast(${boundedNumber(edit.contrast, 0, 300, 100)}%)`,
    `saturate(${boundedNumber(edit.saturation, 0, 300, 100)}%)`,
    `grayscale(${grayscale}%)`,
  ].join(" ");
  drawing.translate(outputWidth / 2, outputHeight / 2);
  drawing.rotate(rotation);
  drawing.scale(edit.flipHorizontal ? -1 : 1, edit.flipVertical ? -1 : 1);
  drawing.drawImage(image, crop.x, crop.y, crop.width, crop.height, -crop.width / 2, -crop.height / 2, crop.width, crop.height);
  if (edit.removeBackground) removeBackgroundPixels(drawing, outputWidth, outputHeight, boundedNumber(edit.backgroundTolerance, 10, 120, 45));
  const outputMime = edit.removeBackground ? "image/png" : mimeType === "image/jpeg" ? "image/jpeg" : "image/png";
  return canvas.toDataURL(outputMime, outputMime === "image/jpeg" ? 0.92 : undefined);
}

async function applyInlinePictureOperation(context: OfficeContextLike, body: OfficeBodyLike, change: ProposedChange, operation: WordRangeOperation, selectedPicture?: OfficeInlinePictureLike): Promise<{ success: boolean; message?: string }> {
  const collection = body.inlinePictures;
  if (!collection) return { success: false, message: "This Word version does not expose embedded pictures to add-ins." };
  let picture = selectedPicture;
  if (!picture) {
    collection.load("items");
    await context.sync();
    const index = change.target.visualIndex ?? Number(/inline-picture-(\d+)/u.exec(change.target.id)?.[1] ?? -1);
    picture = index >= 0 ? collection.items[index] : undefined;
  }
  if (!picture) return { success: false, message: "The target picture moved or was removed. Ask again so OpenWordCode can refresh its location." };
  const formatSupported = supportsRequirementSet("WordApiDesktop", "1.1");
  picture.load(`width,height,altTextTitle,altTextDescription,hyperlink,lockAspectRatio${formatSupported ? ",imageFormat" : ""}`);
  await context.sync();
  const descriptor = operationDescriptor(operation);

  if (operation.name === "deleteImage") {
    if (typeof picture.delete !== "function") return { success: false, message: "This Word version does not expose picture deletion." };
    picture.delete();
    return { success: true };
  }
  if (operation.name === "resizeImage") {
    const width = finiteNumber(descriptor.width);
    const height = finiteNumber(descriptor.height);
    if (width === undefined && height === undefined) return { success: false, message: "resizeImage requires width and/or height in Word points." };
    if ("lockAspectRatio" in descriptor) picture.lockAspectRatio = Boolean(descriptor.lockAspectRatio);
    if (width !== undefined) picture.width = boundedNumber(width, 1, 2000, picture.width);
    if (height !== undefined) picture.height = boundedNumber(height, 1, 2000, picture.height);
    return { success: true };
  }
  if (operation.name === "setImageAltText") {
    if (typeof descriptor.title === "string") picture.altTextTitle = descriptor.title.slice(0, 512);
    if (typeof descriptor.description === "string") picture.altTextDescription = descriptor.description.slice(0, 2000);
    if (typeof descriptor.hyperlink === "string") picture.hyperlink = descriptor.hyperlink.slice(0, 2000);
    return { success: true };
  }
  if (operation.name !== "cropImage" && operation.name !== "editImage" && operation.name !== "removeBackground") return { success: false, message: `${operation.name} requires a floating Word shape, not an inline picture.` };
  if (typeof picture.delete !== "function") return { success: false, message: "This Word version cannot replace an embedded picture after editing it." };

  const result = picture.getBase64ImageSrc();
  await context.sync();
  const imageFormat = formatSupported ? picture.imageFormat : undefined;
  const mimeType = imageMimeType(imageFormat, result.value);
  if (!mimeType) return { success: false, message: `The ${imageFormat ?? "embedded"} image format cannot be pixel-edited in this Word host. PNG and JPEG work best.` };
  const edit: ImageEditDescriptor = operation.name === "cropImage"
    ? { crop: descriptor as ImageCropDescriptor }
    : operation.name === "removeBackground"
    ? { removeBackground: true, ...(finiteNumber(descriptor.tolerance ?? descriptor.backgroundTolerance) !== undefined ? { backgroundTolerance: finiteNumber(descriptor.tolerance ?? descriptor.backgroundTolerance) } : {}) }
    : {
        ...(recordValue(descriptor.crop) as ImageCropDescriptor && Object.keys(recordValue(descriptor.crop)).length ? { crop: recordValue(descriptor.crop) as ImageCropDescriptor } : {}),
        ...(finiteNumber(descriptor.rotate) !== undefined ? { rotate: finiteNumber(descriptor.rotate) } : {}),
        ...(typeof descriptor.flipHorizontal === "boolean" ? { flipHorizontal: descriptor.flipHorizontal } : {}),
        ...(typeof descriptor.flipVertical === "boolean" ? { flipVertical: descriptor.flipVertical } : {}),
        ...(finiteNumber(descriptor.brightness) !== undefined ? { brightness: finiteNumber(descriptor.brightness) } : {}),
        ...(finiteNumber(descriptor.contrast) !== undefined ? { contrast: finiteNumber(descriptor.contrast) } : {}),
        ...(finiteNumber(descriptor.saturation) !== undefined ? { saturation: finiteNumber(descriptor.saturation) } : {}),
        ...(typeof descriptor.grayscale === "boolean" || finiteNumber(descriptor.grayscale) !== undefined ? { grayscale: descriptor.grayscale as number | boolean } : {}),
      };
  const transformed = await transformImageBase64(result.value, mimeType, edit);
  const transformedBase64 = bareBase64(transformed);
  const range = typeof picture.getRange === "function" ? picture.getRange() : undefined;
  const replacement = typeof picture.insertInlinePictureFromBase64 === "function"
    ? picture.insertInlinePictureFromBase64(transformedBase64, "After")
    : range?.insertInlinePictureFromBase64?.(transformedBase64, "After");
  if (!replacement) return { success: false, message: "This Word version can read the picture but cannot replace it with the edited pixels." };
  replacement.width = picture.width;
  replacement.height = picture.height;
  replacement.lockAspectRatio = picture.lockAspectRatio;
  replacement.altTextTitle = picture.altTextTitle;
  replacement.altTextDescription = picture.altTextDescription;
  replacement.hyperlink = picture.hyperlink;
  picture.delete();
  return { success: true };
}

async function applyShapeOperation(context: OfficeContextLike, body: OfficeBodyLike, change: ProposedChange, operation: WordRangeOperation): Promise<{ success: boolean; message?: string }> {
  const collection = body.shapes;
  if (!collection || !supportsRequirementSet("WordApiDesktop", "1.2")) return { success: false, message: "Floating-shape editing requires a newer Word desktop build with WordApiDesktop 1.2." };
  collection.load("items/id,items/left,items/top,items/width,items/height,items/rotation,items/altTextDescription,items/lockAspectRatio");
  await context.sync();
  const shapeId = change.target.shapeId ?? Number(/shape-(\d+)/u.exec(change.target.id)?.[1] ?? -1);
  const shape = collection.items.find(item => item.id === shapeId) ?? collection.items[change.target.visualIndex ?? -1];
  if (!shape) return { success: false, message: "The target floating shape moved or was removed. Ask again after refreshing the document." };
  const descriptor = operationDescriptor(operation);

  if (operation.name === "deleteImage") {
    if (typeof shape.delete !== "function") return { success: false, message: "This Word build does not expose shape deletion." };
    shape.delete();
    return { success: true };
  }
  if (operation.name === "setImageAltText") {
    if (typeof descriptor.description !== "string") return { success: false, message: "setImageAltText requires a description." };
    shape.altTextDescription = descriptor.description.slice(0, 2000);
    return { success: true };
  }
  if (operation.name === "moveShape") {
    const x = finiteNumber(descriptor.x);
    const y = finiteNumber(descriptor.y);
    const deltaX = finiteNumber(descriptor.deltaX);
    const deltaY = finiteNumber(descriptor.deltaY);
    if (x === undefined && y === undefined && deltaX === undefined && deltaY === undefined) return { success: false, message: "moveShape requires x/y or deltaX/deltaY in Word points." };
    if (x !== undefined) shape.left = boundedNumber(x, -20_000, 20_000, shape.left ?? 0);
    if (y !== undefined) shape.top = boundedNumber(y, -20_000, 20_000, shape.top ?? 0);
    if (deltaX !== undefined) {
      if (typeof shape.moveHorizontally === "function") shape.moveHorizontally(deltaX);
      else shape.left = (shape.left ?? 0) + deltaX;
    }
    if (deltaY !== undefined) {
      if (typeof shape.moveVertically === "function") shape.moveVertically(deltaY);
      else shape.top = (shape.top ?? 0) + deltaY;
    }
    return { success: true };
  }
  if (operation.name === "rotateShape") {
    const rotation = finiteNumber(descriptor.rotation ?? operation.args?.[0]);
    if (rotation === undefined) return { success: false, message: "rotateShape requires a rotation in degrees." };
    shape.rotation = boundedNumber(rotation, -3600, 3600, 0);
    return { success: true };
  }
  if (operation.name === "resizeShape" || operation.name === "resizeImage") {
    const width = finiteNumber(descriptor.width);
    const height = finiteNumber(descriptor.height);
    if (width === undefined && height === undefined) return { success: false, message: `${operation.name} requires width and/or height in Word points.` };
    if ("lockAspectRatio" in descriptor) shape.lockAspectRatio = Boolean(descriptor.lockAspectRatio);
    if (width !== undefined) shape.width = boundedNumber(width, 1, 2000, shape.width ?? 1);
    if (height !== undefined) shape.height = boundedNumber(height, 1, 2000, shape.height ?? 1);
    return { success: true };
  }
  if (operation.name === "setShapeWrap") {
    if (!shape.textWrap) return { success: false, message: "This Word build does not expose wrapping for this shape." };
    if (typeof shape.textWrap.set === "function") shape.textWrap.set(descriptor);
    else if (typeof descriptor.type === "string") shape.textWrap.type = descriptor.type;
    else return { success: false, message: "setShapeWrap requires a Word wrap type." };
    return { success: true };
  }
  if (operation.name === "cropImage" || operation.name === "editImage" || operation.name === "removeBackground") return { success: false, message: "Office.js exposes floating-shape geometry but not its image pixels. Convert it to an inline picture before cropping or pixel editing." };
  return { success: false, message: `Unsupported floating-shape operation: ${operation.name}` };
}

async function applyVisualOperation(context: OfficeContextLike, body: OfficeBodyLike, change: ProposedChange, operation: WordRangeOperation, selectedPicture?: OfficeInlinePictureLike): Promise<{ success: boolean; message?: string }> {
  if (change.target.visualKind === "shape" || change.target.id.startsWith("shape-")) return applyShapeOperation(context, body, change, operation);
  return applyInlinePictureOperation(context, body, change, operation, selectedPicture);
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character] ?? character);
}

function tableHtml(values: string[][]): string {
  return `<table>${values.map(row => `<tr>${row.map(cell => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}</table>`;
}

type TableFillSource = string[][] | { mode: "sequence"; start: number; step: number };

function normalizedTableValues(value: unknown): string[][] | null {
  if (!Array.isArray(value)) return null;
  const rows = value
    .filter((row): row is unknown[] => Array.isArray(row))
    .map(row => row.map(cell => String(cell ?? "")));
  return rows.length && rows.some(row => row.length > 0) ? rows : null;
}

function tableValuesFromMarkup(markup: string): string[][] | null {
  const rows: string[][] = [];
  for (const rowMatch of markup.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/giu)) {
    const rowMarkup = rowMatch[1] ?? "";
    const cells = [...rowMarkup.matchAll(/<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/giu)]
      .map(cellMatch => plainTextFromMarkup(cellMatch[1] ?? "").trim());
    if (cells.length) rows.push(cells);
  }
  return rows.length ? rows : null;
}

function tableFillSource(operation: WordRangeOperation): TableFillSource | null {
  const args = operation.args ?? [];
  if (operation.name === "fillTable") {
    const source = args[0];
    const values = normalizedTableValues(source);
    if (values) return values;
    if (source && typeof source === "object" && !Array.isArray(source)) {
      const descriptor = source as { mode?: unknown; start?: unknown; step?: unknown };
      if (descriptor.mode === "sequence") {
        const start = Number(descriptor.start ?? 1);
        const step = Number(descriptor.step ?? 1);
        if (Number.isFinite(start) && Number.isFinite(step)) return { mode: "sequence", start, step };
      }
    }
    return null;
  }
  if (operation.name === "insertTable") return normalizedTableValues(args[3]);
  if (operation.name === "insertHtml") return typeof args[0] === "string" ? tableValuesFromMarkup(args[0]) : null;
  if (operation.name === "insertText") {
    if (typeof args[0] !== "string") return null;
    return args[0].split(/\r?\n/u).map(row => row.split(/\t|\s*,\s*/u).map(cell => cell.trim()));
  }
  return null;
}

function isTableFillOperation(operation: WordRangeOperation): boolean {
  return operation.name === "fillTable" || operation.name === "insertTable" || operation.name === "insertHtml" || operation.name === "insertText";
}

async function fillOfficeTable(context: OfficeContextLike, table: OfficeTableLike, source: TableFillSource): Promise<{ success: boolean; message?: string }> {
  if (!table.rows || typeof table.rows.load !== "function") return { success: false, message: "This Word version does not expose the selected table rows. Update Office and retry." };
  try {
    table.rows.load("items");
    await context.sync();
    const rows = table.rows.items ?? [];
    if (!rows.length) return { success: false, message: "Word returned no rows for the selected table." };
    for (const row of rows) row.cells?.load("items");
    await context.sync();

    let written = 0;
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      const cells = rows[rowIndex]?.cells?.items ?? [];
      for (let columnIndex = 0; columnIndex < cells.length; columnIndex += 1) {
        const cell = cells[columnIndex];
        if (!cell?.body || typeof cell.body.insertText !== "function") return { success: false, message: "Word did not expose the body of a selected table cell." };
        const value = Array.isArray(source)
          ? source[rowIndex]?.[columnIndex] ?? ""
          : String(source.start + written * source.step);
        cell.body.insertText(value, "Replace");
        written += 1;
      }
    }
    return written ? { success: true } : { success: false, message: "Word returned a table without writable cells." };
  } catch (error) {
    return { success: false, message: `Word rejected the table-cell edit: ${error instanceof Error ? error.message : String(error)}` };
  }
}

async function selectionTableInfo(context: OfficeContextLike, selection: OfficeRangeLike): Promise<{ isTable: boolean; tableCount: number; table?: OfficeTableLike }> {
  try {
    const tables = selection.tables;
    const parentTable = selection.parentTableOrNullObject;
    if (!tables && !parentTable) return { isTable: false, tableCount: 0 };
    tables?.load("items");
    parentTable?.load("isNullObject");
    await context.sync();
    const table = tables?.items?.[0] ?? (parentTable && parentTable.isNullObject !== true ? parentTable : undefined);
    const tableCount = tables?.items?.length ?? (table ? 1 : 0);
    return { isTable: tableCount > 0 || Boolean(table), tableCount, ...(table ? { table } : {}) };
  } catch {
    // Some older Word hosts expose selection text but not table collections.
    return { isTable: false, tableCount: 0 };
  }
}

function executeOfficeRangeOperation(target: OfficeRangeLike, operation: WordRangeOperation): { success: boolean; message?: string } {
  if (!isKnownWordRangeOperation(operation.name)) return { success: false, message: `Unsupported Word.Range operation: ${operation.name}` };
  if (!isWordRangeMutation(operation.name)) return { success: false, message: `Word.Range operation ${operation.name} is read-only and cannot be applied as a change` };
  const record = target as unknown as Record<string, unknown>;
  if (isWordRangeProperty(operation.name)) {
    if (!("value" in operation)) return { success: false, message: `Word.Range property ${operation.name} requires a value` };
    try {
      const current = record[operation.name];
      if (current && typeof current === "object" && typeof (current as { set?: unknown }).set === "function" && operation.value && typeof operation.value === "object") {
        (current as { set(value: unknown): void }).set(operation.value);
      } else {
        record[operation.name] = operation.value;
      }
      return { success: true };
    } catch {
      return { success: false, message: `Word rejected the ${operation.name} property value` };
    }
  }
  if (!isWordRangeMethod(operation.name)) return { success: false, message: `Unsupported Word.Range member: ${operation.name}` };
  if (operation.name === "insertTable") {
    const rawArgs = operation.args ?? [];
    const rowCount = Math.max(1, Math.min(100, Number(rawArgs[0]) || 1));
    const colCount = Math.max(1, Math.min(100, Number(rawArgs[1]) || 1));
    const location = typeof rawArgs[2] === "string" ? rawArgs[2] : "After";
    const rawValues = Array.isArray(rawArgs[3]) ? rawArgs[3] : [];
    const values: string[][] = Array.from({ length: rowCount }, (_unused, rIdx) => {
      const row = Array.isArray(rawValues[rIdx]) ? rawValues[rIdx] : [];
      return Array.from({ length: colCount }, (_cell, cIdx) => String(row[cIdx] ?? ""));
    });
    const hasNonEmptyText = values.some(row => row.some(cell => cell.trim().length > 0));
    if (typeof target.insertTable === "function") {
      try {
        if (hasNonEmptyText) {
          target.insertTable(rowCount, colCount, location, values);
        } else {
          target.insertTable(rowCount, colCount, location);
        }
        return { success: true };
      } catch {
        // fallback to insertHtml below
      }
    }
    if (typeof target.insertHtml === "function") {
      try {
        target.insertHtml(tableHtml(values), location);
        return { success: true };
      } catch {
        // fallback
      }
    }
    if (typeof target.insertText === "function") {
      target.insertText(values.map(row => row.join("\t")).join("\n"), location);
      return { success: true };
    }
    return { success: false, message: "Word host cannot insert tables at this location" };
  }
  const member = record[operation.name];
  if (typeof member !== "function") {
    if (operation.name === "insertHtml" && typeof target.insertText === "function" && typeof operation.args?.[0] === "string") {
      const location = typeof operation.args?.[1] === "string" ? operation.args[1] as string : "After";
      target.insertText(plainTextForRangeInsertion(operation.args[0] as string), location);
      return { success: true };
    }
    return { success: false, message: `This Word version does not expose Range.${operation.name}` };
  }
  const args = operation.args ?? (operation.name === "set" && "value" in operation ? [operation.value] : []);
  try {
    member.apply(target, args);
    return { success: true };
  } catch (error) {
    return { success: false, message: error instanceof Error ? error.message : `Word rejected Range.${operation.name}` };
  }
}

function normalizeBuiltInTableStyle(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim().replace(/[–—−]/g, "-").replace(/\s+/g, " ");
  const match = /^(Grid|List)\s*Table\s*(\d+)\s*(Light|Dark|Colorful)?(?:\s*[-_]\s*)?Accent\s*(\d+)$/i.exec(normalized);
  if (match) return `${match[1]}Table${match[2]}${match[3] ?? ""}_Accent${match[4]}`;
  const compact = normalized.replace(/\s+/g, "");
  if (/^(?:TableGrid|PlainTable[1-5]|GridTable(?:1Light|[2-4]|5Dark|6Colorful|7Colorful)(?:_Accent[1-6])?|ListTable(?:1Light|[2-4]|5Dark|6Colorful|7Colorful)(?:_Accent[1-6])?)$/i.test(compact)) return compact;
  return null;
}

function executeOfficeTableOperation(table: OfficeTableLike, operation: WordRangeOperation): { success: boolean; message?: string } {
  if (operation.name === "styleBuiltIn") {
    if (!("value" in operation)) return { success: false, message: "Word.Table styleBuiltIn requires a value" };
    const style = normalizeBuiltInTableStyle(operation.value);
    if (!style) return { success: false, message: "Use a valid Word built-in table style, such as GridTable4_Accent1" };
    return executeOfficeRangeOperation(table as unknown as OfficeRangeLike, { ...operation, value: style });
  }
  if (operation.name === "style") {
    const style = normalizeBuiltInTableStyle(operation.value);
    if (style) return executeOfficeRangeOperation(table as unknown as OfficeRangeLike, { ...operation, name: "styleBuiltIn", value: style });
  }
  if (operation.name === "set" && operation.value && typeof operation.value === "object" && !Array.isArray(operation.value)) {
    const properties = { ...(operation.value as Record<string, unknown>) };
    const style = normalizeBuiltInTableStyle(properties.styleBuiltIn ?? properties.style);
    if (style) {
      delete properties.style;
      properties.styleBuiltIn = style;
      return executeOfficeRangeOperation(table as unknown as OfficeRangeLike, { ...operation, value: properties });
    }
  }
  return executeOfficeRangeOperation(table as unknown as OfficeRangeLike, operation);
}

type MemoryOperationAction = "replace" | "append" | "prepend" | "clear" | "noop";

function plainTextFromMarkup(value: string): string {
  return value.replace(/<br\s*\/?\s*>/gi, "\n").replace(/<[^>]+>/g, "").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">").replace(/&quot;/gi, '"').replace(/&#39;/gi, "'");
}

function plainTextForRangeInsertion(value: string): string {
  return value
    .replace(/<br\s*\/?\s*>/gi, "\n")
    .replace(/<\/td\s*>/gi, "\t")
    .replace(/<\/tr\s*>/gi, "\n")
    .replace(/<\/(?:p|div|li|h[1-6])\s*>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/[ \t]+\n/gu, "\n")
    .replace(/\n+$/u, "");
}

function memoryRangeOperation(operation: WordRangeOperation): { success: boolean; message?: string; action?: MemoryOperationAction; text?: string } {
  if (!isKnownWordRangeOperation(operation.name)) return { success: false, message: `Unsupported Word.Range operation: ${operation.name}` };
  if (!isWordRangeMutation(operation.name)) return { success: false, message: `Word.Range operation ${operation.name} is read-only and cannot be applied as a change` };
  const args = operation.args ?? (operation.name === "set" && "value" in operation ? [operation.value] : []);
  const location = typeof args[1] === "string" ? args[1].toLowerCase() : "after";
  const actionForLocation = location.includes("replace") ? "replace" : location.includes("before") ? "prepend" : "append";
  if (operation.name === "clear" || operation.name === "delete") return { success: true, action: "clear", text: "" };
  if (operation.name === "insertText" || operation.name === "insertParagraph" || operation.name === "insertBreak") {
    return { success: true, action: actionForLocation, text: typeof args[0] === "string" ? args[0] : operation.name === "insertBreak" ? "\n" : "" };
  }
  if (operation.name === "insertHtml" || operation.name === "insertOoxml") {
    return { success: true, action: actionForLocation, text: typeof args[0] === "string" ? plainTextFromMarkup(args[0]) : "" };
  }
  if (operation.name === "insertTable") {
    const rows = Array.isArray(args[3]) ? args[3] : Array.from({ length: Number(args[0]) || 0 }, () => Array.from({ length: Number(args[1]) || 0 }, () => ""));
    const text = rows.map(row => Array.isArray(row) ? row.map(cell => String(cell ?? "")).join("\t") : "").join("\n");
    return { success: true, action: actionForLocation, text };
  }
  if (operation.name === "set" && operation.value && typeof operation.value === "object" && "text" in operation.value) {
    const text = (operation.value as { text?: unknown }).text;
    if (typeof text === "string") return { success: true, action: "replace", text };
  }
  if (["insertBookmark", "insertCanvas", "insertComment", "insertContentControl", "insertEndnote", "insertField", "insertFileFromBase64", "insertFootnote", "insertGeometricShape", "insertInlinePictureFromBase64", "insertPictureFromBase64", "insertTextBox", "highlight", "removeHighlight", "select"].includes(operation.name)) return { success: true, action: "noop" };
  return { success: false, message: `Preview adapter cannot render Range.${operation.name} yet` };
}

function snapshotFromParts(selectionText: string, documentText: string, paragraphs: DocumentParagraph[], host?: string, platform?: string, selectionInfo?: { isTable?: boolean; tableCount?: number; rangeStart?: number; rangeEnd?: number; selectedVisualElementIds?: string[] }, visualInfo?: { elements: DocumentVisualElement[]; truncated: boolean }): DocumentSnapshot {
  const selection = capText(selectionText, 20_000);
  const fullDocument = capText(documentText, 30_000);
  const targetText = selection.value;
  return {
    documentId: "word-active-document",
    selection: {
      text: targetText,
      isEmpty: targetText.length === 0,
      ...(selectionInfo?.isTable !== undefined ? { isTable: selectionInfo.isTable } : {}),
      ...(selectionInfo?.tableCount !== undefined ? { tableCount: selectionInfo.tableCount } : {}),
      ...(validRangePosition(selectionInfo?.rangeStart) ? { rangeStart: selectionInfo.rangeStart } : {}),
      ...(validRangePosition(selectionInfo?.rangeEnd) ? { rangeEnd: selectionInfo.rangeEnd } : {}),
      ...(selectionInfo?.selectedVisualElementIds?.length ? { selectedVisualElementIds: selectionInfo.selectedVisualElementIds } : {}),
      target: { kind: "selection", id: "selection", beforeText: targetText, beforeFingerprint: textFingerprint(targetText) },
    },
    documentText: fullDocument.value,
    paragraphs: paragraphs.slice(0, 200),
    ...(visualInfo?.elements.length ? { visualElements: visualInfo.elements } : {}),
    ...(visualInfo?.truncated ? { visualContentTruncated: true } : {}),
    outline: paragraphs.filter(paragraph => /^heading\s*\d*$/i.test(paragraph.style ?? "") || /^title$/i.test(paragraph.style ?? "")).map(paragraph => ({ id: paragraph.id, text: paragraph.text, level: Number(/(\d+)/.exec(paragraph.style ?? "")?.[1] ?? 1), index: paragraph.index })),
    capabilities: { canRead: true, canWrite: true, canComment: false, canFormat: false, host, platform },
    truncated: selection.truncated || fullDocument.truncated,
  };
}

export class OfficeWordAdapter implements WordApplicationAdapter {
  readonly kind = "office" as const;

  private get word(): OfficeApiLike {
    const word = officeGlobals().Word;
    if (!word) throw new Error("Office.js Word API is unavailable");
    return word;
  }

  async readSnapshot(): Promise<DocumentSnapshot> {
    try {
      return await this.word.run(async context => {
        const selection = context.document.getSelection();
        const body = context.document.body;
        const selectionRangeSupported = supportsWordApi("1.3");
        selection.load(selectionRangeSupported ? "text,start,end" : "text");
        body.load("text");
        body.paragraphs.load("items/text,items/style");
        await context.sync();
        const tableInfo = await selectionTableInfo(context, selection);
        const paragraphs = body.paragraphs.items.map((paragraph, index) => ({ id: paragraphId(index, paragraph.text), index, text: paragraph.text, ...(paragraph.style ? { style: paragraph.style } : {}) }));
        const visualInfo = await readOfficeVisualElements(context, body, paragraphs);
        const selectionStart = selectionRangeSupported && validRangePosition(selection.start) ? selection.start : undefined;
        const selectionEnd = selectionRangeSupported && validRangePosition(selection.end) ? selection.end : undefined;
        const selectedVisualElementIds = selectionStart !== undefined && selectionEnd !== undefined
          ? visualInfo.elements.filter(element => typeof element.rangeStart === "number" && typeof element.rangeEnd === "number" && element.rangeStart >= selectionStart && element.rangeEnd <= selectionEnd).map(element => element.id)
          : [];
        return snapshotFromParts(selection.text, body.text, paragraphs, officeGlobals().Office?.host, officeGlobals().Office?.platform, {
          ...tableInfo,
          ...(selectionStart !== undefined ? { rangeStart: selectionStart } : {}),
          ...(selectionEnd !== undefined ? { rangeEnd: selectionEnd } : {}),
          ...(selectedVisualElementIds.length ? { selectedVisualElementIds } : {}),
        }, visualInfo);
      });
    } catch (error) {
      throw new Error(`Word could not provide the active document: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async applyChange(change: ProposedChange): Promise<{ success: boolean; message?: string }> {
    const after = change.after;
    const operation = change.type === "range_operation" ? change.operation : undefined;
    const legacyTextChange = (change.type === "replace_text" || change.type === "insert_text") && typeof after === "string";
    if (!legacyTextChange && (!operation || !isWordRangeMutation(operation.name))) return { success: false, message: `Word adapter does not yet apply ${change.type} changes` };
    const textAfter = after ?? "";
    const tableValues = tableValuesForChange(change);
    try {
      return await this.word.run(async context => {
        const document = context.document;
        let restoreTrackingMode: string | undefined;
        try {
          // Word 2016 builds can expose changeTrackingMode even when their
          // requirement-set probe reports an older WordApi version. Try the
          // property whenever the host exposes it, but keep the edit usable
          // on hosts that genuinely do not support it.
          if (typeof document.load === "function" && "changeTrackingMode" in document) {
            try {
              document.load("changeTrackingMode");
              await context.sync();
              if (document.changeTrackingMode && document.changeTrackingMode !== "Off") {
                restoreTrackingMode = document.changeTrackingMode;
                document.changeTrackingMode = "Off";
                await context.sync();
              }
            } catch {
              // Older Word hosts may reject the property at sync time. In
              // that case continue with the requested edit and do not claim
              // that tracking was changed.
            }
          }
          if (change.target.kind === "visual") {
            if (!operation) return { success: false, message: "Visual document elements require a structured image or shape operation." };
            const applied = await applyVisualOperation(context, document.body, change, operation);
            if (!applied.success) return applied;
            await context.sync();
            return { success: true };
          }
          let target: OfficeRangeLike;
          if (change.target.kind === "selection") {
            const selection = context.document.getSelection();
            const selectedPictures = operation && (WORD_VISUAL_OPERATIONS as readonly string[]).includes(operation.name)
              ? selection.inlinePictures
              : undefined;
            if (selectedPictures) selectedPictures.load("items");
            selection.load("text");
            await context.sync();
            if (selection.text !== (change.before ?? change.target.beforeText)) return { success: false, message: "The Word selection changed before this edit could be applied." };
            if (operation && (WORD_VISUAL_OPERATIONS as readonly string[]).includes(operation.name) && selectedPictures?.items.length) {
              const applied = await applyVisualOperation(context, document.body, change, operation, selectedPictures.items[0]);
              if (!applied.success) return applied;
              await context.sync();
              return { success: true };
            }
            const tableInfo = await selectionTableInfo(context, selection);
            const shouldUseSelectedTable = Boolean(operation && (operation.scope === "table" || (tableInfo.isTable && isTableFillOperation(operation))));
            if (shouldUseSelectedTable && operation) {
              if (!tableInfo.table) return { success: false, message: "Word did not expose the selected table. Select the table again and retry." };
              if (operation.name === "delete") {
                if (typeof tableInfo.table.delete !== "function") return { success: false, message: "This Word version does not expose table deletion." };
                tableInfo.table.delete();
              } else if (isTableFillOperation(operation)) {
                const source = tableFillSource(operation);
                if (!source) return { success: false, message: "OpenWordCode could not read the requested table values. Please specify the cell values or ask for sequential labels." };
                const filled = await fillOfficeTable(context, tableInfo.table, source);
                if (!filled.success) return filled;
              } else {
                const executed = executeOfficeTableOperation(tableInfo.table, operation);
                if (!executed.success) return executed;
              }
              await context.sync();
              return { success: true };
            }
            target = selection;
          } else if (change.target.kind === "document") {
            const body = context.document.body;
            body.load("text");
            await context.sync();
            const expected = change.before ?? change.target.beforeText;
            if (body.text !== expected && comparableText(body.text) !== comparableText(expected)) return { success: false, message: "The Word document changed before this edit could be applied. Refresh the document context and retry." };
            target = body;
          } else {
            const body = context.document.body;
            body.load("text");
            body.paragraphs.load("items/text");
            await context.sync();
            const expected = change.before ?? change.target.beforeText;
            const expectedComparable = comparableText(expected);
            const index = change.target.paragraphIndex ?? Number(/paragraph-(\d+)-/.exec(change.target.id)?.[1] ?? -1);
            let paragraph = index >= 0 ? body.paragraphs.items[index] : undefined;
            // Prefer the stable index, but recover when Word inserted/removed
            // a paragraph or normalized its paragraph terminator after the
            // suggestion was created.
            if (!paragraph || comparableText(paragraph.text) !== expectedComparable) {
              paragraph = body.paragraphs.items.find(item => comparableText(item.text) === expectedComparable);
            }
            if (!paragraph) return { success: false, message: "The target paragraph changed before this edit could be applied. Refresh the document context and try again." };
            target = paragraph;
          }
          if (operation) {
            const executed = executeOfficeRangeOperation(target, operation);
            if (!executed.success) return executed;
          } else if (tableValues) {
            if (typeof target.insertTable === "function") {
              if (target.text) target.insertText("", "Replace");
              const hasNonEmpty = tableValues.some(row => row.some(cell => cell.trim().length > 0));
              if (hasNonEmpty) target.insertTable(tableValues.length, tableValues[0]!.length, "After", tableValues);
              else target.insertTable(tableValues.length, tableValues[0]!.length, "After");
            } else if (typeof target.insertHtml === "function") {
              target.insertHtml(tableHtml(tableValues), "After");
            } else {
              return { success: false, message: "This Word host does not expose table insertion. Update Office or use WordApi 1.3+ in the desktop host." };
            }
          } else if (change.type === "replace_text") {
            target.insertText(textAfter, "Replace");
          } else {
            target.insertText(textAfter, "After");
          }
          await context.sync();
          return { success: true };
        } finally {
          if (restoreTrackingMode !== undefined) {
            document.changeTrackingMode = restoreTrackingMode;
            await context.sync();
          }
        }
      });
    } catch (error) {
      return { success: false, message: `Word rejected this edit: ${error instanceof Error ? error.message : String(error)}` };
    }
  }

  diagnostics(): Record<string, string | boolean | undefined> {
    return { officeAvailable: Boolean(officeGlobals().Office), wordApiAvailable: Boolean(officeGlobals().Word), changeTrackingAvailable: supportsWordApi("1.4"), host: officeGlobals().Office?.host, platform: officeGlobals().Office?.platform };
  }
}

export class MemoryWordAdapter implements WordApplicationAdapter {
  readonly kind = "memory" as const;
  private selectionText: string;
  private documentText: string;
  private paragraphs: DocumentParagraph[];
  private visualElements: DocumentVisualElement[];

  constructor(initial?: { selection?: string; documentText?: string; visualElements?: DocumentVisualElement[] }) {
    this.selectionText = initial?.selection ?? "The customers needs to submit the form.";
    this.documentText = initial?.documentText ?? this.selectionText;
    this.paragraphs = this.makeParagraphs(this.documentText);
    this.visualElements = initial?.visualElements ?? [];
  }

  async readSnapshot(): Promise<DocumentSnapshot> {
    return snapshotFromParts(this.selectionText, this.documentText, this.paragraphs, "Memory preview", "Browser", undefined, { elements: this.visualElements, truncated: false });
  }

  async applyChange(change: ProposedChange): Promise<{ success: boolean; message?: string }> {
    const expected = change.before ?? change.target.beforeText;
    const paragraphAtIndex = change.target.paragraphIndex !== undefined ? this.paragraphs[change.target.paragraphIndex] : undefined;
    const visual = change.target.kind === "visual" ? this.visualElements.find(item => item.id === change.target.id) ?? this.visualElements.find(item => item.kind === change.target.visualKind && item.index === change.target.visualIndex) : undefined;
    const current = change.target.kind === "selection" ? this.selectionText : change.target.kind === "document" ? this.documentText : change.target.kind === "visual" ? (visual ? visualElementTargetText(visual) : undefined) : this.paragraphs.find(item => item.id === change.target.id)?.text ?? paragraphAtIndex?.text;
    if (current !== expected) return { success: false, message: "The preview document changed before this edit could be applied." };
    const operation = change.type === "range_operation" ? change.operation : undefined;
    const legacyTextChange = (change.type === "replace_text" || change.type === "insert_text") && typeof change.after === "string";
    if (!legacyTextChange && (!operation || !isWordRangeMutation(operation.name))) return { success: false, message: `Preview adapter does not yet apply ${change.type} changes` };
    if (change.target.kind === "visual" && operation && visual) {
      const descriptor = operationDescriptor(operation);
      if (operation.name === "deleteImage") this.visualElements = this.visualElements.filter(item => item.id !== visual.id);
      else if (operation.name === "resizeImage" || operation.name === "resizeShape") {
        const width = finiteNumber(descriptor.width);
        const height = finiteNumber(descriptor.height);
        if (width !== undefined) visual.width = width;
        if (height !== undefined) visual.height = height;
      } else if (operation.name === "moveShape") {
        const x = finiteNumber(descriptor.x);
        const y = finiteNumber(descriptor.y);
        if (x !== undefined) visual.x = x;
        if (y !== undefined) visual.y = y;
      } else if (operation.name === "rotateShape") {
        const rotation = finiteNumber(descriptor.rotation);
        if (rotation !== undefined) visual.rotation = rotation;
      } else if (operation.name === "setImageAltText") {
        if (typeof descriptor.title === "string") visual.altTextTitle = descriptor.title;
        if (typeof descriptor.description === "string") visual.altTextDescription = descriptor.description;
      }
      return { success: true };
    }
    const operationResult = operation ? memoryRangeOperation(operation) : { success: true, action: change.type === "replace_text" ? "replace" as const : "append" as const, text: change.after as string };
    if (!operationResult.success) return { success: false, message: operationResult.message };
    if (operationResult.action === "noop") return { success: true };
    if (change.target.kind === "selection") {
      const replacement = operationResult.text ?? "";
      const nextSelection = operationResult.action === "replace" || operationResult.action === "clear"
        ? replacement
        : operationResult.action === "prepend" ? `${replacement}${expected}` : `${expected}${replacement}`;
      this.selectionText = nextSelection;
      this.documentText = this.documentText.replace(expected, nextSelection);
    } else if (change.target.kind === "document") {
      const replacement = operationResult.text ?? "";
      this.documentText = operationResult.action === "replace" || operationResult.action === "clear"
        ? replacement
        : operationResult.action === "prepend" ? `${replacement}${this.documentText}` : `${this.documentText}${replacement}`;
      this.selectionText = "";
    } else {
      const idIndex = this.paragraphs.findIndex(item => item.id === change.target.id);
      const index = idIndex >= 0 ? idIndex : change.target.paragraphIndex ?? -1;
      if (index < 0) return { success: false, message: "Preview paragraph no longer exists." };
      const existing = this.paragraphs[index]!.text;
      const replacement = operationResult.text ?? "";
      const nextText = operationResult.action === "replace" || operationResult.action === "clear"
        ? replacement
        : operationResult.action === "prepend" ? `${replacement}${existing}` : `${existing}${replacement}`;
      this.paragraphs[index] = { ...this.paragraphs[index]!, text: nextText };
      this.documentText = this.paragraphs.map(item => item.text).join("\n");
    }
    this.paragraphs = this.makeParagraphs(this.documentText);
    return { success: true };
  }

  diagnostics(): Record<string, string | boolean | undefined> {
    return { officeAvailable: false, wordApiAvailable: false, host: "Memory preview", platform: "Browser" };
  }

  setSelection(text: string): void { this.selectionText = text; }

  private makeParagraphs(text: string): DocumentParagraph[] {
    return text.split(/\r?\n/).map((value, index) => ({ id: paragraphId(index, value), index, text: value }));
  }
}

export function createWordAdapter(): WordApplicationAdapter {
  const globals = officeGlobals();
  return globals.Office && globals.Word ? new OfficeWordAdapter() : new MemoryWordAdapter();
}

export function isOfficeHost(): boolean {
  return Boolean(officeGlobals().Office);
}

export async function waitForOfficeReady(timeoutMs = 8_000): Promise<boolean> {
  const ready = officeGlobals().Office?.onReady;
  if (!ready) return false;
  try {
    await Promise.race([
      new Promise<void>(resolve => {
        let settled = false;
        const finish = () => { if (!settled) { settled = true; resolve(); } };
        try {
          const result = ready(() => finish());
          if (result && typeof (result as Promise<unknown>).then === "function") void (result as Promise<unknown>).then(finish, finish);
        } catch {
          finish();
        }
      }),
      new Promise<void>(resolve => { globalThis.setTimeout(resolve, timeoutMs); }),
    ]);
  } catch {
    return false;
  }
  return Boolean(officeGlobals().Word?.run);
}
