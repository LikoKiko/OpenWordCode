import { describe, expect, it } from "vitest";
import { imageCropBounds, MemoryWordAdapter, OfficeWordAdapter } from "../packages/app-word/src/index.js";
import { textFingerprint, visualElementTargetText } from "../packages/shared/src/index.js";

describe("memory Word adapter", () => {
  it("reads and applies a selection replacement", async () => {
    const adapter = new MemoryWordAdapter({ selection: "customers needs help", documentText: "customers needs help" });
    const snapshot = await adapter.readSnapshot();
    const change = { id: "change", type: "replace_text" as const, target: snapshot.selection.target, description: "grammar", before: snapshot.selection.text, after: "the customer needs help", status: "approved" as const, createdAt: new Date().toISOString() };
    const applied = await adapter.applyChange(change);
    expect(applied.success).toBe(true);
    expect((await adapter.readSnapshot()).selection.text).toBe("the customer needs help");
    expect(textFingerprint((await adapter.readSnapshot()).selection.text)).not.toBe(snapshot.selection.target.beforeFingerprint);
  });

  it("fails closed when the target is stale", async () => {
    const adapter = new MemoryWordAdapter({ selection: "old" });
    const snapshot = await adapter.readSnapshot();
    adapter.setSelection("new");
    const result = await adapter.applyChange({ id: "change", type: "replace_text", target: snapshot.selection.target, description: "test", before: "old", after: "updated", status: "approved", createdAt: new Date().toISOString() });
    expect(result.success).toBe(false);
  });

  it("applies inserted text to an empty selection", async () => {
    const adapter = new MemoryWordAdapter({ selection: "", documentText: "" });
    const snapshot = await adapter.readSnapshot();
    const result = await adapter.applyChange({ id: "change", type: "insert_text", target: snapshot.selection.target, description: "Insert text", before: "", after: "1\t2\n3\t4", status: "approved", createdAt: new Date().toISOString() });
    expect(result.success).toBe(true);
    expect((await adapter.readSnapshot()).documentText).toBe("1\t2\n3\t4");
  });

  it("applies a Word.Range operation in the memory preview", async () => {
    const adapter = new MemoryWordAdapter({ selection: "Hello", documentText: "Hello" });
    const snapshot = await adapter.readSnapshot();
    const result = await adapter.applyChange({
      id: "range-change",
      type: "range_operation",
      target: snapshot.selection.target,
      description: "Add a polite greeting",
      before: snapshot.selection.text,
      operation: { name: "insertHtml", args: ["<p> world</p>", "After"] },
      status: "approved",
      createdAt: new Date().toISOString(),
    });
    expect(result.success).toBe(true);
    expect((await adapter.readSnapshot()).documentText).toBe("Hello world");
  });

  it("calculates percentage and pixel crop bounds without removing the whole image", () => {
    expect(imageCropBounds(1000, 500, { left: 10, right: 20, top: 5, bottom: 5 })).toEqual({ x: 100, y: 25, width: 700, height: 450 });
    expect(imageCropBounds(1000, 500, { left: 40, top: 20, unit: "pixels" })).toEqual({ x: 40, y: 20, width: 960, height: 480 });
    expect(() => imageCropBounds(100, 100, { left: 60, right: 60 })).toThrow("entire image");
  });

  it("tracks and resizes a visual element in the memory preview", async () => {
    const visual = { id: "inline-picture-0", kind: "inlinePicture" as const, index: 0, width: 120, height: 80, contentAvailable: false };
    const adapter = new MemoryWordAdapter({ selection: "", documentText: "", visualElements: [visual] });
    const beforeText = visualElementTargetText(visual);
    const result = await adapter.applyChange({
      id: "resize-picture",
      type: "range_operation",
      target: { kind: "visual", id: visual.id, visualKind: visual.kind, visualIndex: visual.index, beforeText, beforeFingerprint: textFingerprint(beforeText) },
      description: "Resize picture",
      before: beforeText,
      operation: { name: "resizeImage", scope: "image", args: [{ width: 240, height: 160 }] },
      status: "approved",
      createdAt: new Date().toISOString(),
    });
    expect(result.success).toBe(true);
    expect((await adapter.readSnapshot()).visualElements?.[0]).toMatchObject({ width: 240, height: 160 });
  });

  it("routes visual operations from an image selection to the selected picture", async () => {
    const picture = { width: 120, height: 80, load: () => undefined, delete: () => undefined };
    const selectedPictures = { items: [picture], load: () => undefined };
    const selection = { text: "", load: () => undefined, inlinePictures: selectedPictures, insertText: () => undefined };
    const body = { text: "", load: () => undefined, inlinePictures: { items: [picture], load: () => undefined }, paragraphs: { items: [], load: () => undefined } };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const result = await adapter.applyChange({
        id: "selected-picture-resize",
        type: "range_operation",
        target: { kind: "selection", id: "selection", beforeText: "", beforeFingerprint: textFingerprint("") },
        description: "Resize the selected picture",
        before: "",
        operation: { name: "resizeImage", scope: "image", args: [{ width: 240, height: 160 }] },
        status: "approved",
        createdAt: new Date().toISOString(),
      });
      expect(result.success).toBe(true);
      expect(picture).toMatchObject({ width: 240, height: 160 });
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it("reads embedded Word pictures as bounded visual document context", async () => {
    const png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=";
    const imageResult = { value: png };
    const picture = {
      width: 144,
      height: 72,
      altTextTitle: "Revenue chart",
      altTextDescription: "Quarterly revenue",
      imageFormat: "png",
      paragraph: { text: "Chart paragraph", load: () => undefined, insertText: () => undefined },
      load: () => undefined,
      getBase64ImageSrc: () => imageResult,
    };
    const selection = { text: "", start: 0, end: 0, load: () => undefined, insertText: () => undefined };
    const body = {
      text: "Chart paragraph",
      load: () => undefined,
      paragraphs: { items: [{ text: "Chart paragraph", style: "Normal", load: () => undefined, insertText: () => undefined }], load: () => undefined },
      inlinePictures: { items: [picture], load: () => undefined },
    };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: (name: string, version: string) => name === "WordApi" && version === "1.2" } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const snapshot = await new OfficeWordAdapter().readSnapshot();
      expect(snapshot.visualElements?.[0]).toMatchObject({ id: "inline-picture-0", kind: "inlinePicture", width: 144, height: 72, mimeType: "image/png", anchorParagraphIndex: 0, contentAvailable: true });
      expect(snapshot.visualElements?.[0]?.imageFormat).toBeUndefined();
      expect(snapshot.visualElements?.[0]?.dataUrl).toBe(`data:image/png;base64,${png}`);
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it("falls back from unsupported Range.insertHtml to plain text insertion", async () => {
    const calls: string[][] = [];
    const selection = {
      text: "",
      load: () => undefined,
      insertText: (text: string, mode: string) => { calls.push([text, mode]); },
    };
    const body = { text: "", load: () => undefined, paragraphs: { items: [], load: () => undefined } };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const snapshot = await adapter.readSnapshot();
      const result = await adapter.applyChange({
        id: "office-html-fallback",
        type: "range_operation",
        target: snapshot.selection.target,
        description: "Insert HTML content",
        before: "",
        operation: { name: "insertHtml", args: ["<table><tr><td>1</td><td>2</td></tr></table>", "After"] },
        status: "approved",
        createdAt: new Date().toISOString(),
      });
      expect(result.success).toBe(true);
      expect(calls).toEqual([["1\t2", "After"]]);
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it("turns tracking off while applying a clean edit and restores the user's setting", async () => {
    const trackingAtWrite: string[] = [];
    let trackingMode = "TrackAll";
    const selection = {
      text: "old",
      load: () => undefined,
      insertText: () => { trackingAtWrite.push(trackingMode); },
    };
    const body = { text: "old", load: () => undefined, paragraphs: { items: [], load: () => undefined } };
    const officeDocument = {
      getSelection: () => selection,
      body,
      get changeTrackingMode() { return trackingMode; },
      set changeTrackingMode(value: string) { trackingMode = value; },
      load: () => undefined,
    };
    const context = { document: officeDocument, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    // The real Word 2016 desktop host may expose this property while
    // reporting an older requirement set.
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const snapshot = await adapter.readSnapshot();
      const result = await adapter.applyChange({ id: "clean-edit", type: "replace_text", target: snapshot.selection.target, description: "Clean edit", before: "old", after: "new", status: "approved", createdAt: new Date().toISOString() });
      expect(result.success).toBe(true);
      expect(trackingAtWrite).toEqual(["Off"]);
      expect(trackingMode).toBe("TrackAll");
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it("recovers a paragraph target when Word shifts its body index", async () => {
    const calls: string[][] = [];
    const first = { text: "a different paragraph", load: () => undefined, insertText: (text: string, mode: string) => { calls.push([text, mode]); } };
    const targetParagraph = { text: "target paragraph", load: () => undefined, insertText: (text: string, mode: string) => { calls.push([text, mode]); } };
    const selection = { text: "", load: () => undefined, insertText: () => undefined };
    const body = { text: "a different paragraph\ntarget paragraph", load: () => undefined, paragraphs: { items: [first, targetParagraph], load: () => undefined } };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const result = await adapter.applyChange({
        id: "office-paragraph-rebase",
        type: "replace_text",
        target: { kind: "paragraph", id: "paragraph-0-stale", paragraphIndex: 0, beforeText: "target paragraph", beforeFingerprint: textFingerprint("target paragraph") },
        description: "Update paragraph",
        before: "target paragraph",
        after: "updated paragraph",
        status: "approved",
        createdAt: new Date().toISOString(),
      });
      expect(result.success).toBe(true);
      expect(calls).toEqual([["updated paragraph", "Replace"]]);
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it("does not reject Hebrew paragraphs because Word adds bidi or paragraph-mark characters", async () => {
    const calls: string[][] = [];
    const paragraph = { text: "שלום עולם\u200f\r", load: () => undefined, insertText: (text: string, mode: string) => { calls.push([text, mode]); } };
    const selection = { text: "", load: () => undefined, insertText: () => undefined };
    const body = { text: "שלום עולם\u200f\r", load: () => undefined, paragraphs: { items: [paragraph], load: () => undefined } };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const result = await adapter.applyChange({
        id: "hebrew-paragraph",
        type: "replace_text",
        target: { kind: "paragraph", id: "paragraph-0-stable", paragraphIndex: 0, beforeText: "שלום עולם", beforeFingerprint: textFingerprint("שלום עולם") },
        description: "Rewrite Hebrew paragraph",
        before: "שלום עולם",
        after: "שלום עולם חדש",
        status: "approved",
        createdAt: new Date().toISOString(),
      });
      expect(result.success).toBe(true);
      expect(calls).toEqual([["שלום עולם חדש", "Replace"]]);
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it("applies whole-document replacements across Word paragraph serialization", async () => {
    const calls: string[][] = [];
    const selection = { text: "", load: () => undefined, insertText: () => undefined };
    const body = {
      text: "שלום עולם\u200f\r",
      load: () => undefined,
      insertText: (text: string, mode: string) => { calls.push([text, mode]); },
      paragraphs: { items: [], load: () => undefined },
    };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const result = await adapter.applyChange({
        id: "office-document-normalization",
        type: "replace_text",
        target: { kind: "document", id: "document", beforeText: "שלום עולם", beforeFingerprint: textFingerprint("שלום עולם") },
        description: "Replace the document",
        before: "שלום עולם",
        after: "מכתב חדש",
        status: "approved",
        createdAt: new Date().toISOString(),
      });
      expect(result.success).toBe(true);
      expect(calls).toEqual([["מכתב חדש", "Replace"]]);
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it("applies an insertTable operation through the Office adapter", async () => {
    const calls: unknown[][] = [];
    const selection = {
      text: "",
      load: () => undefined,
      insertText: () => undefined,
      insertTable: (...args: unknown[]) => { calls.push(args); },
    };
    const body = { text: "", load: () => undefined, paragraphs: { items: [], load: () => undefined } };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const snapshot = await adapter.readSnapshot();
      const result = await adapter.applyChange({
        id: "office-table",
        type: "range_operation",
        target: snapshot.selection.target,
        description: "Insert table",
        before: "",
        operation: { name: "insertTable", args: [2, 2, "After", [["1", "2"], ["3", "4"]]] },
        status: "approved",
        createdAt: new Date().toISOString(),
      });
      expect(result.success).toBe(true);
      expect(calls).toEqual([[2, 2, "After", [["1", "2"], ["3", "4"]]]]);
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as unknown as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as unknown as { Word?: unknown }).Word = previousWord;
    }
  });

  it("detects an empty-text table selection and deletes the selected table", async () => {
    let deleted = false;
    const table = { isNullObject: false, load: () => undefined, delete: () => { deleted = true; } };
    const selection = {
      text: "",
      load: () => undefined,
      tables: { items: [table], load: () => undefined },
      insertText: () => undefined,
    };
    const body = { text: "", load: () => undefined, paragraphs: { items: [], load: () => undefined } };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const snapshot = await adapter.readSnapshot();
      expect(snapshot.selection.isTable).toBe(true);
      expect(snapshot.selection.text).toBe("");
      const result = await adapter.applyChange({
        id: "office-delete-table",
        type: "range_operation",
        target: snapshot.selection.target,
        description: "Delete the selected table",
        before: "",
        operation: { name: "delete", scope: "table" },
        status: "approved",
        createdAt: new Date().toISOString(),
      });
      expect(result.success).toBe(true);
      expect(deleted).toBe(true);
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as unknown as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as unknown as { Word?: unknown }).Word = previousWord;
    }
  });

  it("applies a built-in style to the selected table", async () => {
    let styleBuiltIn = "TableGrid";
    const table = { isNullObject: false, styleBuiltIn, load: () => undefined, delete: () => undefined };
    const selection = {
      text: "",
      load: () => undefined,
      tables: { items: [table], load: () => undefined },
      insertText: () => undefined,
    };
    const body = { text: "", load: () => undefined, paragraphs: { items: [], load: () => undefined } };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const snapshot = await adapter.readSnapshot();
      const result = await adapter.applyChange({
        id: "office-style-table",
        type: "range_operation",
        target: snapshot.selection.target,
        description: "Style the selected table",
        before: "",
        operation: { name: "styleBuiltIn", scope: "table", value: "Grid Table 4 – Accent 1" },
        status: "approved",
        createdAt: new Date().toISOString(),
      });
      expect(result.success).toBe(true);
      expect(table.styleBuiltIn).toBe("GridTable4_Accent1");
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it("fills every selected table cell directly through Word.Table cells", async () => {
    const values: string[][] = Array.from({ length: 3 }, () => Array.from({ length: 4 }, () => ""));
    const table = {
      isNullObject: false,
      load: () => undefined,
      rows: {
        items: values.map(row => ({
          cells: {
            items: row.map((_value, columnIndex) => ({
              body: { insertText: (text: string) => { row[columnIndex] = text; } },
            })),
            load: () => undefined,
          },
        })),
        load: () => undefined,
      },
    };
    const selection = {
      text: "",
      load: () => undefined,
      tables: { items: [table], load: () => undefined },
      insertText: () => undefined,
    };
    const body = { text: "", load: () => undefined, paragraphs: { items: [], load: () => undefined } };
    const context = { document: { getSelection: () => selection, body }, sync: async () => undefined };
    const previousOffice = (globalThis as unknown as { Office?: unknown }).Office;
    const previousWord = (globalThis as unknown as { Word?: unknown }).Word;
    (globalThis as unknown as { Office?: unknown }).Office = { host: "Word", platform: "Windows", context: { requirements: { isSetSupported: () => false } } };
    (globalThis as unknown as { Word?: unknown }).Word = { run: async (callback: (value: typeof context) => Promise<unknown>) => callback(context) };
    try {
      const adapter = new OfficeWordAdapter();
      const snapshot = await adapter.readSnapshot();
      const result = await adapter.applyChange({
        id: "office-fill-table",
        type: "range_operation",
        target: snapshot.selection.target,
        description: "Fill selected table",
        before: "",
        operation: { name: "fillTable", scope: "table", args: [{ mode: "sequence", start: 1, step: 1 }] },
        status: "approved",
        createdAt: new Date().toISOString(),
      });
      expect(result.success).toBe(true);
      expect(values).toEqual([
        ["1", "2", "3", "4"],
        ["5", "6", "7", "8"],
        ["9", "10", "11", "12"],
      ]);
    } finally {
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });
});
