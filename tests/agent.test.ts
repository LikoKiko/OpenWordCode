import { describe, expect, it } from "vitest";
import { loadDocxSkill, runAgent } from "../packages/agent-core/src/index.js";
import { textFingerprint, type DocumentSnapshot } from "../packages/shared/src/index.js";
import type { ProviderChatRequest, ProviderRuntime, ProviderStreamEvent } from "../packages/providers/src/index.js";

const selection = "The customers needs to submit the form.";
const document: DocumentSnapshot = {
  documentId: "test-document",
  selection: { text: selection, isEmpty: false, target: { kind: "selection", id: "selection", beforeText: selection, beforeFingerprint: textFingerprint(selection) } },
  documentText: selection,
  paragraphs: [{ id: "paragraph-0", index: 0, text: selection }],
  outline: [],
  capabilities: { canRead: true, canWrite: true, canComment: false, canFormat: false },
};

class ScriptedProvider implements ProviderRuntime {
  readonly config = { id: "test", displayName: "Test", kind: "demo" as const, baseUrl: "demo://test", enabled: true, local: true, auth: { method: "none" as const }, privacyNote: "test" };
  private turn = 0;
  async listModels() { return []; }
  async *streamChat(): AsyncGenerator<ProviderStreamEvent> {
    this.turn += 1;
    if (this.turn === 1) {
      yield { type: "tool_call", call: { id: "tool-1", name: "get_selection", arguments: "{}" } };
    } else {
      yield { type: "text", delta: JSON.stringify({ answer: "I prepared a grammar fix.", proposedChanges: [{ type: "replace_text", targetId: "selection", before: selection, after: "The customer needs to submit the form.", description: "Fix subject-verb agreement." }] }) };
    }
    yield { type: "done" };
  }
}

class RangeScriptedProvider implements ProviderRuntime {
  readonly config = { id: "range-test", displayName: "Range Test", kind: "demo" as const, baseUrl: "demo://range-test", enabled: true, local: true, auth: { method: "none" as const }, privacyNote: "test" };
  async listModels() { return []; }
  async *streamChat(): AsyncGenerator<ProviderStreamEvent> {
    yield { type: "text", delta: JSON.stringify({ answer: "I prepared a table insertion.", proposedChanges: [{ type: "range_operation", targetId: "selection", before: selection, operation: { name: "insertTable", args: [2, 2, "After", [["1", "2"], ["3", "4"]]] }, description: "Insert the requested table." }] }) };
    yield { type: "done" };
  }
}

class RepairProvider implements ProviderRuntime {
  readonly config = { id: "repair-test", displayName: "Repair Test", kind: "demo" as const, baseUrl: "demo://repair-test", enabled: true, local: true, auth: { method: "none" as const }, privacyNote: "test" };
  private turn = 0;
  async listModels() { return []; }
  async *streamChat(): AsyncGenerator<ProviderStreamEvent> {
    this.turn += 1;
    if (this.turn === 1) yield { type: "text", delta: "I will rewrite that for you." };
    else yield { type: "text", delta: JSON.stringify({ answer: "I prepared the rewrite.", proposedChanges: [{ type: "replace_text", targetId: "selection", before: selection, after: "The revised wording is clearer.", description: "Rewrite the selected text." }] }) };
    yield { type: "done" };
  }
}

class VisualProvider implements ProviderRuntime {
  readonly config = { id: "visual-test", displayName: "Visual Test", kind: "demo" as const, baseUrl: "demo://visual-test", enabled: true, local: true, auth: { method: "none" as const }, privacyNote: "test" };
  lastRequest?: ProviderChatRequest;
  async listModels() { return []; }
  async *streamChat(request: ProviderChatRequest): AsyncGenerator<ProviderStreamEvent> {
    this.lastRequest = request;
    yield { type: "text", delta: JSON.stringify({ answer: "I cropped the embedded picture.", proposedChanges: [{ type: "range_operation", targetId: "inline-picture-0", operation: { name: "cropImage", scope: "image", args: [{ left: 10, right: 10, unit: "percent" }] }, description: "Crop the document picture." }] }) };
    yield { type: "done" };
  }
}

class WholeDocumentProvider implements ProviderRuntime {
  readonly config = { id: "whole-document-test", displayName: "Whole Document Test", kind: "demo" as const, baseUrl: "demo://whole-document-test", enabled: true, local: true, auth: { method: "none" as const }, privacyNote: "test" };
  async listModels() { return []; }
  async *streamChat(): AsyncGenerator<ProviderStreamEvent> {
    yield { type: "text", delta: JSON.stringify({ answer: "I prepared the letter.", proposedChanges: [{ type: "replace_text", targetId: "paragraph-0", before: selection, after: "Dear teacher,\n\nI will not be able to attend school tomorrow.\n\nSincerely,\nTom", description: "Replace the document with the requested letter." }] }) };
    yield { type: "done" };
  }
}

describe("agent runtime", () => {
  it("loads the project DOCX skill automatically", () => {
    expect(loadDocxSkill()).toContain("# DOCX creation, editing, and analysis");
  });

  it("executes bounded read tools and returns an approval proposal", async () => {
    const result = await runAgent({ provider: new ScriptedProvider(), modelId: "test", instruction: "Fix grammar", mode: "manual", document });
    expect(result.answer).toContain("grammar fix");
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.status).toBe("pending");
    expect(result.changes[0]?.after).toBe("The customer needs to submit the form.");
  });

  it("keeps manual approval changes pending", async () => {
    const result = await runAgent({ provider: new ScriptedProvider(), modelId: "test", instruction: "Fix grammar", mode: "manual", document });
    expect(result.changes[0]?.status).toBe("pending");
  });

  it("repairs a provider response that forgot to return a Word change", async () => {
    const result = await runAgent({ provider: new RepairProvider(), modelId: "test", instruction: "Rewrite the selected text", mode: "manual", document });
    expect(result.answer).toContain("prepared the rewrite");
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.after).toBe("The revised wording is clearer.");
  });

  it("accepts a safe Word.Range operation proposal", async () => {
    const result = await runAgent({ provider: new RangeScriptedProvider(), modelId: "test", instruction: "Add a table", mode: "manual", document });
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.type).toBe("range_operation");
    expect(result.changes[0]?.operation?.name).toBe("insertTable");
  });

  it("automatically sends embedded Word pictures to vision and targets them as document elements", async () => {
    const provider = new VisualProvider();
    const png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=";
    const visualDocument: DocumentSnapshot = {
      ...document,
      selection: { ...document.selection, selectedVisualElementIds: ["inline-picture-0"] },
      visualElements: [{ id: "inline-picture-0", kind: "inlinePicture", index: 0, width: 120, height: 80, mimeType: "image/png", dataUrl: `data:image/png;base64,${png}`, size: 68, contentAvailable: true }],
    };
    const result = await runAgent({ provider, modelId: "vision", instruction: "Crop the selected picture", mode: "manual", document: visualDocument });
    expect(provider.lastRequest?.attachments?.[0]).toMatchObject({ id: "document-inline-picture-0", mimeType: "image/png" });
    expect(provider.lastRequest?.messages.at(-1)?.content).toContain("inline-picture-0");
    expect(result.changes[0]?.target).toMatchObject({ kind: "visual", visualKind: "inlinePicture", visualIndex: 0 });
    expect(result.changes[0]?.operation).toMatchObject({ name: "cropImage", scope: "image" });

    const directCrop = await runAgent({ provider, modelId: "vision", instruction: "Crop it to half", mode: "manual", document: visualDocument });
    expect(directCrop.changes[0]?.target.kind).toBe("visual");
    expect(directCrop.changes[0]?.operation).toMatchObject({ name: "cropImage", scope: "image" });

    const directBackgroundRemoval = await runAgent({ provider, modelId: "vision", instruction: "Remove its background", mode: "manual", document: visualDocument });
    expect(directBackgroundRemoval.changes[0]?.operation).toMatchObject({ name: "removeBackground", scope: "image" });
  });

  it("turns a direct table request into an actionable insertion", async () => {
    const result = await runAgent({ provider: new ScriptedProvider(), modelId: "test", instruction: "Create a table with numbers 1 2 3 4 1234 1234 1234", mode: "manual", document });
    expect(result.answer).toContain("3×4");
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.operation?.name).toBe("insertTable");
    expect(result.changes[0]?.operation?.args?.[3]).toEqual([
      ["1", "2", "3", "4"],
      ["1", "2", "3", "4"],
      ["1", "2", "3", "4"],
    ]);
  });

  it("turns an empty-text table selection into a table deletion proposal", async () => {
    const tableDocument: DocumentSnapshot = {
      ...document,
      selection: { text: "", isEmpty: true, isTable: true, tableCount: 1, target: { kind: "selection", id: "selection", beforeText: "", beforeFingerprint: textFingerprint("") } },
    };
    const result = await runAgent({ provider: new ScriptedProvider(), modelId: "test", instruction: "Delete the selected table", mode: "manual", document: tableDocument });
    expect(result.answer).toContain("selected table");
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.operation).toEqual({ name: "delete", scope: "table" });
  });

  it("turns a selected-table labeling request into a direct cell-fill operation", async () => {
    const tableDocument: DocumentSnapshot = {
      ...document,
      selection: { text: "", isEmpty: true, isTable: true, tableCount: 1, target: { kind: "selection", id: "selection", beforeText: "", beforeFingerprint: textFingerprint("") } },
    };
    const result = await runAgent({ provider: new ScriptedProvider(), modelId: "test", instruction: "add label numbers inside the cells", mode: "manual", document: tableDocument });
    expect(result.answer).toContain("sequential labels");
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.operation).toEqual({ name: "fillTable", scope: "table", args: [{ mode: "sequence", start: 1, step: 1 }] });
  });

  it("targets the whole document for a clear-all request", async () => {
    const documentWithText: DocumentSnapshot = { ...document, documentText: "First paragraph\nSecond paragraph" };
    const result = await runAgent({ provider: new ScriptedProvider(), modelId: "test", instruction: "Delete all text from the document", mode: "manual", document: documentWithText });
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.target.kind).toBe("document");
    expect(result.changes[0]?.operation?.name).toBe("clear");
  });

  it("retargets a model paragraph edit when the user asks to replace the whole document", async () => {
    const documentWithText: DocumentSnapshot = { ...document, documentText: selection };
    const result = await runAgent({ provider: new WholeDocumentProvider(), modelId: "test", instruction: "Replace all existing content in the document with this letter", mode: "manual", document: documentWithText });
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]?.target.kind).toBe("document");
    expect(result.changes[0]?.before).toBe(selection);
    expect(result.changes[0]?.after).toContain("Dear teacher");

    const hebrewResult = await runAgent({ provider: new WholeDocumentProvider(), modelId: "test", instruction: "תמחק את הכל ותכניס מכתב חדש במקום התוכן הקיים", mode: "manual", document: documentWithText });
    expect(hebrewResult.changes[0]?.target.kind).toBe("document");
  });
});
