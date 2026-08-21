import { describe, expect, it } from "vitest";
import { ChangeStore, StaleChangeError } from "../apps/core/src/changes.js";
import { textFingerprint, type ProposedChange } from "../packages/shared/src/index.js";

function change(): ProposedChange {
  const before = "Original wording";
  return { id: "change-1", type: "replace_text", target: { kind: "selection", id: "selection", beforeText: before, beforeFingerprint: textFingerprint(before) }, description: "Improve wording", before, after: "Improved wording", status: "pending", createdAt: new Date().toISOString() };
}

describe("change approval", () => {
  it("rejects stale content before approval", () => {
    const store = new ChangeStore();
    store.add(change());
    expect(() => store.approve("change-1", "Newer wording")).toThrow(StaleChangeError);
    expect(store.get("change-1")?.status).toBe("pending");
  });

  it("requires approval before completion", () => {
    const store = new ChangeStore();
    store.add(change());
    expect(() => store.complete("change-1", true)).toThrow(/approved/);
    expect(store.approve("change-1", "Original wording").status).toBe("approved");
    expect(store.complete("change-1", true).status).toBe("applied");
  });

  it("accepts Word-only bidi and paragraph-mark normalization for paragraph targets", () => {
    const before = "שלום עולם";
    const store = new ChangeStore();
    store.add({ ...change(), id: "hebrew-change", target: { kind: "paragraph", id: "paragraph-0", paragraphIndex: 0, beforeText: before, beforeFingerprint: textFingerprint(before) }, before });
    expect(store.approve("hebrew-change", `${before}\u200f\r`).status).toBe("approved");
  });

  it("accepts serialization-only normalization for selection targets", () => {
    const before = "Original wording";
    const store = new ChangeStore();
    store.add({ ...change(), id: "selection-normalized", before, target: { kind: "selection", id: "selection", beforeText: before, beforeFingerprint: textFingerprint(before) } });
    expect(store.approve("selection-normalized", `${before}\u200f\r`).status).toBe("approved");
  });
});
