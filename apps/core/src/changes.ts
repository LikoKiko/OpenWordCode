import { comparableText, textFingerprint, type ProposedChange, type ProposedChangeStatus } from "../../../packages/shared/src/index.js";

export class StaleChangeError extends Error {
  readonly code = "stale_change";
  constructor() { super("This content changed before the edit could be applied. Refresh the document context and try again."); }
}

export class ChangeStore {
  private readonly values = new Map<string, ProposedChange>();
  private readonly maxEntries = 100;

  add(change: ProposedChange): ProposedChange {
    this.values.set(change.id, change);
    while (this.values.size > this.maxEntries) {
      const oldest = this.values.keys().next().value as string | undefined;
      if (!oldest) break;
      this.values.delete(oldest);
    }
    return change;
  }

  get(id: string): ProposedChange | null { return this.values.get(id) ?? null; }

  list(): ProposedChange[] { return [...this.values.values()].reverse(); }

  reject(id: string): ProposedChange {
    return this.transition(id, "rejected");
  }

  approve(id: string, currentBefore: string): ProposedChange {
    const change = this.require(id);
    if (change.status !== "pending") throw new Error(`change is already ${change.status}`);
    const expected = change.before ?? change.target.beforeText;
    const normalizedEquivalent = (change.target.kind === "selection" || change.target.kind === "paragraph" || change.target.kind === "document")
      && comparableText(currentBefore) === comparableText(expected);
    if ((!normalizedEquivalent && currentBefore !== expected) || (!normalizedEquivalent && textFingerprint(currentBefore) !== change.target.beforeFingerprint)) throw new StaleChangeError();
    return this.transition(id, "approved");
  }

  complete(id: string, success: boolean, reason?: string): ProposedChange {
    const change = this.require(id);
    if (change.status !== "approved") throw new Error(`change must be approved before completion (currently ${change.status})`);
    return this.transition(id, success ? "applied" : "failed", reason);
  }

  private transition(id: string, status: ProposedChangeStatus, reason?: string): ProposedChange {
    const current = this.require(id);
    const next = { ...current, status, ...(reason ? { failureReason: reason } : {}) };
    this.values.set(id, next);
    return next;
  }

  private require(id: string): ProposedChange {
    const change = this.values.get(id);
    if (!change) throw new Error("change not found");
    return change;
  }
}
