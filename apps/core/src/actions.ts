import type { AgentAction, AgentActionStatus } from "../../../packages/shared/src/index.js";

export class AgentActionStore {
  private readonly values = new Map<string, AgentAction>();
  private readonly maxEntries = 100;

  add(action: AgentAction): AgentAction {
    this.values.set(action.id, action);
    while (this.values.size > this.maxEntries) {
      const oldest = this.values.keys().next().value as string | undefined;
      if (!oldest) break;
      this.values.delete(oldest);
    }
    return action;
  }

  get(id: string): AgentAction | null { return this.values.get(id) ?? null; }

  list(): AgentAction[] { return [...this.values.values()].reverse(); }

  transition(id: string, status: AgentActionStatus, patch: Partial<AgentAction> = {}): AgentAction {
    const current = this.values.get(id);
    if (!current) throw new Error("action not found");
    const next = { ...current, ...patch, status };
    this.values.set(id, next);
    return next;
  }

  reject(id: string): AgentAction {
    const current = this.requirePending(id);
    return this.transition(current.id, "rejected");
  }

  approve(id: string): AgentAction {
    const current = this.requirePending(id);
    return this.transition(current.id, "approved");
  }

  private requirePending(id: string): AgentAction {
    const action = this.values.get(id);
    if (!action) throw new Error("action not found");
    if (action.status !== "pending") throw new Error(`action is already ${action.status}`);
    return action;
  }
}
