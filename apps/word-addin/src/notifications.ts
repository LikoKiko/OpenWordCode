import { Notyf, NotyfEvent, type NotyfNotification } from "notyf";

/**
 * Toast notifications for the task pane.
 *
 * The pane polls sign-in and Core status on a timer, so the same failure can be
 * reported many times in a row. The old inline banner absorbed that by simply
 * overwriting its text; toasts would stack instead. Identical messages are
 * therefore deduplicated while their toast is still on screen.
 */

type ToastKind = "error" | "success";

const ERROR_DURATION = 6_000;
const SUCCESS_DURATION = 3_000;

let instance: Notyf | null = null;
const active = new Map<string, NotyfNotification>();

function notyf(): Notyf {
  if (instance) return instance;
  instance = new Notyf({
    duration: ERROR_DURATION,
    ripple: false,
    dismissible: true,
    position: { x: "center", y: "top" },
    types: [
      { type: "error", className: "owc-toast owc-toast-error", duration: ERROR_DURATION, icon: false },
      { type: "success", className: "owc-toast owc-toast-success", duration: SUCCESS_DURATION, icon: false },
    ],
  });
  return instance;
}

function show(kind: ToastKind, message: string): void {
  const text = message.trim();
  if (!text || typeof window === "undefined") return;
  const key = `${kind}:${text}`;
  if (active.has(key)) return;
  const toast = notyf().open({ type: kind, message: text });
  active.set(key, toast);
  const forget = (): void => { active.delete(key); };
  toast.on(NotyfEvent.Dismiss, forget);
  window.setTimeout(forget, kind === "error" ? ERROR_DURATION : SUCCESS_DURATION);
}

export function notifyError(message: string): void {
  show("error", message);
}

export function notifySuccess(message: string): void {
  show("success", message);
}

export function dismissNotifications(): void {
  if (!instance) return;
  instance.dismissAll();
  active.clear();
}
