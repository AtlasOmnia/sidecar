/**
 * R4i (R3c D4): content-free task-pane notice for stored reconciliation markers.
 *
 * Pure and DOM-agnostic. It only reads a count through the injected source, only while the pane
 * is idle, at most one read at a time, and renders a bounded plain-text line. It never writes to
 * the store and never sees record contents.
 */

export const RECONCILIATION_NOTICE_MAX_DISPLAY = 999;

export type ReconciliationNoticeView =
  | { readonly kind: "unavailable" }
  | { readonly kind: "none" }
  | { readonly kind: "show"; readonly text: string };

/** Content-free text for a marker count; anything that is not a safe non-negative integer is "unavailable". */
export function reconciliationNoticeView(count: unknown): ReconciliationNoticeView {
  if (typeof count !== "number" || !Number.isSafeInteger(count) || count < 0) return { kind: "unavailable" };
  if (count === 0) return { kind: "none" };
  if (count === 1) return { kind: "show", text: "1 workbook action needs reconciliation" };
  const shown = count > RECONCILIATION_NOTICE_MAX_DISPLAY ? `${RECONCILIATION_NOTICE_MAX_DISPLAY}+` : String(count);
  return { kind: "show", text: `${shown} workbook actions need reconciliation` };
}

export interface ReconciliationCountSource {
  reconciliationRequiredCount(): Promise<number | undefined>;
}

export interface ReconciliationNoticeTarget {
  textContent: string | null;
  hidden: boolean;
}

export interface ReconciliationNotice {
  refresh(): Promise<void>;
}

export function createReconciliationNotice(deps: {
  readonly source: ReconciliationCountSource;
  readonly target: () => ReconciliationNoticeTarget | null;
  readonly isIdle: () => boolean;
}): ReconciliationNotice {
  let pending: Promise<void> | null = null;
  // R4j (R4i lane-B B-A1): set when a refresh joins the pending read; at most one follow-up is queued.
  let followUp = false;

  const readCount = async (): Promise<unknown> => {
    try {
      return await deps.source.reconciliationRequiredCount();
    } catch {
      return undefined;
    }
  };

  const apply = (view: ReconciliationNoticeView): void => {
    if (view.kind === "unavailable") return;
    const element = deps.target();
    if (element === null) return;
    if (view.kind === "none") {
      element.hidden = true;
      element.textContent = "";
    } else {
      element.textContent = view.text;
      element.hidden = false;
    }
  };

  const readOnce = async (): Promise<void> => {
    try {
      apply(reconciliationNoticeView(await readCount()));
    } catch {
      // Rendering is best effort; the notice never fails a caller.
    }
  };

  const idleNow = (): boolean => {
    try {
      return deps.isIdle() === true;
    } catch {
      return false;
    }
  };

  const run = async (): Promise<void> => {
    try {
      await readOnce();
      // One coalesced follow-up read after the joined read settles, so the line ends at the latest
      // count. Still one read at a time, and only while the pane is idle.
      while (followUp) {
        followUp = false;
        if (!idleNow()) break;
        await readOnce();
      }
    } finally {
      followUp = false;
      pending = null;
    }
  };

  return {
    refresh(): Promise<void> {
      if (pending !== null) {
        followUp = true;
        return pending;
      }
      let idle = false;
      try {
        idle = deps.isIdle() === true;
      } catch {
        idle = false;
      }
      if (!idle) return Promise.resolve();
      pending = run();
      return pending;
    },
  };
}
