export type CheckpointDecision = "continue" | "stop";

interface PendingCheckpoint {
  resolve: (decision: CheckpointDecision) => void;
  onAbort: () => void;
  signal: AbortSignal;
}

/** A one-shot, abort-aware checkpoint prompt coordinator for the task pane. */
export class CheckpointController {
  private pending: PendingCheckpoint | null = null;

  begin(signal: AbortSignal): Promise<CheckpointDecision> {
    if (this.pending) return Promise.resolve("continue");
    return new Promise<CheckpointDecision>((resolve) => {
      const onAbort = () => this.finish("stop");
      this.pending = { resolve, onAbort, signal };
      if (signal.aborted) onAbort();
      else signal.addEventListener("abort", onAbort, { once: true });
    });
  }

  choose(decision: CheckpointDecision): void {
    this.finish(decision);
  }

  cancel(): void {
    this.finish("stop");
  }

  get isPending(): boolean {
    return this.pending !== null;
  }

  private finish(decision: CheckpointDecision): void {
    const pending = this.pending;
    if (!pending) return;
    this.pending = null;
    pending.signal.removeEventListener("abort", pending.onAbort);
    pending.resolve(decision);
  }
}
