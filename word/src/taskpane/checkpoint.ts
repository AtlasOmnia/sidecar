/**
 * Wait for an accessible checkpoint decision while preserving the caller's
 * abort reason and removing every listener when the decision settles.
 */
export function waitForCheckpointDecision(element: HTMLElement, signal: AbortSignal): Promise<boolean> {
  const continueButton = element.querySelector<HTMLButtonElement>(".checkpoint-continue");
  const stopButton = element.querySelector<HTMLButtonElement>(".checkpoint-stop");
  if (!continueButton || !stopButton) {
    return Promise.reject(new Error("Checkpoint controls are missing."));
  }
  if (signal.aborted) {
    return Promise.reject(signal.reason ?? new DOMException("The checkpoint was aborted.", "AbortError"));
  }

  return new Promise<boolean>((resolve, reject) => {
    let settled = false;
    const cleanup = (): void => {
      continueButton.removeEventListener("click", onContinue);
      stopButton.removeEventListener("click", onStop);
      signal.removeEventListener("abort", onAbort);
    };
    const settle = (callback: () => void): void => {
      if (settled) return;
      settled = true;
      cleanup();
      callback();
    };
    const onContinue = (): void => settle(() => resolve(true));
    const onStop = (): void => settle(() => resolve(false));
    const onAbort = (): void => settle(() => reject(signal.reason ?? new DOMException("The checkpoint was aborted.", "AbortError")));

    continueButton.addEventListener("click", onContinue);
    stopButton.addEventListener("click", onStop);
    signal.addEventListener("abort", onAbort, { once: true });
  });
}
