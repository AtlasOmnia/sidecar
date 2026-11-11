export interface LinkedDeadline {
  signal: AbortSignal;
  cleanup: () => void;
}
function abortReason(signal: AbortSignal, fallbackName = "AbortError", fallbackMessage = "Operation aborted."): unknown {
  return signal.reason ?? new DOMException(fallbackMessage, fallbackName);
}

export function createLinkedDeadline(
  parent: AbortSignal,
  timeoutMs: number,
  timeoutMessage: string,
): LinkedDeadline {
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  let cleaned = false;

  const abortFromParent = (): void => {
    if (!controller.signal.aborted) {
      controller.abort(abortReason(parent));
    }
  };

  if (parent.aborted) {
    abortFromParent();
  } else {
    parent.addEventListener("abort", abortFromParent, { once: true });
  }

  if (!controller.signal.aborted && Number.isFinite(timeoutMs)) {
    timer = setTimeout(() => {
      if (!controller.signal.aborted) {
        controller.abort(new DOMException(timeoutMessage, "TimeoutError"));
      }
    }, Math.max(0, timeoutMs));
  }

  return {
    signal: controller.signal,
    cleanup: () => {
      if (cleaned) return;
      cleaned = true;
      if (timer !== undefined) {
        clearTimeout(timer);
        timer = undefined;
      }
      parent.removeEventListener("abort", abortFromParent);
    },
  };
}

export function raceWithSignal<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
  if (signal.aborted) {
    // The operation has already started. Observe a later rejection even
    // though the abort result wins immediately, so it cannot become unhandled.
    void promise.catch(() => undefined);
    return Promise.reject(abortReason(signal));
  }

  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const onAbort = (): void => {
      if (settled) return;
      settled = true;
      signal.removeEventListener("abort", onAbort);
      reject(abortReason(signal));
    };

    signal.addEventListener("abort", onAbort, { once: true });
    promise.then(
      (value) => {
        if (settled) return;
        settled = true;
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      },
      (error) => {
        if (settled) return;
        settled = true;
        signal.removeEventListener("abort", onAbort);
        reject(error);
      },
    );
  });
}

export function interruptibleSleep(ms: number, signal: AbortSignal): Promise<void> {
  if (signal.aborted) {
    return Promise.reject(abortReason(signal));
  }
  if (ms <= 0) {
    return Promise.resolve();
  }

  return new Promise<void>((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      signal.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    const onAbort = (): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal.removeEventListener("abort", onAbort);
      reject(abortReason(signal));
    };
    signal.addEventListener("abort", onAbort, { once: true });
  });
}
