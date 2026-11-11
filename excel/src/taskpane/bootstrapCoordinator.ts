declare const __HERMES_OFFICE_ACCEPTANCE__: boolean;

/**
 * Narrow typed dependencies for the bootstrap decision coordinator.
 *
 * All callback names are deliberately generic so the coordinator module can
 * be bundled into both normal and acceptance production builds without
 * leaking harness-specific string literals.  The call site wraps the real
 * dynamic imports.
 */
export interface BootstrapDeps {
  /** Exact URL validator (production-only guard). */
  isFullFunctionalityAcceptanceUrl(): boolean;
  /**
   * Dynamic import + install for the full-functionality acceptance harness.
   * Resolves `true` when the harness installed successfully (normal boot is
   * suppressed).  Resolves `false` to fall back to normal boot.  Rejects on
   * import failure — treated as fall-back.
   */
  loadAndInstallFullHarness(): Promise<boolean>;
  /**
   * Dynamic import + install for the focused office acceptance harness.
   * Rejection is consumed silently; normal boot always runs first.
   */
  loadAndInstallFocusedHarness(): Promise<void>;
  /** Normal task-pane boot entry point. */
  bootTaskpane(): void;
}

/**
 * Pure bootstrap decision coordinator.
 *
 * Preserves the compile guard (`__HERMES_OFFICE_ACCEPTANCE__`), exact URL
 * validator, production dead-code exclusion, and the void-import fire-and-forget
 * pattern.  All side effects flow through the injected dependencies so the
 * coordinator can be tested without DOM, Office.js, or real dynamic imports.
 */
export function bootstrap(deps: BootstrapDeps): void {
  const compiledAcceptanceEnabled =
    typeof __HERMES_OFFICE_ACCEPTANCE__ !== "undefined" &&
    __HERMES_OFFICE_ACCEPTANCE__ === true;

  if (compiledAcceptanceEnabled && deps.isFullFunctionalityAcceptanceUrl()) {
    void deps
      .loadAndInstallFullHarness()
      .then((installed) => {
        if (!installed) {
          deps.bootTaskpane();
        }
      })
      .catch(() => {
        deps.bootTaskpane();
      });
  } else {
    if (compiledAcceptanceEnabled) {
      void deps.loadAndInstallFocusedHarness().catch(() => {
        // Harness install failed; normal boot already runs below.
      });
    }
    deps.bootTaskpane();
  }
}
