import { installWorkbookPlanAcceptanceHarness, type WorkbookPlanAcceptanceWorkbook } from "./workbookPlanAcceptanceHarness";

declare const __HERMES_OFFICE_ACCEPTANCE__: boolean;

/** Dedicated acceptance-build entry, never imported by taskpane.ts.
 * The later native launcher supplies the disposable workbook authority only
 * after its separate live-session permission has been granted. */
export function attachWorkbookPlanAcceptance(
  workbook: WorkbookPlanAcceptanceWorkbook,
  target: Record<string, unknown> = globalThis as unknown as Record<string, unknown>,
  location: { href: string } = globalThis.location,
): boolean {
  if (typeof __HERMES_OFFICE_ACCEPTANCE__ === "undefined" || !__HERMES_OFFICE_ACCEPTANCE__) return false;
  return installWorkbookPlanAcceptanceHarness(target, workbook, location);
}

if (typeof __HERMES_OFFICE_ACCEPTANCE__ !== "undefined" && __HERMES_OFFICE_ACCEPTANCE__) {
  (globalThis as unknown as Record<string, unknown>).__hermesExcelAttachWorkbookPlanAcceptance = attachWorkbookPlanAcceptance;
}
