# Full-functionality Excel acceptance harness

Current product scope: the limited desktop Excel for Mac bridge/table/formula flow is recorded in
[the README](../README.md). Windows desktop Excel is untested. That smoke test does not close the broader
WorkbookPlan/native acceptance gates below; historical slice results are not claims about full product acceptance.

This is the acceptance-enabled B2/B3 fixture lane. It is not a replacement for the native 27-tool matrix, the Hermes bridge lane, or model/workflow acceptance.

## Guard and launch contract

The harness is compiled only when `HERMES_OFFICE_ACCEPTANCE=1` is set while webpack is invoked. Runtime installation then requires the exact HTTPS task-pane location:

```text
https://localhost:3000/taskpane.html?fullFunctionalityAcceptance=1
```

The URL may contain exactly one nonempty `_host_Info` parameter supplied by Office. It may not contain a fragment, unknown parameter, duplicate parameter, empty metadata value, or the focused harness query `?acceptance=1`. The normal production build fails closed and does not install this global.

Building and serving the acceptance artifact needs a `dev-server` npm script run with `NODE_ENV=development`, the Hermes
Python as `PYTHON` and `HERMES_OFFICE_ACCEPTANCE=1`. This package currently defines no `dev-server` script
(see `package.json`), and the controller below refuses to start without one, so this lane cannot be started from
this checkout until that script is restored.

For the server/readiness controller, use:

```bash
node scripts/run-full-excel-acceptance.mjs
```

The controller starts the repository's `webpack-cli serve --mode development` directly with the same environment as that `dev-server` script, waits for the exact HTTPS URL and acceptance-only dynamic chunk markers, then prints the URL and:

```text
await window.__hermesExcelFullFunctionalityAcceptance.run()
```

It stays alive for the task-pane session and terminates only its directly owned webpack child on signals, escalating from `SIGTERM` to `SIGKILL` after a bounded timeout if needed. A generic HTTP 200 is insufficient: readiness requires the generated full-functionality chunk to contain the bounded harness markers. It does not start, stop, probe, or own the Hermes API, proxy, companion, model, or any other service.

Open the URL in a real Excel desktop task pane from a newly created disposable workbook. Do not run this lane against a user workbook.

## B2/B3 scope

`runFullFunctionalityAcceptance()` requires the real global Excel host and has no injectable Office runtime. It creates, seeds, reads back, and finally deletes only its own seven sheets. It does not invoke `executeValidatedExcelBridgeTool`, a model, a gateway, a companion, or a proxy. Therefore its successful fixture setup/readback is preparation evidence only; it is not a native tool result and does not claim Phase C or later acceptance.

Every run creates a secure lowercase eight-character run ID and these seven unique sheet names. The final suffix is the plan role:

| Role | Sheet suffix | Fixture contents |
|---|---:|---|
| Data | `_D` | Fixed values, dates, blanks, formulas, deterministic error formulas, and data sentinels |
| Formats | `_F` | Header/title band, long text, dates, currency, percentages, and format sentinels |
| Tables | `_T` | A deterministic named 20-row table, a separate insert target, duplicate values, three regions, and table sentinels |
| Pivot | `_P` | Empty pivot destination and pivot sentinels |
| Comments | `_C` | Disposable comment/update/clear target cells and comment sentinels |
| Protected | `_R` | Passwordless protected target and protection sentinels |
| Expected failures | `_X` | Passwordless-protected pre-merged span, occupied cell, invalid-target area, and failure sentinels |

Names have the form `Hermes_Acceptance_<runId>_<suffix>` and are at most 31 characters. A preflight collision check fails before fixture creation. The prior active user sheet is restored when possible. Cleanup uses the exact names recorded as created by the current run, never a prefix wildcard and never a user-selected name. Cleanup is in `finally`, including setup/readback failure paths.

Fixture values and sentinels are fixed synthetic values. Workbook values, formulas, errors, comments, passwords, prompts, and raw Office error text are never returned as evidence. No password argument is present in the acceptance registry.

## Deterministic 27-tool registry

`FULL_FUNCTIONALITY_ACCEPTANCE_CASES` is the reusable registry for the canonical 27 tools from `src/excelTools.ts`. Case IDs are unique and use `tool:positive`, `tool:negative`, or `tool:boundary` (with an optional bounded suffix). Each tool has one positive and one negative/boundary slot. Every definition requires a fresh baseline fixture and provides explicit prerequisite metadata for temporary sheets or comments, so Phase C does not depend on hidden case order. Sheet and table arguments are derived from the generated run sheet set; no bridge target is exposed. Pivot and comment cases explicitly allow a bounded unsupported-host result without treating it as a successful native operation.

| Tool | Positive slot | Negative/boundary slot |
|---|---|---|
| `get_workbook_overview` | populated disposable workbook overview | boundary overview with failure fixture present |
| `read_range` | fixed data range values/formulas | malformed A1 address |
| `summarize_range_stats` | numeric data statistics | invalid/zero-row address |
| `write_range` | rectangular synthetic write | dimension mismatch |
| `format_range` | documented formatting properties | invalid alignment with no mutation |
| `conditional_format_range` | cell-value rule | invalid operator |
| `freeze_panes` | one row and one column | negative row boundary |
| `set_data_validation` | fixed list source | empty list source |
| `create_sheet` | generated temporary name | duplicate generated sheet name |
| `delete_sheet` | generated temporary sheet | missing generated name |
| `activate_sheet` | generated formats sheet | missing generated name |
| `insert_table` | separate four-row source with headers | overlap with the deterministic named table |
| `sort_table` | generated table name and valid column | out-of-range column index |
| `filter_table` | generated table name and fixed values | out-of-range column index |
| `create_chart` | clean two-column source | invalid chart type |
| `find_replace` | fixed case-sensitive replacement | empty find boundary |
| `clear_range` | exact disposable range | protected target |
| `autofill_formula` | one formula source through a blank target range | multi-cell source |
| `rename_sheet` | generated failure sheet to generated temporary name | duplicate generated destination |
| `copy_sheet` | generated data sheet to generated temporary name | same source/destination |
| `protect_sheet` | passwordless protection | already-protected target |
| `unprotect_sheet` | passwordless protected target | already-unprotected success boundary |
| `create_pivot_table` | generated table source and empty pivot destination | missing-field rejection branch |
| `run_calculation` | fixed formula/error fixture calculation | error-formula boundary |
| `add_comment` | new-comment and explicitly seeded update targets | multi-cell target |
| `clear_comments` | explicitly seeded target with an out-of-range comment sentinel | invalid address |
| `import_csv_to_sheet` | fixed small RFC-style CSV | malformed quoted CSV |

The registry is a definition layer, not an execution claim. Mutating definitions require sentinel evidence. The pure evaluator requires captured, expectation-matching pre-state, tool-result evidence, post-state, and sentinel evidence. It rejects missing or wrong evidence and records `noPartialMutation`; a tool is never marked passed from its return status alone.

## Evidence contract

The returned object is schema-versioned and content-free. It contains only:

- run ID;
- bounded case metadata and status flags;
- pass/fail counts;
- exact-sheet cleanup counts/status;
- pending future-native slots.

`serializeFullFunctionalityAcceptanceEvidence()` projects an explicit allow-list, drops injected fields such as `values`, `formulas`, `rawError`, `arguments`, and `csv`, and refuses output over 64 KiB. Setup/readback can populate only the precondition slot. Tool result, post-state, and post-mutation sentinel slots remain pending because B2/B3 does not execute Phase C.

Formula, workflow, isolation, and recovery slots are always present as:

```text
status: pending
nativeRequired: true
resultCount: 0
```

They must never be reported as passed by this harness. Phase C direct native tool execution, Phase D Hermes-native bridge execution, Phase E formula/workflow acceptance, and Phase F cancellation/recovery acceptance remain pending until their respective live lanes are run in disposable workbooks.

Static tests, unit tests, successful builds, fixture preparation, HTTP readiness, and this documentation do not prove native Excel acceptance. This B2/B3 harness does not prove native Excel acceptance. Native acceptance requires real Excel Office.js execution and readback evidence from the disposable workbook.
