# Hermes Excel Add-in Capability Matrix

Latest live run: 2026-11-11 about 00:26 in desktop Excel for Mac, plugin and sidebar built from
`a7b3fd7`. Connect and Allow succeeded; a 3-row table with header in A1:B4 and `=SUM(B2:B4)` in B5
were read back from the workbook. Revoke dropped the sidebar to “needs your OK”. The bridge answered
200 to the configured origin and 403 to another origin. The sidebar used a private test site via `addin_base`;
the CA was trusted in the login keychain, and this Mac also has it in the System keychain.

Earlier the same night at `35ef09e` (behaviour unchanged since): reversal on request by clearing the range
(no real undo), plugin reload with the sidebar open (it reconnected by itself and edited again), approval
remembered after restarting Excel, and the next Connect after revoke asked for approval again.
Windows desktop Excel is untested. Also not tested: dark theme, keyboard-only use, the public Pages origin,
TLS-only trust, Excel for the web, or full tool/WorkbookPlan coverage. Reopen the sidebar if it does not reconnect.
“Supported now” means implemented, not live-verified on every host. Reversing an edit is another edit,
not restoration of Excel's undo history.
See [desk-bridge.md](desk-bridge.md) for manual CA trust and installer-supplied `addin_base`.

Each feature is tagged with its status:

| Status | Meaning |
|--------|---------|
| **Supported now** | Implemented as an Office.js tool in `excelTools.ts`. |
| **Supported through formulas** | Achievable by writing Excel formulas via `write_range` (e.g. `=SUMIFS`, `=XLOOKUP`). |
| **Planned** | On the roadmap, not yet implemented. |
| **Out of scope** | Not planned — better handled by other tooling. |
| **Requires confirmation** | May work depending on Excel host version. |
| **Requires external service** | Needs a network API call, not pure Office.js. |

---

## Workbook tools

| Feature | Status | Notes |
|---------|--------|-------|
| Read values / formulas / number formatting | **Supported now** | `read_range` |
| Write values and formulas | **Supported now** | `write_range` |
| Get workbook overview | **Supported now** | `get_workbook_overview` |
| Recalculate workbook | **Supported now** | `run_calculation` |
| Find and replace | **Supported now** | `find_replace` |
| Clear range | **Supported now** | `clear_range` |
| Import CSV data | **Supported now** | `import_csv_to_sheet` — RFC 4180 via PapaParse |
| Export to CSV | **Planned** | Read range and format as CSV string |
| Clear sheet | **Supported through formulas** | `clear_range` with sheet's used range address |
| Undo / redo | **Out of scope** | Sidecar supplies no real undo; a compensating edit is not Excel Undo |

## Sheet management

| Feature | Status | Notes |
|---------|--------|-------|
| Create sheet | **Supported now** | `create_sheet` |
| Delete sheet | **Supported now** | `delete_sheet` |
| Activate / switch sheet | **Supported now** | `activate_sheet` |
| Rename sheet | **Supported now** | `rename_sheet` |
| Copy sheet | **Supported now** | `copy_sheet` |
| Protect sheet | **Supported now** | `protect_sheet` — with optional password |
| Unprotect sheet | **Supported now** | `unprotect_sheet` |
| Hide / unhide sheet | **Planned** | Office.js supports `sheet.visibility` |
| Move sheet (reorder) | **Planned** | Office.js `sheet.position` |
| Group / ungroup sheets | **Out of scope** | No practical Office.js API |

## Formatting

| Feature | Status | Notes |
|---------|--------|-------|
| Bold / italic / font color | **Supported now** | `format_range` |
| Fill color | **Supported now** | `format_range` |
| Number format | **Supported now** | `format_range` |
| Borders (all edges) | **Supported now** | `format_range` |
| Horizontal / vertical alignment | **Supported now** | `format_range` — canonical enum: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed |
| CenterAcrossSelection (span without merge) | **Supported now** | `format_range` — single-row multi-cell span, precondition checks enforced |
| Column autofit | **Supported now** | `format_range` — explicit opt-in only (`columnAutofit: true`) |
| Row autofit | **Supported now** | `format_range` — explicit opt-in only (`rowAutofit: true`) |
| Font name / size | **Supported now** | `format_range` |
| Underline style | **Supported now** | `format_range` |
| Wrap text | **Supported now** | `format_range` |
| Indent level | **Supported now** | `format_range` |
| Text orientation | **Supported now** | `format_range` |
| Custom border color / weight | **Supported now** | `format_range` |
|| Merge / unmerge cells | **Out of scope** | No merge/unmerge capability; use CenterAcrossSelection instead |
| Conditional formatting: cell value rules | **Supported now** | `conditional_format_range` — GreaterThan, etc. |
| Conditional formatting: contains text | **Supported now** | `conditional_format_range` |
| Conditional formatting: color scale | **Supported through formulas** | Use conditional formatting with `cellValue` rules |
| Conditional formatting: data bar | **Requires confirmation** | Office.js supports `dataBar` in some versions |
| Conditional formatting: icon sets | **Planned** | Office.js `iconSet` conditional format |
| Freeze panes | **Supported now** | `freeze_panes` — freeze rows/columns |
| Cell styles / named styles | **Planned** | Office.js `range.worksheet.names` |

## Tables

| Feature | Status | Notes |
|---------|--------|-------|
| Insert table from range | **Supported now** | `insert_table` |
| Sort table | **Supported now** | `sort_table` |
| Filter table | **Supported now** | `filter_table` — by values or custom criteria |
| Add / remove table columns | **Planned** | Office.js `table.columns.add()` |
| Table style / banded rows | **Supported now** | Default `TableStyleMedium2` on insert |
| Structured references in formulas | **Supported through formulas** | Write `=SUM(Table1[Amount])` via `write_range` |

## Charts

| Feature | Status | Notes |
|---------|--------|-------|
| Create chart from range | **Supported now** | `create_chart` — column/line/pie/bar |
| Chart title / legend | **Supported now** | Set on creation |
| Modify chart type | **Planned** | Office.js `chart.type` |
| Chart axis labels / formatting | **Planned** | Office.js `chart.axes` |
| Sparklines | **Planned** | Office.js `worksheet.sparklineGroups` |

## Data analysis

| Feature | Status | Notes |
|---------|--------|-------|
| Range statistics (count, sum, avg, min, max) | **Supported now** | `summarize_range_stats` |
| Pivot table | **Supported now** | `create_pivot_table` — ExcelApi 1.8+ |
| Data validation (list dropdown) | **Supported now** | `set_data_validation` |
| Data validation (date, number) | **Planned** | Office.js supports date/number/time/wholeNumber validations |
| Auto-fill formulas | **Supported now** | `autofill_formula` — with R1C1 fallback |
| Optional deep analysis with PandasAI | **Requires external service** | Runs as a separate Hermes-side Python tool; not bundled into the add-in. See [`pandasai-integration.md`](pandasai-integration.md). |
| Named ranges | **Planned** | `workbook.names.add()` |
| Remove duplicates | **Planned** | Office.js `range.removeDuplicates()` |
| Group / outline rows | **Planned** | Office.js `range.worksheet.rowPageBreaks` |

## Comments and notes

| Feature | Status | Notes |
|---------|--------|-------|
| Add comment to cell | **Supported now** | `add_comment` — ExcelApi 1.10+ required, uses getItemByCell for updates |
| Update existing comment | **Supported now** | `add_comment` detects via getItemByCell and updates .content |
| Clear comments in range | **Supported now** | `clear_comments` — filters by isAddressInRange, ExcelApi 1.10+ required |
| Get/read comments | **Planned** | Useful for reading existing annotations |
| Resolve / reopen thread | **Planned** | Office.js `comment.resolved` |
| Legacy cell notes (Excel 2019) | **Out of scope** | Comments API is the modern path |

## Security

| Feature | Status | Notes |
|---------|--------|-------|
| Protect sheet | **Supported now** | `protect_sheet` — optional password, allow sort/filter |
| Unprotect sheet | **Supported now** | `unprotect_sheet` |
| Protect workbook structure | **Planned** | Office.js `workbook.protection.protect()` |
| Read-only / edit restrictions | **Out of scope** | Application-level controls |

## Advanced

| Feature | Status | Notes |
|---------|--------|-------|
| Run arbitrary JavaScript macros | **Out of scope** | Office.js cannot run VBA; use Excel on the web automation |
| External data connections | **Requires external service** | Use the proxy or Excel's built-in Power Query |
| Print / PDF export | **Out of scope** | Use Excel's native File > Print or `Excel.run` print APIs in future |
| Custom ribbon / commands | **Requires external service** | Requires manifest-level customizations beyond taskpane scope |
| Real-time collaboration | **Requires confirmation** | Office.js co-authoring is read-only for add-ins |
| Notification / alert dialogs | **Planned** | Possible via `Office.context.ui.displayDialogAsync` |

## Documented elsewhere

For the current Hermes plugin integration and tool-calling loop, see:

- [`desk-bridge.md`](desk-bridge.md) — the plugin-owned HTTPS loopback bridge; no separate proxy is required
- [`pandasai-integration.md`](pandasai-integration.md) — optional external PandasAI analysis boundary and wrapper
- `src/agent.ts` — the tool-calling loop implementation
- `src/taskpane/taskpane.ts` — the chat UI and settings
