# WorkbookPlan acceptance harness (offline only)

Current product scope: the limited desktop Excel for Mac bridge/table/formula flow is recorded in
[the README](../README.md). Windows desktop Excel is untested. That smoke test does not close the broader
WorkbookPlan/native acceptance gates below; historical slice results are not claims about full product acceptance.

## Pin and boundaries

Historical harness baseline: `b1b3422915de340c5d4350e435a12c7aaa6b9465`. At this pin, `SERVER_INSPECTION_SUPPORTED` is **true**, server inspect/execute phase tickets share a turn key, the closed not-executed reasons include `already_executed`, and the production rollout defaults remain off. The primitive acceptance harness still forces WorkbookPlan rollout off; the 27-tool full-functionality harness is still the primitive lane. Neither is changed here.

`workbookPlanAcceptanceEntry.ts` is a separate webpack entry added **only** when `HERMES_OFFICE_ACCEPTANCE=1`. The production taskpane entry has no static or dynamic import of either new acceptance module. The guarded entry publishes `__hermesExcelAttachWorkbookPlanAcceptance`; attaching a runtime requires the exclusive HTTPS location `https://localhost:3000/taskpane.html?workbookPlanAcceptance=1` (one optional `_host_Info`). The run API is `__hermesExcelWorkbookPlanAcceptance.run(runId)`. Undefined compile definitions fail closed, unlike the older primitive harness's test fallback.

The native workbook creation/identity/cleanup authority is deliberately injected through `WorkbookPlanAcceptanceWorkbook`. A later approved native launcher must implement it using a **new disposable workbook** containing exactly the seven `R7H_<runId>_<index>` fixture sheets; it must fail if the active workbook changes or cannot prove host-owned identity. Import/attach does not create a workbook. Each run checks identity before native reads and every adapter mutation handoff, and cleans up in `finally`, including partial creation failures. It does not edit settings, install plugins, use browser persistence, contact providers or start Excel. Synthetic local rollout dependencies are enabled only after the compile guard; the rollout-off matrix row passes false.

## Matrix and oracles

Both direct `createWorkbookPlanTurn` and the actual `createServerOwnedPlanRoute` with `TaskPaneTurnCoordinator` execute:

| Case | Required evidence |
|---|---|
| execute | Real validator/policy/compiler/executor/native adapter and verified native readback |
| inspect-execute | One native read, then one execute; server phase tickets share turn key |
| ask / refuse | Closed outcome, no execute-start acknowledgement |
| invalid | Invalid protocol cannot mutate; direct correction capped at one |
| replay | Seed once, then stored execution replay; server emits `already_executed`, no second acknowledgement |
| Stop mid-turn | Stop inside the native inspection, before model continuation/mutation; cancelled, no rollback claim |
| rotation | Unclean request forces coordinator session rotation; old inspection generation loses authority, fresh turn commits |
| caps | 31-operation schema cap rejection plus schema-valid 5,001-cell read and formula-write resource budget rejection |
| rollout-off | Direct rejected / server not_started; no mutation or acknowledgement |

Server replay advances only the injected ticket-binding clock by 90,001 ms, expiring the old ticket tombstone while retaining the browser execution receipt. This is **not** a real wait, persisted browser reload, live companion or live transport test. Synthetic trusted ticket envelopes and execute-start controls are injected into the real server route. The existing wire/bridge/restart gates independently cover those seams.

The offline oracle is a local strict stateful `Excel.run` subset: load/sync-gated reads, sync-committed writes, exact synthetic sheets, and persistent cell state. It runs the actual native adapter and plan-aware precondition capture; no fake adapter claims success. All twenty route/case receipts must pass; the full run must perform exactly eight writes and zero premature host reads. Identity refusal must make zero writes and still clean up.

Receipts contain only route, case ID, closed outcome/reason, execution outcome, verification status, inspection/correction counters and pass. They exclude summaries, cell values, formula text, sheet names, document/session IDs and inspection bodies.

## What remains pending

No standalone webpack build and no live Excel run occurred for this slice. Static entry-graph exclusion and config-factory tests are offline proof, **not** compiled bundle absence proof. At the later frozen-candidate build, separately record:

1. Production `dist` manifest/hash (fresh export).
2. Acceptance `dist` manifest/hash (different fresh export).
3. Hash of the `workbookPlanAcceptance` entry chunk(s), separately from production.
4. Zero production hits / acceptance presence for `workbookPlanAcceptance`, `__hermesExcelWorkbookPlanAcceptance`, `__hermesExcelAttachWorkbookPlanAcceptance`, and `workbookPlanAcceptance=1`.

Full WorkbookPlan acceptance still requires a native launcher implementing workbook creation/identity/cleanup, Mac and Windows live runs, host-version evidence, save/reopen, provider testing and network/manifest validation. This harness documents offline route proof only, not full V1 operation-family native coverage. The limited live flow in the README does not close these broader checks.
