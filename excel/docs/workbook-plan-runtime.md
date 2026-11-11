# WorkbookPlan runtime boundary (R6d)

Scope: current repository source and synthetic offline evidence. Rollout is default off. This document neither enables the pane setting nor approves adoption.

## Route and ownership

`src/taskpane/taskpane.ts` snapshots `settings.workbookPlanRolloutEnabled === true` into the turn. `src/taskpane/taskPaneTurnCoordinator.ts` owns turn/session rotation, cancellation, primitive mutation authority and bound confirmation. `src/taskpane/serverOwnedPlanRoute.ts` handles authenticated companion plan calls; server-owned `hermes-agent` responses are not executed by a second caller-owned tool loop. `src/agent.ts` retains the separate direct-provider route. See AGENTS.md for trusted session-header routing.

Server inspection is supported when rollout is enabled. One trusted turn permits at most one inspect followed by at most one execute; narration after an execute is allowed, another plan is not. Inspection is optional only when facts required for preparation are already known. Host association supplies identity; the model cannot supply session, policy, confirmation or replay authority in plan arguments. Generated guidance bounds inspection output at 1 MiB per read, 1.5 MiB body, depth64, 5000 cells; terminal text is bounded to2048 Unicode code points.

## Execution versus proposing

`src/workbookPlanProtocol.ts` normalizes the static response. `src/workbookPlanValidation.ts` enforces the closed V1 shape. `src/workbookPlanRequest.ts` classifies the trusted user request independently from model arguments. `src/workbookPlanRoute.ts` assembles policy facts; `src/workbookPlanPolicy.ts` requires a trusted snapshot. The host-backed `prepareWorkbookPlanExecution` path in `src/workbookPlanCompiler.ts` materializes policy, compile context and precondition evidence before minting executable authority. A plain compiler inspection result is not an executable authorization.

`src/workbookPlanOrchestrator.ts` bounds the turn to90000ms and one correction. A rejected or ambiguous plan does not mutate through a parser-only bypass. Confirmation is bound to current turn/session/document/plan identity; cancel, expiry or stale generation cannot reuse the candidate. Production expiry coverage is RB5 in `src/__tests__/task-pane-session-rotation-r3b.test.ts`; synthetic evaluator callbacks are not that coverage.

`src/workbookPlanOfficeAdapter.ts` is the native Office preparation/execute adapter. `src/workbookPlanExecutor.ts` revalidates prepared authority and tracks replay/execution outcomes. Direct primitive writes independently pass `src/primitiveMutationAuthority.ts` and its claimed raw dispatcher. This is not permission for unsupported destructive legacy calls.

## Replay and uncertainty

Do not retry `already_executed` or an uncertain effect. A partial/unknown report means an operation may have been applied; it does not establish that nothing changed or that everything succeeded. The unknown-only replay narration is: "This workbook operation may have been applied earlier and was not executed again." Read the recorded report and stop further plan calls in that turn.

## Evidence limits

`src/__tests__/workbook-plan-conformance-runtime.test.ts` covers synthetic evaluator behavior, not native fidelity. `src/__tests__/workbook-plan-shadow-comparison.test.ts` exercises real primitive dispatch against a synthetic host and pure shadow protocol/validation/policy/compile with Office execution trapped. Comparison mismatch, unavailable and contractDrift are not native success. The primary committed label derives the raw successful response; it is not an independent final-state/readback verdict. No provider latency, provider diversity, installed hooks or live Excel acceptance is established here.
