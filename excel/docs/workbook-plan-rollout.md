# WorkbookPlan rollout and rollback

The limited desktop Excel for Mac bridge/table/formula flow is recorded in [the README](../README.md).
Windows desktop Excel is untested. WorkbookPlan rollout is default off. Offline synthetic evidence is
not whole-release acceptance and does not enable settings or adopt a deployment.

## Before enabling workbook plans

- Preserve the exact candidate SHA, schema/corpus hashes and check results, including unavailable or drift rows.
- Exercise native Mac and Windows cases in disposable workbooks: mutation readback, zero-mutation rejection,
  replay, cancellation, session/document rotation, partial/unknown execution and confirmation expiry
  (the tests' historical RB5 label). Harness preparation is not proof those hosts passed.
- Verify installed guidance and actual provider behavior; generated repository guidance alone proves neither.
- Record dependency/lockfile review and `npm audit` results. This package defines no `validate` npm script; record any
  separate online manifest validation as its own result. Do not silently change dependencies.
- Keep full WorkbookPlan native acceptance (historically called Task 23) separate from the limited README smoke.
  HOLD in older gate records means missing acceptance evidence, not a blanket prohibition on public release.

## Rollback

Turning the rollout setting off prevents new plan authority but cannot undo prior workbook changes.
For partial/unknown results, preserve and reconcile execution state; never replay to discover whether a write
happened. Retain the previous artifact/version and private lifecycle state for a scoped rollback exercise.
Do not stop unrelated Hermes services. Verify rollback against the actual host before claiming success.

Unit, documentation, schema and synthetic checks do not prove native/provider acceptance, live rollback,
or installed guidance. Publication instructions are in [publish-runbook.md](publish-runbook.md).
