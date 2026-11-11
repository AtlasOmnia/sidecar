# WorkbookPlan conformance evidence

The limited desktop Excel for Mac bridge/table/formula flow is recorded in [the README](../README.md).
Windows desktop Excel is untested. That smoke test does not establish full WorkbookPlan acceptance.

## Reproduce offline checks

Record the exact SHA, tree, schema fingerprint, corpus hashes, tool versions and command results.
When comparing a candidate and a landed change, record their patch-id; source review is not runtime testing.

`node scripts/run-workbook-plan-conformance.mjs --out <absolute-existing-scratch-directory> --candidate <40-lowercase-hex-SHA>`
checks frozen original responses and synthetic fixtures. An unavailable disposition is not a success;
contractDrift records a pinned legacy-shape contradiction, not a current-contract PASS. Exit zero requires
zero unexpected failures/errors, but is not whole-release acceptance. Do not repair frozen responses to pass.

The same command with `--shadow` selects the synthetic shadow corpus. Primary execution uses primitive
authority; shadow execution is trapped. An expected comparison mismatch can pass its fixture expectation
without proving the paths equivalent. Synthetic latency/token inputs are not provider measurements.
Unknown execution outcomes remain unknown, not verified success.

The shadow comparison tests retain case labels S1 (primary/shadow execution) through S4 (unchanged default
records), plus outcome and exit-status checks. See `src/__tests__/workbook-plan-shadow-comparison.test.ts`.

## Build and acceptance boundaries

Build production and acceptance artifacts separately. Record inventories, entrypoints and taskpane.js sizes;
production must exclude acceptance entrypoints and synthetic fixtures. Offline tests do not prove native
Mac/Windows execution, installed guidance, provider compliance or rollback. Those need separate live evidence.

Historical gate label H-NET-1 refers to dependency/lockfile, audit and validation checks. It is not a current
publication prohibition; publication steps are in [publish-runbook.md](publish-runbook.md).
