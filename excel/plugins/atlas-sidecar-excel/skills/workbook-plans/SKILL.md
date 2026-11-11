---
name: workbook-plans
description: Submit bounded WorkbookPlan V1 inspection and execution through the static Excel plan tool.
---

# Workbook plans

At most one plan call per response. When submitting a plan, use the static `sidecar_excel_execute_plan` tool. Do not mix plans with primitive tool calls.
After an execute plan call returns, or after an already_executed / uncertain reply, make no further plan calls in that turn; report the outcome instead. An inspect may be followed by at most one execute for the same turn.
Allowed top-level fields (additionalProperties: false): `version`, `outcome`, `reason`, `summary`, `message`, `reads`, `operations`, `verification`.
Outcomes: inspect, execute, ask, refuse.
Maximum operations: 30.

- Server inspection is supported when the plan rollout is enabled; rollout is default off.
- When workbook facts are needed, emit an inspection plan first and read its result before proposing execution.
- Per server turn: at most one inspect, then at most one execute; no plan after execute. Inspect and execute must share the same trusted turnKey, associated by the host, not supplied in model arguments. Inspection is optional when the required workbook facts are already known.
- Plan within these limits for serialized server inspection results: 1 MiB per read, 1.5 MiB for the body, nesting depth 64, and 5,000 cells returned to the server model. Byte limits count compact JSON in UTF-8; depth counts the read object as level 1.
- Bound ask / refuse / finalSummary text to 2,048 Unicode code points. Ask/refuse messages are forwarded as finalSummary; execute summary has the same 2,048-code-point bound.
- Wait for and read the tool result before claiming success; nothing executes from the plan alone.
- Keep plans bounded and explicit: destructive and replacement fields are not part of V1. Do not invent private routing, session, policy, limit, request, idempotency or fingerprint fields.

Non-executing reasons: ask, refuse, rejected, cancelled, not_started, invalid, already_executed.
Do not retry after already_executed: an earlier attempt of this same operation may already have changed the workbook. Do not run it again. Read the recorded execution report in the reply and tell the user what it shows (it may be partial or unknown). Do not treat an uncertain result as proof that nothing changed.

Schema fingerprint: 5fcab038f26bef8423d0ec2271b283947c87256bfe69f250205ea348cbfc83c1.
