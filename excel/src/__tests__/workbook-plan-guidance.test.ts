import { describe, expect, it } from "vitest";
import { cpSync, existsSync, mkdirSync, mkdtempSync, readFileSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { workbookPlanSchema, workbookPlanFingerprint } from "../generated/workbookPlanSchema";
import { SERVER_INSPECTION_SUPPORTED } from "../taskpane/serverOwnedPlanRoute";
import { getDirectSystemPrompt, getServerOwnedSystemPrompt } from "../agent";

const root = resolve(__dirname, "../..");
const schemaPath = join(root, "plugins/atlas-sidecar-excel/tool_schemas.json");
const skillPath = join(root, "plugins/atlas-sidecar-excel/skills/workbook-plans/SKILL.md");
const generator = join(root, "scripts/generate-office-plugin-schemas.mjs");
const hash = (s: string | Buffer) => createHash("sha256").update(s).digest("hex");
const run = (...args: string[]) => spawnSync(process.execPath, [generator, ...args], { cwd: root, encoding: "utf8", timeout: 60000 });
const temp = () => mkdtempSync(join(tmpdir(), "r6a-guidance-"));
const skill = () => existsSync(skillPath) ? readFileSync(skillPath, "utf8") : "";
const identities = () => [schemaPath, skillPath, join(root, "src/generated/workbookPlanSchema.ts"), join(root, "src/agent.ts")].map(p => existsSync(p) ? hash(readFileSync(p)) : "missing");

describe("generated workbook plan guidance", () => {
  it("G1 real drift check passes and the generated skill exists", () => {
    expect(run("--check").status).toBe(0);
    expect(existsSync(skillPath)).toBe(true);
  }, 60_000);
  it("G2/G3 skill drift and missing skill fail without writing real or temp artifacts", () => {
    const d = temp(); const output = join(d, "schemas.json"); const target = join(d, "SKILL.md");
    writeFileSync(output, readFileSync(schemaPath));
    writeFileSync(target, skill() + "x");
    const before = identities(); const staleBefore = readFileSync(target);
    const stale = run("--check", "--output", output, "--skill", target);
    expect(stale.status).toBe(1);
    expect(stale.stderr).toContain("Generated workbook plan skill is stale");
    expect(readFileSync(target)).toEqual(staleBefore);
    expect(identities()).toEqual(before);
    const missing = run("--check", "--output", output, "--skill", join(d, "missing.md"));
    expect(missing.status).toBe(1);
    expect(missing.stderr).toContain("Generated workbook plan skill is missing");
    expect(existsSync(join(d, "missing.md"))).toBe(false);
    expect(identities()).toEqual(before);
  }, 120000);
  it("G4 schema drift is checked independently", () => {
    const d = temp(); const output = join(d, "schemas.json"); const target = join(d, "SKILL.md");
    writeFileSync(output, "[]\n"); writeFileSync(target, skill());
    const before = identities();
    const result = run("--check", "--output", output, "--skill", target);
    expect(result.status).toBe(1);
    expect(result.stderr).toContain("Generated Excel plugin schemas are stale");
    expect(result.stderr).not.toContain("Generated workbook plan skill is stale");
    expect(readFileSync(output, "utf8")).toBe("[]\n");
    expect(identities()).toEqual(before);
  }, 60_000);
  it("G5/G6 derives fingerprint, closed fields, outcomes and operation bound from the canonical schema", () => {
    const text = skill();
    const tool = JSON.parse(readFileSync(schemaPath, "utf8")).find((s: { name: string }) => s.name === "sidecar_excel_execute_plan");
    expect(text).toContain(`Schema fingerprint: ${workbookPlanFingerprint}.`);
    expect(tool["x-workbook-plan-fingerprint"]).toBe(workbookPlanFingerprint);
    expect(text).toContain("additionalProperties: false");
    const allowed = [...Object.keys(workbookPlanSchema.properties), "sidecar_excel_execute_plan"];
    const tokens = [...text.matchAll(/`([^`]+)`/g)].map(m => m[1]);
    expect(tokens.length).toBeGreaterThan(1);
    for (const token of tokens) expect(allowed).toContain(token);
    expect(text).toContain(`Outcomes: ${workbookPlanSchema.properties.outcome.enum.join(", ")}.`);
    expect(text).toContain(`Maximum operations: ${workbookPlanSchema.properties.operations.maxItems}.`);
    expect(text).not.toContain("idempotencyKey");
  });
  it("G7 pins supported server inspection, no retry, caps and default-off truth", () => {
    expect(SERVER_INSPECTION_SUPPORTED).toBe(true);
    const text = skill();
    expect(text).toContain("at most one inspect, then at most one execute; no plan after execute");
    expect(text).toContain("Server inspection is supported");
    expect(text).toContain("rollout is default off");
    expect(text).toContain("1 MiB per read, 1.5 MiB for the body, nesting depth 64, and 5,000 cells");
    expect(text).toContain("Non-executing reasons: ask, refuse, rejected, cancelled, not_started, invalid, already_executed.");
    expect(text).toContain("Do not retry after already_executed");
    expect(text).toContain("read the tool result before claiming success");
    expect(text).toContain("nothing executes from the plan alone");
    expect(text).toContain("destructive and replacement fields are not part of V1");
  });
  it("G8 write overrides are deterministic, LF-only and isolated; CRLF check is accepted", () => {
    const a = temp(); const b = temp(); const before = identities();
    for (const d of [a, b]) expect(run("--output", join(d, "schemas.json"), "--skill", join(d, "SKILL.md")).status).toBe(0);
    for (const f of ["schemas.json", "SKILL.md"]) expect(readFileSync(join(a, f))).toEqual(readFileSync(join(b, f)));
    const text = readFileSync(join(a, "SKILL.md"), "utf8");
    expect(text).toBe(skill()); expect(text).not.toContain("\r"); expect(text.endsWith("\n")).toBe(true);
    expect(text.startsWith("---\nname: workbook-plans\ndescription:")).toBe(true);
    writeFileSync(join(a, "SKILL.md"), text.replace(/\n/g, "\r\n"));
    expect(run("--check", "--output", join(a, "schemas.json"), "--skill", join(a, "SKILL.md")).status).toBe(0);
    expect(identities()).toEqual(before);
  }, 180000);
  it("rejects duplicate, missing-value and unknown flags without writes", () => {
    const before = identities();
    for (const args of [["--skill", "a", "--skill", "b"], ["--output"], ["--skill", "--check"], ["--unknown"], ["--check", "--check"]]) {
      const result = run(...args); expect(result.status).not.toBe(0); expect(result.stderr).toMatch(/argument|flag|value/i);
    }
    expect(identities()).toEqual(before);
  }, 120000);
  it("R1 guidance allows terminal narration and describes replay uncertainty and text limits", () => {
    const text = skill();
    expect(text).toContain("At most one plan call per response.");
    expect(text).not.toContain("Use exactly one");
    expect(text).toContain("After an execute plan call returns, or after an already_executed / uncertain reply, make no further plan calls in that turn; report the outcome instead.");
    expect(text).toContain("An inspect may be followed by at most one execute for the same turn.");
    expect(text).toContain("same trusted turnKey");
    expect(text).toContain("an earlier attempt of this same operation may already have changed the workbook. Do not run it again.");
    expect(text).toContain("Read the recorded execution report in the reply and tell the user what it shows (it may be partial or unknown).");
    expect(text).not.toContain("was already applied earlier");
    expect(text).toContain("2,048 Unicode code points");
    expect(text).toContain("ask / refuse / finalSummary text");
    expect(text).toContain("serialized server inspection results");
    expect(text).toContain("Plan within these limits");
  });
  // Run destructive regression controls only against complete scratch copies.
  it.each(["skill-default", "output-default", "explicit", "source", "source-link", "file-link", "parent-link", "missing-parent-link", "dangling-link"])("R1 rejects %s alias before any artifact or source write", (kind) => {
    const d = temp();
    mkdirSync(join(d, "scripts"));
    cpSync(generator, join(d, "scripts/generate-office-plugin-schemas.mjs"));
    cpSync(join(root, "src"), join(d, "src"), { recursive: true });
    cpSync(join(root, "plugins"), join(d, "plugins"), { recursive: true });
    symlinkSync(join(root, "node_modules"), join(d, "node_modules"), "dir");
    const schema = join(d, "plugins/atlas-sidecar-excel/tool_schemas.json");
    const target = join(d, "plugins/atlas-sidecar-excel/skills/workbook-plans/SKILL.md");
    const source = join(d, "src/generated/workbookPlanSchema.ts");
    let args: string[];
    if (kind === "skill-default") args = ["--skill", schema];
    else if (kind === "output-default") args = ["--output", target];
    else if (kind === "explicit") args = ["--output", schema, "--skill", schema];
    else if (kind === "source") args = ["--output", source];
    else if (kind === "source-link") {
      symlinkSync(source, join(d, "alias")); args = ["--skill", join(d, "alias")];
    } else if (kind === "file-link") {
      symlinkSync(schema, join(d, "alias")); args = ["--skill", join(d, "alias")];
    } else if (kind === "dangling-link") {
      symlinkSync("missing.md", join(d, "alias"));
      args = ["--output", join(d, "missing.md"), "--skill", join(d, "alias")];
    } else {
      symlinkSync(join(d, "plugins/atlas-sidecar-excel"), join(d, "alias"), "dir");
      const suffix = kind === "missing-parent-link" ? "new/nested/file" : "tool_schemas.json";
      args = ["--output", join(d, "plugins/atlas-sidecar-excel", suffix), "--skill", join(d, "alias", suffix)];
    }
    const paths = [schema, target, source, join(d, "scripts/generate-office-plugin-schemas.mjs")];
    const before = paths.map(p => hash(readFileSync(p)));
    const result = spawnSync(process.execPath, [join(d, "scripts/generate-office-plugin-schemas.mjs"), ...args], { cwd: d, encoding: "utf8", timeout: 60000 });
    expect(result.status).not.toBe(0);
    expect(result.stderr).toMatch(/different files|alias|input|source/i);
    expect(paths.map(p => hash(readFileSync(p)))).toEqual(before);
    expect(existsSync(join(d, "plugins/atlas-sidecar-excel/new"))).toBe(false);
    expect(existsSync(join(d, "missing.md"))).toBe(false);
  }, 120000);
  it("G9 base oracle: prompt hashes remain unchanged", () => {
    expect(hash(getDirectSystemPrompt(false))).toBe("5f2c25809f4d0c24a16f3626da1ba82dd1d7d2b717571d612d04bdcc2e781e11");
    expect(hash(getDirectSystemPrompt(true))).toBe("d89aaed26ade8dcf784af6dcc78e57eccf1c6ff3e640f61eedf1e4a260cbc70f");
    expect(hash(getServerOwnedSystemPrompt())).toBe("7ee6b2428175d749c63376dd0b58872f32e53bdb8b15d8694f9b490580f2bea2"); // NAMES: only desk_excel_* -> sidecar_excel_* in the prompt;
  });
  const doc = (name: string) => readFileSync(join(root, "docs", name), "utf8");
  it("S7 runtime docs trace current routing, defaults and trust boundary", () => {
    const text = doc("workbook-plan-runtime.md");
    for (const s of ["default off", "serverOwnedPlanRoute", "taskPaneTurnCoordinator", "prepareWorkbookPlanExecution", "unknown", "may have been applied", "correction", "precondition", "synthetic"]) expect(text).toContain(s);
    expect(text).not.toMatch(/server inspection (?:is )?unsupported/i);
  });
  it("S8 skill docs describe generated artifact without installing it", () => {
    const text = doc("workbook-plan-skill.md");
    for (const s of ["SKILL.md", "generate-office-plugin-schemas.mjs", "--check", "additionalProperties: false", "already_executed", "default off", "HOLD", workbookPlanFingerprint]) expect(text).toContain(s);
    for (const key of Object.keys(workbookPlanSchema.properties)) expect(text).toContain(key);
    expect(text).toContain("at most one inspect");
  });
  it("S9 PR docs separate offline proof from production acceptance", () => {
    const text = doc("workbook-plan-pr-rollout.md");
    for (const s of ["exact SHA", "frozen", "contractDrift", "unavailable", "S1", "S4", "taskpane.js", "H-NET-1", "provider", "unknown", "patch-id"]) expect(text).toContain(s);
  });
  it("S10 rollout docs preserve real adoption/native/provider HOLDs", () => {
    const text = doc("workbook-plan-rollout.md");
    for (const s of ["default off", "HOLD", "Task 23", "Mac", "Windows", "RB5", "rollback", "installed", "provider", "npm run validate", "npm audit"]) expect(text).toContain(s);
    expect(text).toContain("not whole-release acceptance");
  });
});
