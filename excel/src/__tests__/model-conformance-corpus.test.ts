import { describe, expect, it, beforeAll } from "vitest";
import { readdir, readFile } from "fs/promises";
import { join } from "path";
import { createHash } from "crypto";

const FIXTURE_DIR = join(__dirname, "fixtures", "model-conformance");

/* ── enums ─────────────────────────────────────────────────────── */
const ALLOWED_TRANSPORTS = new Set([
  "openai-tool-call",
  "xml-fallback",
]);
const ALLOWED_DISPOSITIONS = new Set([
  "accepted",
  "accepted-normalized",
  "rejected-syntax-alias",
  "rejected-malformed-json",
  "rejected-unknown-tool",
  "rejected-conflicting-alias",
  "rejected-duplicate-write",
  "rejected-invalid-operation",
  "rejected-unsupported-host",
  "rejected-plan-validation",
]);

/* ── required coverage ─────────────────────────────────────────── */
const REQUIRED_FAILURE_CLASSES = [
  "command-drift",
  "malformed-json",
  "hallucinated-tool",
  "conflicting-aliases",
  "duplicate-write",
  "partial-multi-operation",
  "unsupported-host",
];

const REQUIRED_METADATA_FIELDS = [
  "caseId",
  "sourceClass",
  "transport",
  "description",
  "syntheticInput",
  "modelOutput",
  "expectedDisposition",
  "failureClass",
];

/* ── sensitive-data patterns ───────────────────────────────────── */
const SENSITIVE_PATTERNS = [
  /\/Users\//,
  /\/home\//,
  /\/C:\\/,
  /\\\\Users\\\\/,
  /sk-[A-Za-z0-9]{20,}/,
  /api[_-]?key/i,
  /token[=:\s]/i,
  /"token"\s*:/i,
  /"access_token"\s*:/i,
  /"refresh_token"\s*:/i,
  /"api_key"\s*:/i,
  /"api-key"\s*:/i,
  /uuid[:=]/i,
  /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i,
  /bridgeId/i,
  /\bSyntheticOrg\b/i,
  /\bPrivateOrgSentinel\b/,
  /\bpayroll\b/i,
  /\bclaim[s]?\b/i,
  /\bhealth\b/i,
  /\bpersonal\b/i,
  /\bsocial.security\b/i,
  /\bmedicare\b/i,
  /\bdiagnosis\b/i,
  /\bclaimNumber\b/i,
];

/* ── fixture type ──────────────────────────────────────────────── */
interface FixtureMeta {
  caseId: string;
  sourceClass: string;
  transport: string;
  description: string;
  syntheticInput: unknown;
  modelOutput: unknown;
  expectedDisposition: string;
  failureClass: string | null;
  notes?: string;
}

/* ── helpers ───────────────────────────────────────────────────── */
function sha256hex(data: Buffer): string {
  return createHash("sha256").update(data).digest("hex");
}

/* ── tests ─────────────────────────────────────────────────────── */
describe("model-conformance-corpus", () => {
  let fixtures: FixtureMeta[];
  let sortedFileNames: string[];
  let fileHashes: Record<string, string>;
  let corpusSha256: string;

  beforeAll(async () => {
    const files = (await readdir(FIXTURE_DIR)).filter(
      (f) => f.endsWith(".json") && f !== "manifest.json"
    );
    sortedFileNames = files.sort();
    fixtures = [];
    fileHashes = {};
    for (const file of sortedFileNames) {
      const buf = await readFile(join(FIXTURE_DIR, file));
      const raw = buf.toString("utf-8");
      fixtures.push(JSON.parse(raw));
      fileHashes[file] = sha256hex(buf);
    }
    // Compute corpus aggregate: SHA-256 of sorted "filename:digest" pairs
    const combined = sortedFileNames
      .map((f) => `${f}:${fileHashes[f]}`)
      .join("");
    corpusSha256 = sha256hex(Buffer.from(combined, "utf-8"));
  });

  /* ── 1. count and unique caseId ──────────────────────────────── */
  describe("fixture count and caseId uniqueness", () => {
    it("has exactly 12 fixture files", () => {
      expect(fixtures.length).toBe(12);
    });

    it("has unique caseId values across all fixtures", () => {
      const ids = fixtures.map((f) => f.caseId);
      expect(new Set(ids).size).toBe(ids.length);
    });
  });

  /* ── 2. valid JSON and stable sorted inventory ───────────────── */
  describe("valid JSON and file inventory", () => {
    it("all fixture files parse as valid JSON", () => {
      // If beforeAll succeeded, all JSON.parse calls succeeded. This is a
      // structural guarantee: the loader would have thrown on bad JSON.
      expect(fixtures.length).toBeGreaterThan(0);
    });

    it("sorted file inventory is stable", () => {
      expect(sortedFileNames).toEqual([
        "conflicting-alias-001.json",
        "deepseek-drift-001.json",
        "deepseek-drift-002.json",
        "deepseek-drift-003.json",
        "duplicate-write-001.json",
        "hallucinated-tool-001.json",
        "malformed-json-001.json",
        "openai-tool-call-001.json",
        "partial-multi-op-001.json",
        "syntax-alias-001.json",
        "unsupported-host-001.json",
        "xml-fallback-001.json",
      ]);
    });
  });

  /* ── 3. transport enum ───────────────────────────────────────── */
  describe("transport coverage", () => {
    it("every fixture uses an allowed transport", () => {
      for (const f of fixtures) {
        expect(ALLOWED_TRANSPORTS.has(f.transport)).toBe(true);
      }
    });

    it("covers both accepted transports", () => {
      const transports = new Set(fixtures.map((f) => f.transport));
      expect(transports.has("openai-tool-call")).toBe(true);
      expect(transports.has("xml-fallback")).toBe(true);
    });
  });

  /* ── 4. expectedDisposition enum ─────────────────────────────── */
  describe("expectedDisposition enum", () => {
    it("every fixture uses an allowed disposition", () => {
      for (const f of fixtures) {
        expect(ALLOWED_DISPOSITIONS.has(f.expectedDisposition)).toBe(true);
      }
    });
  });

  /* ── 5. required metadata fields ─────────────────────────────── */
  describe("required metadata fields", () => {
    it("every fixture has all required fields", () => {
      for (const f of fixtures) {
        for (const field of REQUIRED_METADATA_FIELDS) {
          expect(f).toHaveProperty(field);
        }
      }
    });

    it("caseId is a non-empty string", () => {
      for (const f of fixtures) {
        expect(typeof f.caseId).toBe("string");
        expect(f.caseId.length).toBeGreaterThan(0);
      }
    });

    it("transport is a non-empty string", () => {
      for (const f of fixtures) {
        expect(typeof f.transport).toBe("string");
        expect(f.transport.length).toBeGreaterThan(0);
      }
    });

    it("description is a non-empty string", () => {
      for (const f of fixtures) {
        expect(typeof f.description).toBe("string");
        expect(f.description.length).toBeGreaterThan(0);
      }
    });
  });

  /* ── 5b. fixture metadata consistency ───────────────────────── */
  describe("fixture metadata consistency", () => {
    const ALLOWED_SOURCE_CLASSES = new Set(["generic", "deepseek-v4-flash"]);
    const REQUIRED_FAILURE_CLASSES_SET = new Set(REQUIRED_FAILURE_CLASSES);

    it("sourceClass is a non-empty string from the frozen corpus source set", () => {
      for (const f of fixtures) {
        expect(typeof f.sourceClass).toBe("string");
        expect(f.sourceClass.length).toBeGreaterThan(0);
        expect(ALLOWED_SOURCE_CLASSES.has(f.sourceClass)).toBe(true);
      }
    });

    it("failureClass is either null or a non-empty string from REQUIRED_FAILURE_CLASSES", () => {
      for (const f of fixtures) {
        if (f.failureClass === null) {
          // allowed
        } else {
          expect(typeof f.failureClass).toBe("string");
          expect(f.failureClass.length).toBeGreaterThan(0);
          expect(REQUIRED_FAILURE_CLASSES_SET.has(f.failureClass)).toBe(true);
        }
      }
    });

    it("expectedDisposition beginning with 'accepted' implies failureClass is null", () => {
      for (const f of fixtures) {
        if (f.expectedDisposition.startsWith("accepted")) {
          expect(f.failureClass).toBe(null);
        }
      }
    });

    it("expectedDisposition beginning with 'rejected' implies failureClass is a non-null REQUIRED_FAILURE_CLASSES value", () => {
      for (const f of fixtures) {
        if (f.expectedDisposition.startsWith("rejected")) {
          expect(f.failureClass).not.toBe(null);
          expect(typeof f.failureClass).toBe("string");
          expect(f.failureClass!.length).toBeGreaterThan(0);
          expect(REQUIRED_FAILURE_CLASSES_SET.has(f.failureClass!)).toBe(true);
        }
      }
    });
  });

  /* ── 6. honest synthetic provenance ──────────────────────────── */
  describe("synthetic provenance", () => {
    it("no fixture claims to be a raw provider response", () => {
      for (const f of fixtures) {
        const raw = JSON.stringify(f).toLowerCase();
        expect(raw).not.toContain("raw provider response");
        expect(raw).not.toContain("actual provider output");
        expect(raw).not.toContain("captured response");
      }
    });

    it("every fixture has a non-null syntheticInput", () => {
      for (const f of fixtures) {
        expect(f.syntheticInput).not.toBeNull();
        expect(typeof f.syntheticInput).toBe("object");
      }
    });

    it("every fixture has a non-null modelOutput", () => {
      for (const f of fixtures) {
        expect(f.modelOutput).not.toBeNull();
        // XML fallback uses a raw string envelope; OpenAI uses an object
        expect(
          typeof f.modelOutput === "object" ||
            typeof f.modelOutput === "string"
        ).toBe(true);
      }
    });
  });

  /* ── 7. hygiene: no sensitive data ───────────────────────────── */
  describe("sensitive data hygiene", () => {
    it("detects neutral synthetic organization sentinels", () => {
      for (const sentinel of ["SyntheticOrg", "PrivateOrgSentinel"]) {
        expect(SENSITIVE_PATTERNS.some(pattern => pattern.test(sentinel))).toBe(true);
      }
    });

    it("no absolute macOS/Windows paths or home-directory markers", () => {
      for (const f of fixtures) {
        const raw = JSON.stringify(f);
        for (const pattern of SENSITIVE_PATTERNS) {
          expect(raw).not.toMatch(pattern);
        }
      }
    });

    it("no UUID/session ID patterns", () => {
      for (const f of fixtures) {
        const raw = JSON.stringify(f);
        // Match UUID v4 pattern but exclude synthetic call_XXX IDs
        const uuidPattern =
          /[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}/i;
        expect(raw).not.toMatch(uuidPattern);
      }
    });

    it("no API key or token patterns", () => {
      for (const f of fixtures) {
        const raw = JSON.stringify(f);
        expect(raw).not.toMatch(/sk-[A-Za-z0-9]{20,}/);
        expect(raw).not.toMatch(/api[_-]?key["']?\s*[:=]/i);
        expect(raw).not.toMatch(/bearer\s+[A-Za-z0-9]/i);
        // Quoted-JSON-key variants are caught by SENSITIVE_PATTERNS
      }
    });

    it("no sensitive SyntheticOrg/PrivateOrgSentinel/payroll/claims/health/personal terms", () => {
      const sensitiveTerms = [
        /\bSyntheticOrg\b/,
        /\bPrivateOrgSentinel\b/,
        /\bpayroll\b/i,
        /\bclaim[s]?\b/i,
        /\bhealth\b/i,
        /\bpersonal\b/i,
        /\bsocial.security\b/i,
        /\bmedicare\b/i,
        /\bdiagnosis\b/i,
      ];
      for (const f of fixtures) {
        const raw = JSON.stringify(f);
        for (const term of sensitiveTerms) {
          expect(raw).not.toMatch(term);
        }
      }
    });

    it("no bridge IDs", () => {
      for (const f of fixtures) {
        const raw = JSON.stringify(f);
        expect(raw).not.toMatch(/bridgeId/i);
        expect(raw).not.toMatch(/bridge[_-]?session/i);
      }
    });
  });

  /* ── 8. failure-class coverage ───────────────────────────────── */
  describe("failure-class coverage", () => {
    it("covers every required failure class", () => {
      const classes = new Set(
        fixtures.filter((f) => f.failureClass !== null).map((f) => f.failureClass)
      );
      for (const cls of REQUIRED_FAILURE_CLASSES) {
        expect(classes.has(cls)).toBe(true);
      }
    });

    it("at least 3 command-drift fixtures exist", () => {
      const driftCount = fixtures.filter(
        (f) => f.failureClass === "command-drift"
      ).length;
      expect(driftCount).toBeGreaterThanOrEqual(3);
    });
  });

  /* ── 9. accepted-transport coverage ──────────────────────────── */
  describe("accepted-transport coverage", () => {
    it("at least one accepted fixture per transport", () => {
      const acceptedByTransport = new Map<string, number>();
      for (const f of fixtures) {
        if (f.expectedDisposition.startsWith("accepted")) {
          acceptedByTransport.set(
            f.transport,
            (acceptedByTransport.get(f.transport) || 0) + 1
          );
        }
      }
      expect(acceptedByTransport.get("openai-tool-call")).toBeGreaterThanOrEqual(
        1
      );
      expect(acceptedByTransport.get("xml-fallback")).toBeGreaterThanOrEqual(1);
    });
  });

  /* ── 9b. transport-envelope structure (corpus-shape only) ───── */
  describe("transport envelope structure", () => {
    it("openai-tool-call fixtures expose a non-empty tool_calls array", () => {
      const openai = fixtures.filter(
        (f) => f.transport === "openai-tool-call"
      );
      expect(openai.length).toBeGreaterThan(0);
      for (const f of openai) {
        const out = f.modelOutput as { tool_calls: unknown[] };
        expect(Array.isArray(out.tool_calls)).toBe(true);
        expect(out.tool_calls.length).toBeGreaterThan(0);
      }
    });

    it("openai-tool-call tool_calls entries have function-call envelope fields", () => {
      const openai = fixtures.filter(
        (f) => f.transport === "openai-tool-call"
      );
      for (const f of openai) {
        const out = f.modelOutput as {
          tool_calls: Array<{
            function: { name: unknown; arguments: unknown };
          }>;
        };
        for (const tc of out.tool_calls) {
          expect(typeof tc.function.name).toBe("string");
          expect((tc.function.name as string).length).toBeGreaterThan(0);
          expect(typeof tc.function.arguments).toBe("string");
        }
      }
    });

    it("xml-fallback fixtures expose a bounded <tool_call>...</tool_call> envelope", () => {
      const xml = fixtures.filter((f) => f.transport === "xml-fallback");
      expect(xml.length).toBeGreaterThan(0);
      for (const f of xml) {
        const out = f.modelOutput as string;
        // Must contain a bounded tool_call tag pair directly in the string
        expect(out).toMatch(/<tool_call>[\s\S]*?<\/tool_call>/);
      }
    });

    it("xml-fallback envelope inner text is JSON with non-empty name and object-valued arguments", () => {
      const xml = fixtures.filter((f) => f.transport === "xml-fallback");
      for (const f of xml) {
        const out = f.modelOutput as string;
        const match = out.match(/<tool_call>([\s\S]*?)<\/tool_call>/);
        expect(match).not.toBeNull();
        const inner = match![1];
        expect(() => JSON.parse(inner)).not.toThrow();
        const parsed = JSON.parse(inner) as Record<string, unknown>;
        expect(typeof parsed.name).toBe("string");
        expect((parsed.name as string).length).toBeGreaterThan(0);
        expect(parsed.arguments).not.toBeNull();
        expect(typeof parsed.arguments).toBe("object");
        expect(Array.isArray(parsed.arguments)).toBe(false);
      }
    });
  });

  /* ── 9c. token-hygiene sentinel ──────────────────────────────── */
  describe("token hygiene detector catches ordinary JSON fields", () => {
    it("SENSITIVE_PATTERNS catches quoted JSON key patterns: token, access_token, refresh_token, api_key, api-key", () => {
      // Sentinel: raw strings containing dangerous field names as quoted JSON keys.
      // The live SENSITIVE_PATTERNS array must flag them without storing any real secret.
      const sentinelCases = [
        '"token":"sentinel_token_value"',
        '"access_token":"sentinel_access_token_value"',
        '"refresh_token":"sentinel_refresh_token_value"',
        '"api_key":"sentinel_api_key_value"',
        '"api-key":"sentinel_api-key_value"',
      ];
      for (const sentinel of sentinelCases) {
        let matched = false;
        for (const pat of SENSITIVE_PATTERNS) {
          if (pat.test(sentinel)) {
            matched = true;
            break;
          }
        }
        expect(
          matched,
          `Sentinel "${sentinel}" should be caught by SENSITIVE_PATTERNS`
        ).toBe(true);
      }
    });
  });

  /* ── 10. manifest.json contract ──────────────────────────────── */
  describe("manifest.json contract", () => {
    let manifest: Record<string, unknown>;

    beforeAll(async () => {
      const raw = await readFile(
        join(FIXTURE_DIR, "manifest.json"),
        "utf-8"
      );
      manifest = JSON.parse(raw);
    });

    it("has fixtureCount equal to 12", () => {
      expect(manifest.fixtureCount).toBe(12);
    });

    it("has transports equal to exactly ['openai-tool-call', 'xml-fallback']", () => {
      expect(manifest.transports).toEqual(["openai-tool-call", "xml-fallback"]);
    });

    it("has files mapping every sorted fixture filename to its SHA-256 digest", () => {
      const files = manifest.files as Record<string, string>;
      expect(files).toBeDefined();

      // No extra files
      const manifestFilenames = Object.keys(files).sort();
      expect(manifestFilenames).toEqual(sortedFileNames);

      // Every digest matches the computed hash
      for (const fname of sortedFileNames) {
        expect(files[fname]).toBe(fileHashes[fname]);
      }
    });

    it("has corpusSha256 matching the deterministic aggregate", () => {
      expect(manifest.corpusSha256).toBe(corpusSha256);
    });

    it("has no generatedAt field", () => {
      expect(manifest).not.toHaveProperty("generatedAt");
    });

    it("has no plain-json transport in transports", () => {
      const transports = manifest.transports as string[];
      expect(transports).not.toContain("plain-json");
    });

    it("has no extra/missing manifest file entries", () => {
      const files = manifest.files as Record<string, string>;
      expect(Object.keys(files).sort()).toEqual(sortedFileNames);
    });
  });

  /* ── 11. no skipped tests ────────────────────────────────────── */
  // (all tests above are active — no .skip or .todo calls)

  /* ── 12. deepseek-drift-002 semantic fixture ─────────────────── */
  describe("deepseek-drift-002: insert_table with range vs address drift", () => {
    it("tool name is canonical 'insert_table' (not create_table)", () => {
      const f = fixtures.find((f) => f.caseId === "deepseek-drift-002");
      expect(f).toBeDefined();
      const out = f!.modelOutput as { tool_calls: Array<{ function: { name: string; arguments: string } }> };
      expect(out.tool_calls[0].function.name).toBe("insert_table");
    });

    it("arguments parse as JSON containing 'range' (not 'address')", () => {
      const f = fixtures.find((f) => f.caseId === "deepseek-drift-002");
      const out = f!.modelOutput as { tool_calls: Array<{ function: { name: string; arguments: string } }> };
      const args = JSON.parse(out.tool_calls[0].function.arguments);
      expect(args).toHaveProperty("range");
      expect(args).not.toHaveProperty("address");
    });

    it("description and notes consistently state tool name is canonical with range drifting from address", () => {
      const f = fixtures.find((f) => f.caseId === "deepseek-drift-002");
      // Description says "tool name is canonical" (correct — it IS insert_table)
      // and property 'range' drifts from canonical 'address'
      expect(f!.description).toMatch(/tool name is canonical/i);
      expect(f!.description).toMatch(/range/);
      expect(f!.description).toMatch(/address/);
      // Notes say tool name is correct, only property alias mismatches
      expect(f!.notes).toMatch(/tool name is correct/i);
      expect(f!.notes).toMatch(/range/);
      expect(f!.notes).toMatch(/address/);
    });
  });

  /* ── 13. accepted OpenAI fixture structural validation ───────── */
  describe("accepted OpenAI fixture structural validation", () => {
    const acceptedOpenAI = () =>
      fixtures.filter(
        (f) =>
          f.transport === "openai-tool-call" &&
          (f.expectedDisposition === "accepted" ||
            f.expectedDisposition === "accepted-normalized")
      );

    it("every accepted OpenAI function.arguments parses to an object", () => {
      const accepted = acceptedOpenAI();
      expect(accepted.length).toBeGreaterThan(0);

      for (const f of accepted) {
        const out = f.modelOutput as {
          tool_calls?: Array<{
            function?: { name?: unknown; arguments?: unknown };
          }>;
        };
        expect(Array.isArray(out.tool_calls)).toBe(true);
        expect(out.tool_calls!.length).toBeGreaterThan(0);

        for (const tc of out.tool_calls!) {
          expect(typeof tc.function?.arguments).toBe("string");
          const parsed = JSON.parse(tc.function!.arguments as string);
          expect(parsed).not.toBeNull();
          expect(typeof parsed).toBe("object");
          expect(Array.isArray(parsed)).toBe(false);
        }
      }
    });

    it("accepted excel_execute_plan arguments use the V1 kind discriminators", () => {
      const planCalls = acceptedOpenAI().flatMap((f) => {
        const out = f.modelOutput as {
          tool_calls: Array<{
            function: { name: string; arguments: string };
          }>;
        };
        return out.tool_calls.filter(
          (tc) => tc.function.name === "excel_execute_plan"
        );
      });
      expect(planCalls.length).toBeGreaterThan(0);

      for (const tc of planCalls) {
        const plan = JSON.parse(tc.function.arguments) as {
          version?: unknown;
          outcome?: unknown;
          summary?: unknown;
          operations?: unknown;
          verification?: unknown;
        };
        expect(plan.version).toBe(1);
        expect(plan.outcome).toBe("execute");
        expect(typeof plan.summary).toBe("string");
        expect((plan.summary as string).length).toBeGreaterThan(0);
        expect(Array.isArray(plan.operations)).toBe(true);
        expect((plan.operations as unknown[]).length).toBeGreaterThan(0);
        expect(Array.isArray(plan.verification)).toBe(true);
        expect((plan.verification as unknown[]).length).toBeGreaterThan(0);

        for (const operation of plan.operations as unknown[]) {
          expect(operation).not.toBeNull();
          expect(typeof operation).toBe("object");
          expect(Array.isArray(operation)).toBe(false);
          const record = operation as Record<string, unknown>;
          expect(typeof record.kind).toBe("string");
          expect((record.kind as string).length).toBeGreaterThan(0);
          expect(record).not.toHaveProperty("op");
        }

        for (const rule of plan.verification as unknown[]) {
          expect(rule).not.toBeNull();
          expect(typeof rule).toBe("object");
          expect(Array.isArray(rule)).toBe(false);
          const record = rule as Record<string, unknown>;
          expect(typeof record.kind).toBe("string");
          expect((record.kind as string).length).toBeGreaterThan(0);
          expect(record).not.toHaveProperty("rule");
        }
      }
    });
  });

  /* ── 14. no source-string assertions pretending runtime ──────── */
  // All assertions above test structural/content properties of the fixtures
  // themselves, not runtime parsing behavior. Behavioral assertions
  // (parsing plans, executing operations) are deferred to Phase 5.
});
