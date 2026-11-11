import { beforeEach, describe, expect, it, vi } from "vitest";
import { executeWordTool, executeWordToolParsed, executeWordToolInContext, toolSchemas } from "../wordTools";
import { executeValidatedWordBridgeTool } from "../officeBridgeWord";
import { getServerOwnedSystemPrompt } from "../serverOwnedMode";
import { readFileSync } from "node:fs";
import { HermesAgent } from "../agent";
import * as wordTools from "../wordTools";
import { isSetSupported, wordRun } from "./setup";

// Office boundary double: commands commit at sync, failures can occur there.
// Does not model native range tracking, rich formatting or Undo.
function host(initial = ["source"]) {
  const queued: Array<() => void> = [];
  let syncCount = 0;
  let failAt = 0;
  let query = "";
  const calls: Array<{ query: string; options: unknown }> = [];
  const paragraph = (text: string): any => {
    let style = "Normal", builtIn = "Normal";
    const p: any = {
      text, load: vi.fn(),
      insertText: vi.fn((value: string, location: string) => { queued.push(() => { p.text = location === "Replace" ? value : p.text + value; }); return { font: {} }; }),
      insertParagraph(value: string, location: string) {
        const next = paragraph(value);
        queued.push(() => paragraphs.splice(paragraphs.indexOf(p) + (location === "After" ? 1 : 0), 0, next));
        return next;
      }
    };
    Object.defineProperties(p, {
      style: { get: () => style, set: (v: string) => queued.push(() => { style = v; builtIn = "Other"; }) },
      styleBuiltIn: { get: () => builtIn, set: (v: string) => queued.push(() => { builtIn = v; style = v === "Heading1" ? "Heading 1" : v; }) }
    });
    return p;
  };
  const paragraphs = initial.map(paragraph);
  const collection = { items: paragraphs, load: vi.fn() };
  const selectionParagraph = paragraph("elsewhere");
  const body: any = {
    paragraphs: collection, load: vi.fn(),
    insertText: () => ({ font: {} }), insertOoxml: () => ({}),
    get text() { return paragraphs.map(p => p.text).join("\r"); },
    insertParagraph(value: string, location: string) {
      const p = paragraph(value); queued.push(() => location === "Start" ? paragraphs.unshift(p) : paragraphs.push(p)); return p;
    },
    search(encoded: string, options: any) {
      query = encoded; calls.push({ query: encoded, options });
      const literal = encoded.replace(/\^\^/g, "^");
      const found = paragraphs.filter(p => options.matchCase ? p.text.includes(literal) : p.text.toLowerCase().includes(literal.toLowerCase()));
      return { load: vi.fn(), items: found.map(p => ({
        paragraphs: { items: [p], load: vi.fn() },
        compareLocationWith: () => ({ value: "After" }),
        insertText(value: string) { queued.push(() => { p.text = p.text.replace(literal, value); }); return { font: {} }; }
      })) };
    }
  };
  const context: any = { document: { body, getSelection: () => ({ text: "elsewhere", load: vi.fn(), font: {}, insertText: () => ({ font: {} }), insertOoxml: () => ({}), insertComment: () => ({}), paragraphs: { items: [selectionParagraph], load: vi.fn() } }) },
    sync: async () => { syncCount++; if (syncCount === failAt || query.length > 255) throw new Error("SECRET host debug document text"); queued.splice(0).forEach(f => f()); }
  };
  wordRun.mockImplementation(async work => { try { return await work(context); } finally { queued.splice(0); } });
  return { context, paragraphs, selectionParagraph, calls, fail: (n: number) => { failAt = n; }, syncs: () => syncCount };
}
const args = (tool: string, find: string, replace = "updated") => tool === "find_replace" ? { find, replace } : { edits: [{ find, replace }] };
const run = async (tool: string, value: any) => JSON.parse(await executeWordTool(tool, JSON.stringify(value)));

describe("bounded Word search and style repair candidate", () => {
  for (const tool of ["find_replace", "apply_edit_plan"]) {
    for (const length of [255, 256, 312]) {
      it(`${tool} ${length} UTF-16 source across all executor boundaries`, async () => {
        const source = "x".repeat(length);
        for (const entry of ["raw", "parsed", "context", "bridge"]) {
          const h = host([source]);
          const value = args(tool, source);
          const result = JSON.parse(await (entry === "raw" ? executeWordTool(tool, JSON.stringify(value)) : entry === "parsed" ? executeWordToolParsed(tool, value) : entry === "context" ? executeWordToolInContext(tool, value, h.context) : executeValidatedWordBridgeTool(tool, value)));
          if (length === 255) {
            expect(result.ok).toBe(true); expect(h.paragraphs[0].text).toBe("updated");
          } else {
            expect(result).toMatchObject({ ok: false, phase: "validation", mutationState: "none", sourceLength: length, editIndex: 0 });
            expect(h.calls).toHaveLength(0); expect(h.paragraphs[0].text).toBe(source);
          }
        }
      });
    }
    it(`${tool} encodes literal carets and sets all literal flags`, async () => {
      const h = host([" ^p (*)? "]);
      expect((await run(tool, args(tool, " ^p (*)? "))).ok).toBe(true);
      expect(h.paragraphs[0].text).toBe("updated");
      expect(h.calls).toEqual([{ query: " ^^p (*)? ", options: { matchCase: false, matchWholeWord: false, matchWildcards: false, ignorePunct: false, ignoreSpace: false, matchPrefix: false, matchSuffix: false } }]);
    });
    it(`${tool} rejects encoded overflow, astral overflow and unsupported controls`, async () => {
      for (const source of ["^".repeat(128), "😀".repeat(128), "a\tb", "a\rb", "a\nb", "a\vb", "a\u007fb", "a\u0085b", "a\u2028b", "a\u2029b"]) {
        const h = host([source]);
        expect(await run(tool, args(tool, source))).toMatchObject({ ok: false, phase: "validation", mutationState: "none" });
        expect(h.calls).toHaveLength(0); expect(h.paragraphs[0].text).toBe(source);
      }
    });
    it(`${tool} retains 5000 replacement capacity`, async () => {
      const h = host(); expect((await run(tool, args(tool, "source", "r".repeat(5000)))).ok).toBe(true);
      expect(h.paragraphs[0].text).toBe("r".repeat(5000));
    });
    for (const phase of ["search", "write"]) {
      it(`${tool} sanitizes ${phase} sync failure`, async () => {
        const h = host(); h.fail(phase === "search" ? 1 : 2);
        const result = await run(tool, args(tool, "source"));
        expect(result).toMatchObject({ ok: false, phase, mutationState: phase === "search" ? "none" : "unknown" });
        expect(JSON.stringify(result)).not.toContain("SECRET");
        expect(h.paragraphs[0].text).toBe("source");
      });
    }
  }
  it("validates the whole plan before Word entry or writes", async () => {
    const h = host();
    expect(await run("apply_edit_plan", { edits: [{ find: "source", replace: "changed" }, { find: "x".repeat(256), replace: "" }] })).toMatchObject({ ok: false, editIndex: 1, phase: "validation", mutationState: "none" });
    expect(h.calls).toHaveLength(0); expect(h.paragraphs[0].text).toBe("source");
  });
  it("reports comparison failure before writes", async () => {
    const h = host(["source", "second"]); h.fail(2);
    expect(await run("apply_edit_plan", { edits: [{ find: "source", replace: "a" }, { find: "second", replace: "b" }] })).toMatchObject({ ok: false, phase: "compare", mutationState: "none" });
    expect(h.paragraphs.map(p => p.text)).toEqual(["source", "second"]);
  });
  it("keeps ambiguous edit-plan sources untouched", async () => {
    const h = host(["source", "source"]);
    expect(await run("apply_edit_plan", args("apply_edit_plan", "source"))).toMatchObject({ totalApplied: 0 });
    expect(h.paragraphs.map(p => p.text)).toEqual(["source", "source"]);
  });
  it("canonical schemas cap direct finds and retain replacements", () => {
    for (const tool of ["find_replace", "apply_edit_plan"]) {
      const schema: any = toolSchemas.find(s => s.function.name === tool)!.function.parameters;
      const fields = tool === "find_replace" ? schema.properties : schema.properties.edits.items.properties;
      expect(fields.find.maxLength).toBe(255); expect(fields.replace.maxLength).toBe(5000);
      expect(fields.find.description).toContain("UTF-16");
    }
  });
  it("reads actual paragraph styles and untrimmed boundary text", async () => {
    const h = host([" heading ", "body"]);
    h.paragraphs[0].styleBuiltIn = "Heading1"; await h.context.sync();
    const r = await run("get_document_structure", {});
    expect(r.paragraphPreviews[0]).toMatchObject({ text: " heading ", length: 9, style: "Heading 1", styleBuiltIn: "Heading1" });
    expect(r.styleBuiltInSupported).toBe(true);
  });
  it("exposes unsupported built-in style capability without inventing it", async () => {
    host(); isSetSupported.mockReturnValue(false);
    const r = await run("get_document_structure", {});
    expect(r).toMatchObject({ styleBuiltInSupported: false });
    expect(r.paragraphPreviews[0]).toMatchObject({ style: "Normal", styleBuiltIn: null });
    expect(await run("apply_style", { style: "Heading1", targetText: "source" })).toMatchObject({ ok: false, phase: "validation", mutationState: "none" });
  });
  it("repairs exact complete paragraph independent of selection and returns observed style", async () => {
    const h = host(["Quarterly Summary", "body"]);
    const r = await run("apply_style", { style: "Heading1", targetText: "Quarterly Summary" });
    expect(r).toMatchObject({ ok: true, styledParagraphs: 1, observedStyles: [{ style: "Heading 1", styleBuiltIn: "Heading1" }] });
    expect(h.paragraphs[0].styleBuiltIn).toBe("Heading1"); expect(h.paragraphs[1].style).toBe("Normal"); expect(h.selectionParagraph.style).toBe("Normal");
  });
  it("refuses missing, partial and duplicate paragraph targets without styling selection", async () => {
    for (const texts of [["missing"], ["prefix source"], ["source", "source"]]) {
      const h = host(texts);
      expect(await run("apply_style", { style: "Heading1", targetText: "source" })).toMatchObject({ ok: false, mutationState: "none" });
      expect(h.paragraphs.every(p => p.style === "Normal")).toBe(true); expect(h.selectionParagraph.style).toBe("Normal");
    }
  });
  it("reports style readback failure as uncertain after writes", async () => {
    const h = host(); h.fail(5);
    expect(await run("apply_style", { style: "Heading1", targetText: "source" })).toMatchObject({ ok: false, phase: "readback", mutationState: "unknown" });
  });
  it("constructs heading/body with native paragraphs and observed styles without extra empty paragraphs", async () => {
    const h = host([""]);
    const r = await run("insert_text", { paragraphs: [{ text: "Quarterly Summary", styleBuiltIn: "Heading1" }, { text: "Fruit prices rose modestly.", styleBuiltIn: "Normal" }] });
    expect(r).toMatchObject({ ok: true, paragraphCount: 2, observedParagraphs: [{ text: "Quarterly Summary", styleBuiltIn: "Heading1" }, { text: "Fruit prices rose modestly.", styleBuiltIn: "Normal" }] });
    expect(h.paragraphs.map(p => p.text)).toEqual(["Quarterly Summary", "Fruit prices rose modestly."]);
  });
  it("validates all typed paragraphs before insertion", async () => {
    const h = host([""]);
    expect(await run("insert_text", { paragraphs: [{ text: "heading", styleBuiltIn: "Heading1" }, { text: "bad\nbody", styleBuiltIn: "Normal" }] })).toMatchObject({ ok: false, phase: "validation", mutationState: "none" });
    expect(h.paragraphs.map(p => p.text)).toEqual([""]);
  });
  it("local prompt includes shared recovery guidance", () => {
    const source = readFileSync(new URL("../agent.ts", import.meta.url), "utf8");
    expect(source).toContain("${WORD_EDIT_SAFETY_GUIDANCE}");
  });
  it("local completion awaits actual text/boundary/style readback after a mutation", async () => {
    const h = host();
    const fetch = vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response(JSON.stringify({ choices: [{ message: { tool_calls: [{ id: "edit", type: "function", function: { name: "find_replace", arguments: JSON.stringify({ find: "source", replace: "updated" }) } }] } }] }), { headers: { "content-type": "application/json" } }));
    const activity: any[] = [];
    try {
      const r = await new HermesAgent().run("Please change source to updated in this document", { baseUrl: "https://example.test", model: "test", maxIterations: 4 }, { onAssistantStart() {}, onAssistantDelta() {}, onAssistantFinal() {}, onStatus() {}, onActivity: a => activity.push(a) }, new AbortController().signal);
      expect(h.paragraphs[0].text).toBe("updated");
      const reads = activity.filter(a => a.label === "Post-mutation readback").map(a => JSON.parse(a.detail));
      expect(reads).toHaveLength(1);
      expect(reads[0].text.text).toBe("updated");
      expect(reads[0].structure.paragraphPreviews[0]).toMatchObject({ text: "updated", style: "Normal", styleBuiltIn: "Normal" });
      expect(r.finalContent).toContain("unverified"); // No preimage-aware general replacement proof.
    } finally { fetch.mockRestore(); }
  });
  for (const [tool, payload] of [
    ["insert_text", { paragraphs: [{ text: "heading", styleBuiltIn: "Heading1" }, { text: "body", styleBuiltIn: "Normal" }] }],
    ["apply_style", { style: "Heading1", targetText: "source" }]
  ] as const) {
    it(`local ${tool} summary uses observed paragraph/style receipts`, async () => {
      host();
      const fetch = vi.spyOn(globalThis, "fetch").mockResolvedValue(new Response(JSON.stringify({ choices: [{ message: { tool_calls: [{ id: "edit", type: "function", function: { name: tool, arguments: JSON.stringify(payload) } }] } }] }), { headers: { "content-type": "application/json" } }));
      try {
        const r = await new HermesAgent().run("Please change this document", { baseUrl: "https://example.test", model: "test", maxIterations: 4 }, { onAssistantStart() {}, onAssistantDelta() {}, onAssistantFinal() {}, onStatus() {}, onActivity() {} }, new AbortController().signal);
        expect(r.finalContent).toContain("observed");
        expect(r.finalContent).toContain(tool === "insert_text" ? "2 paragraphs" : "Heading1");
        expect(r.finalContent).not.toContain("inserted 0 chars");
      } finally { fetch.mockRestore(); }
    });
  }
  for (const [tool, payload] of [
    ["insert_text", { text: "new" }], ["replace_selection", { text: "new" }],
    ["format_selection", { bold: true }], ["insert_ooxml", { ooxml: "<w:p/>" }], ["insert_comment", { text: "note" }]
  ] as const) {
    it(`${tool} does not label a possible mutation failure none`, async () => {
      const h = host(); h.fail(1);
      expect(await run(tool, payload)).toMatchObject({ ok: false, phase: "write", mutationState: "unknown" });
    });
  }
  it("refuses a target changed between search and full paragraph comparison", async () => {
    const h = host(); const sync = h.context.sync;
    h.context.sync = async () => { await sync(); if (h.syncs() === 2) h.paragraphs[0].text = "changed by user"; };
    expect(await run("apply_style", { style: "Heading1", targetText: "source" })).toMatchObject({ ok: false, error: "paragraph_target_not_exact", mutationState: "none" });
    expect(h.paragraphs[0].style).toBe("Normal");
  });
  it("returns actual custom style rather than assuming requested style", async () => {
    const h = host();
    const r = await run("apply_style", { style: "Custom Local Style", targetText: "source" });
    expect(r).toMatchObject({ ok: true, observedStyles: [{ style: "Custom Local Style", styleBuiltIn: "Other" }] });
  });
  it("paragraph insertion appends boundaries without merging existing content", async () => {
    const h = host(["existing"]);
    expect((await run("insert_text", { paragraphs: [{ text: "heading", styleBuiltIn: "Heading1" }, { text: "body", styleBuiltIn: "Normal" }] })).ok).toBe(true);
    expect(h.paragraphs.map(p => p.text)).toEqual(["existing", "heading", "body"]);
    expect(h.paragraphs[0].style).toBe("Normal");
  });
  it("paragraph mode rejects incompatible, over-budget and unsupported requests", async () => {
    for (const payload of [
      { text: "x", paragraphs: [{ text: "p", styleBuiltIn: "Normal" }] },
      { location: "selection", paragraphs: [{ text: "p", styleBuiltIn: "Normal" }] },
      { paragraphs: [{ text: "a".repeat(3000), styleBuiltIn: "Normal" }, { text: "b".repeat(2001), styleBuiltIn: "Normal" }] }
    ]) {
      const h = host([""]);
      expect(await run("insert_text", payload)).toMatchObject({ ok: false, phase: "validation", mutationState: "none" });
      expect(h.syncs()).toBe(0); expect(h.paragraphs.map(p => p.text)).toEqual([""]);
    }
    const h = host([""]); isSetSupported.mockReturnValue(false);
    expect(await run("insert_text", { paragraphs: [{ text: "p", styleBuiltIn: "Normal" }] })).toMatchObject({ ok: false, error: "built_in_style_unsupported", mutationState: "none" });
    expect(h.syncs()).toBe(0);
  });
  async function local(tool: string, payload: unknown) {
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async () => new Response(JSON.stringify({ choices: [{ message: { tool_calls: [{ id: "amend", type: "function", function: { name: tool, arguments: JSON.stringify(payload) } }] } }] }), { headers: { "content-type": "application/json" } }));
    const activity: any[] = [];
    try {
      const result = await new HermesAgent().run("Please change this document", { baseUrl: "https://example.test", model: "test", maxIterations: 4 }, { onAssistantStart() {}, onAssistantDelta() {}, onAssistantFinal() {}, onStatus() {}, onActivity: a => activity.push(a) }, new AbortController().signal);
      return { ...result, activity, requests: fetch.mock.calls.length };
    } finally { fetch.mockRestore(); }
  }
  for (const tool of ["apply_style", "insert_text"]) {
    const payload = (target: string) => tool === "apply_style"
      ? { style: "Heading1", targetText: target }
      : { location: "end", paragraphs: [{ text: target, styleBuiltIn: "Heading1" }] };
    it(`R1 ${tool} refuses a 203-unit literal-ellipsis preview collision without replay`, async () => {
      const target = "a".repeat(200) + "...";
      const changed = "a".repeat(200) + "XYZ";
      const h = host(tool === "apply_style" ? [target] : ["existing"]);
      const body = h.context.document.body;
      const getText = Object.getOwnPropertyDescriptor(body, "text")!.get!;
      let changedBetweenReads = false;
      Object.defineProperty(body, "text", { get() {
        const value = getText.call(body);
        // Initial snapshot precedes mutation. Change only after the fresh text
        // read has captured the requested content, before structure observation.
        if (!changedBetweenReads && h.syncs() >= (tool === "apply_style" ? 7 : 6)) {
          h.paragraphs[h.paragraphs.length - 1].text = changed;
          changedBetweenReads = true;
        }
        return value;
      } });
      const r = await local(tool, payload(target));
      const readback = JSON.parse(r.activity.find(a => a.label === "Post-mutation readback").detail);
      expect(changedBetweenReads).toBe(true);
      expect(readback.text.text).toContain(target);
      expect(readback.structure.paragraphPreviews.at(-1)).toMatchObject({ text: target, length: 203, truncated: true, styleBuiltIn: "Heading1" });
      expect(h.paragraphs.at(-1).text).toBe(changed);
      expect(r.finalContent).toContain("unverified");
      expect(r.finalContent).not.toContain(`${tool} complete`);
      expect(r.requests).toBe(1);
      expect(h.calls).toHaveLength(tool === "apply_style" ? 1 : 0);
      expect(h.syncs()).toBe(tool === "apply_style" ? 9 : 8);
      if (tool === "insert_text") expect(h.paragraphs.map(p => p.text)).toEqual(["existing", changed]);
    });
    it(`R1 ${tool} refuses missing preview truncation metadata`, async () => {
      const h = host(tool === "apply_style" ? ["source"] : ["existing"]);
      const actualReadback = wordTools.getPostMutationReadback;
      const readback = vi.spyOn(wordTools, "getPostMutationReadback").mockImplementation(async () => {
        const result = await actualReadback();
        // Simulate an older/incomplete observation envelope, not a new producer.
        for (const p of (result.structure as any).paragraphPreviews) delete p.truncated;
        return result;
      });
      try {
        const r = await local(tool, payload("source"));
        expect(r.finalContent).toContain("unverified");
        expect(r.finalContent).not.toContain(`${tool} complete`);
        expect(r.requests).toBe(1);
        expect(h.calls).toHaveLength(tool === "apply_style" ? 1 : 0);
        expect(h.paragraphs.at(-1).text).toBe("source");
        expect(h.paragraphs.at(-1).styleBuiltIn).toBe("Heading1");
      } finally { readback.mockRestore(); }
    });
  }
  it("F1 does not claim replacement completion on contradictory readback or replay", async () => {
    const h = host(); const sync = h.context.sync;
    h.context.sync = async () => { await sync(); if (h.paragraphs[0].text === "updated") h.paragraphs[0].text = "source"; };
    const r = await local("find_replace", { find: "source", replace: "updated" });
    expect(r.finalContent).not.toContain("find_replace complete");
    expect(r.finalContent).toMatch(/unverified|not verified/i); expect(r.requests).toBe(1); expect(h.calls).toHaveLength(1);
  });
  it("F1 rejects observed style mismatch without compensating writes", async () => {
    const h = host(); const sync = h.context.sync;
    h.context.sync = async () => { await sync(); if (h.syncs() >= 4) { h.paragraphs[0].styleBuiltIn = "Normal"; await sync(); } };
    const r = await local("apply_style", { style: "Heading1", targetText: "source" });
    expect(r.finalContent).not.toContain("apply_style complete"); expect(r.finalContent).toMatch(/unverified|not verified/i); expect(r.requests).toBe(1);
  });
  it("F1 stops when fresh readback fails", async () => {
    const h = host(); h.fail(4);
    const r = await local("find_replace", { find: "source", replace: "updated" });
    expect(r.finalContent).toContain("readback failed"); expect(r.requests).toBe(1); expect(h.calls).toHaveLength(1);
  });
  it("F1 target outside bounded readback is not completion", async () => {
    host(["x".repeat(5100), "source"]);
    const r = await local("apply_style", { style: "Heading1", targetText: "source" });
    expect(r.finalContent).not.toContain("apply_style complete"); expect(r.finalContent).toMatch(/unverified|not verified/i); expect(r.requests).toBe(1);
  });
  for (const tool of ["find_replace", "apply_edit_plan"]) {
    it(`F3 local ${tool} has diagnostic parity without entering Word`, async () => {
      const h = host(); const source = "x".repeat(256);
      const payload = tool === "find_replace" ? args(tool, source) : { edits: [{ find: "source", replace: "a" }, { find: source, replace: "b" }] };
      const bridge = JSON.parse(await executeValidatedWordBridgeTool(tool, payload));
      const r = await local(tool, payload);
      const result = r.activity.filter(a => a.id === "amend" && a.detail).map(a => JSON.parse(a.detail))[0];
      expect(result).toEqual(bridge); expect(result).toMatchObject({ phase: "validation", mutationState: "none", sourceLength: 256, encodedLength: 256, editIndex: tool === "find_replace" ? 0 : 1 });
      expect(h.calls).toHaveLength(0); expect(h.syncs()).toBe(1); // Initial read-only context snapshot only.
    });
  }
  it("F4 local custom style displays its name rather than Other", async () => {
    host(); const r = await local("apply_style", { style: "Custom Local Style", targetText: "source" });
    expect(r.finalContent).toContain("Custom Local Style"); expect(r.finalContent).not.toContain("styles: Other");
  });
  it("F2 never replaces a text-empty paragraph containing opaque rich content", async () => {
    const h = host([""]); const p = h.paragraphs[0]; p.inlinePictures = { items: [{}] }; p.contentControls = { items: [{}] };
    expect((await run("insert_text", { paragraphs: [{ text: "new", styleBuiltIn: "Normal" }] })).ok).toBe(true);
    expect(p.insertText).not.toHaveBeenCalledWith("new", "Replace"); expect(p.inlinePictures.items).toHaveLength(1); expect(p.contentControls.items).toHaveLength(1);
  });
  for (const location of ["start", "end"]) {
    it(`typed insertion ${location} preserves existing paragraph and order`, async () => {
      const h = host(["existing"]); const p = h.paragraphs[0];
      await run("insert_text", { location, paragraphs: [{ text: "a", styleBuiltIn: "Heading1" }, { text: "b", styleBuiltIn: "Normal" }] });
      expect(h.paragraphs.map(p => p.text)).toEqual(location === "start" ? ["a", "b", "existing"] : ["existing", "a", "b"]); expect(p.style).toBe("Normal");
    });
  }
  it("typed insertion readback failure is unknown and local does not replay", async () => {
    const h = host(["existing"]); h.fail(5);
    const r = await local("insert_text", { paragraphs: [{ text: "new", styleBuiltIn: "Normal" }] });
    expect(h.paragraphs.map(p => p.text)).toEqual(["existing", "new"]);
    const receipt = r.activity.filter(a => a.id === "amend" && a.detail).map(a => JSON.parse(a.detail))[0];
    expect(receipt).toMatchObject({ ok: false, phase: "readback", mutationState: "unknown" }); expect(r.requests).toBe(1); expect(r.finalContent).toMatch(/inspect/i);
  });
  it("partial write then error stays unknown and never replays", async () => {
    const h = host(["source", "second"]); const sync = h.context.sync;
    h.context.sync = async () => { if (h.syncs() === 3) { h.paragraphs[0].text = "updated"; throw new Error("SECRET partial host write"); } await sync(); };
    const r = await local("apply_edit_plan", { edits: [{ find: "source", replace: "updated" }, { find: "second", replace: "other" }] });
    const receipt = r.activity.filter(a => a.id === "amend" && a.detail).map(a => JSON.parse(a.detail))[0];
    expect(receipt).toMatchObject({ ok: false, phase: "write", mutationState: "unknown" }); expect(h.paragraphs.map(p => p.text)).toEqual(["updated", "second"]); expect(r.requests).toBe(1); expect(r.finalContent).toMatch(/inspect/i);
  });
  it("fresh style readback contradicting a matching receipt stops local completion", async () => {
    const h = host(); const sync = h.context.sync;
    h.context.sync = async () => { await sync(); if (h.syncs() === 7) { h.paragraphs[0].styleBuiltIn = "Normal"; await sync(); } };
    const r = await local("apply_style", { style: "Heading1", targetText: "source" });
    expect(r.finalContent).toContain("unverified"); expect(r.requests).toBe(1);
  });
  it("fresh typed insertion readback contradicting its receipt stops without replay", async () => {
    const h = host(["existing"]); const sync = h.context.sync;
    h.context.sync = async () => { await sync(); if (h.syncs() === 6) h.paragraphs[1].text = "changed by user"; };
    const r = await local("insert_text", { paragraphs: [{ text: "new", styleBuiltIn: "Normal" }] });
    expect(r.finalContent).toContain("unverified"); expect(r.requests).toBe(1); expect(h.paragraphs.map(p => p.text)).toEqual(["existing", "changed by user"]);
  });
  it("server-owned recovery guidance awaits readback and stops destructive recovery", () => {
    const prompt = getServerOwnedSystemPrompt();
    expect(prompt).toContain("Await each mutation"); expect(prompt).toContain("actual text, paragraph boundaries, and styles");
    expect(prompt).toContain("Do not delete/reinsert"); expect(prompt).toContain("native Undo");
  });
});
