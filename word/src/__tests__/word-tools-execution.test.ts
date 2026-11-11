import { beforeEach, describe, expect, it, vi } from "vitest";
import { executeWordTool, executeWordToolInContext, getDocumentSnapshot, truncateText } from "../wordTools";
import { isSetSupported, mockWordContext, wordRun } from "./setup";

describe("executeWordTool Office.js paths", () => {
  it("executes a parsed tool against a supplied Word context without opening Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const raw = await executeWordToolInContext("format_selection", { bold: true }, mockWordContext as unknown as Word.RequestContext);
    const result = JSON.parse(raw);

    expect(result).toMatchObject({ ok: true, action: "format_selection", changedProperties: 1 });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.selection.font.bold).toBe(true);
  });

  it("reads document body through Word.run", async () => {
    const result = await run("read_document_text");
    expect(result).toMatchObject({ ok: true, action: "read_document_text" });
    expect(result.text).toContain("Sample paragraph");
  });

  it("rejects non-positive and fractional maxLength before calling Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const zeroResult = await run("read_document_text", { maxLength: 0 });
    const fractionalResult = await run("read_document_text", { maxLength: 1.5 });

    expect(zeroResult).toMatchObject({
      ok: false,
      error: 'Argument "maxLength" must be an integer between 1 and 20000.'
    });
    expect(fractionalResult).toMatchObject({
      ok: false,
      error: 'Argument "maxLength" must be an integer between 1 and 20000.'
    });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
  });

  it("rejects maxLength above 20000 without calling Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("read_document_text", { maxLength: 20001 });

    expect(result).toMatchObject({
      ok: false,
      error: 'Argument "maxLength" must be an integer between 1 and 20000.'
    });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
  });

  it("honors a valid integer maxLength and preserves bounded read metadata", async () => {
    const result = await run("read_document_text", { maxLength: 10 });

    expect(result).toMatchObject({
      ok: true,
      action: "read_document_text",
      truncated: true,
      returnedLength: 10
    });
  });

  it("supports read aliases", async () => {
    expect((await run("get_document_text")).ok).toBe(true);
  });

  it("normalizes tool names for executeWordTool", async () => {
    const result = await run("  GET_DOCUMENT_TEXT ");
    expect(result).toMatchObject({ ok: true, action: "read_document_text" });

    const writeResult = await run("  INSERT_TEXT ", { text: "hello" });
    expect(writeResult).toMatchObject({ ok: true, action: "insert_text", location: "end", insertedLength: 5 });

    expect(await run("  unknown_tool ")).toMatchObject({ ok: false, error: "Unknown tool:   unknown_tool " });
  });

  it("rejects empty or whitespace-only insert_text before calling Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const emptyResult = await run("insert_text", { text: "" });
    const whitespaceResult = await run("insert_text", { text: "   \n\t" });

    expect(emptyResult).toMatchObject({ ok: false, error: 'Argument "text" must not be empty.' });
    expect(whitespaceResult).toMatchObject({ ok: false, error: 'Argument "text" must not be empty.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
    expect(mockWordContext.body.insertText).not.toHaveBeenCalled();
  });

  it("replaces selection and reports previous length in a single sync", async () => {
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const result = await run("replace_selection", { text: "new" });
    expect(mockWordContext.selection.insertText).toHaveBeenCalledWith("new", "Replace");
    expect(result.previousLength).toBe(mockWordContext.selection.text.length);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore + 1);
  });

  it("rejects empty or whitespace-only replace_selection text before calling Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const emptyResult = await run("replace_selection", { text: "" });
    const whitespaceResult = await run("replace_selection", { text: "   \n\t" });

    expect(emptyResult).toMatchObject({ ok: false, error: 'Argument "text" must not be empty.' });
    expect(whitespaceResult).toMatchObject({ ok: false, error: 'Argument "text" must not be empty.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
    expect(mockWordContext.selection.insertText).not.toHaveBeenCalled();
  });

  it("rejects oversized replace_selection text before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const result = await run("replace_selection", { text: "x".repeat(5001) });

    expect(result).toMatchObject({ ok: false, error: 'Argument "text" must not exceed 5000 characters.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
    expect(mockWordContext.selection.insertText).not.toHaveBeenCalled();
  });

  it("rejects oversized insert_text before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const result = await run("insert_text", { text: "x".repeat(5001) });

    expect(result).toMatchObject({ ok: false, error: 'Argument "text" must not exceed 5000 characters.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
    expect(mockWordContext.body.insertText).not.toHaveBeenCalled();
  });

  it("rejects oversized insert_comment text before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const result = await run("insert_comment", { text: "x".repeat(5001) });

    expect(result).toMatchObject({ ok: false, error: 'Argument "text" must not exceed 5000 characters.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
    expect(mockWordContext.selection.insertComment).not.toHaveBeenCalled();
  });

  it("rejects oversized insert_ooxml before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const result = await run("insert_ooxml", { ooxml: "x".repeat(20001) });

    expect(result).toMatchObject({ ok: false, error: 'Argument "ooxml" must not exceed 20000 characters.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
    expect(mockWordContext.body.insertOoxml).not.toHaveBeenCalled();
    expect(mockWordContext.selection.insertOoxml).not.toHaveBeenCalled();
  });

  it("forwards find options and reports replacement count", async () => {
    const result = await run("find_replace", { find: "draft", replace: "final", matchCase: true });
    expect(mockWordContext.body.search).toHaveBeenCalledWith("draft", { matchCase: true, matchWholeWord: false, matchWildcards: false, ignorePunct: false, ignoreSpace: false, matchPrefix: false, matchSuffix: false });
    expect(result.matchCase).toBe(true);
    expect(result.matchWholeWord).toBe(false);
    expect(result.replacements).toBe(2);
  });

  it("rejects an invalid highlightColor for replace_selection before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("replace_selection", { text: "updated", highlightColor: "not-a-color" });

    expect(result).toMatchObject({ ok: false });
    expect(result.error).toContain("invalid highlightColor");
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.selection.insertText).not.toHaveBeenCalled();
  });

  it("rejects an empty find_replace search before calling Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("find_replace", { find: "", replace: "replacement" });

    expect(result).toMatchObject({ ok: false, error: 'Argument "find" must not be empty.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.body.search).not.toHaveBeenCalled();
  });

  it("rejects identical find and replace in find_replace before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("find_replace", { find: "same", replace: "same" });

    expect(result).toMatchObject({ ok: false, error: 'Argument "find" and "replace" must be different.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.body.search).not.toHaveBeenCalled();
  });

  it("rejects oversized find in find_replace before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("find_replace", { find: "x".repeat(5001), replace: "y" });

    expect(result).toMatchObject({ ok: false, error: "search_source_too_long", phase: "validation", mutationState: "none", sourceLength: 5001, encodedLength: 5001 });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.body.search).not.toHaveBeenCalled();
  });

  it("rejects oversized replace in find_replace before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("find_replace", { find: "x", replace: "y".repeat(5001) });

    expect(result).toMatchObject({ ok: false, error: 'Argument "replace" must not exceed 5000 characters.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.body.search).not.toHaveBeenCalled();
  });

  it("rejects an invalid highlightColor for find_replace before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("find_replace", { find: "draft", replace: "final", highlightColor: "chartreuse" });

    expect(result).toMatchObject({ ok: false });
    expect(result.error).toContain("invalid highlightColor");
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.body.search).not.toHaveBeenCalled();
  });

  it("normalizes valid hex highlightColor in find_replace results", async () => {
    const result = await run("find_replace", {
      find: "draft",
      replace: "final",
      highlightColor: "#ff00ff"
    });

    expect(result).toMatchObject({ ok: true, action: "find_replace", highlightColor: "#FF00FF" });
  });

  it("normalizes yellow highlightColor in format_selection", async () => {
    const result = await run("format_selection", { bold: true, highlightColor: "yellow" });

    expect(result).toMatchObject({ ok: true, action: "format_selection" });
    expect(mockWordContext.selection.font.highlightColor).toBe("#FFFF00");
  });

  it("does not submit a mutation sync when find_replace finds no matches", async () => {
    const noMatches = { items: [], load: vi.fn() };
    mockWordContext.body.search = vi.fn().mockReturnValue(noMatches);
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const result = await run("find_replace", { find: "missing", replace: "replacement" });

    expect(result).toMatchObject({ ok: true, action: "find_replace", replacements: 0 });
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore + 1);
  });

  it("does not submit a mutation when no formatting properties are provided", async () => {
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("format_selection");

    expect(result).toMatchObject({
      ok: false,
      error: "format_selection requires at least one formatting property."
    });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
  });

  it("rejects non-positive font sizes before Word.run and preserves fractional sizes", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const zeroResult = await run("format_selection", { fontSize: 0 });
    const negativeResult = await run("format_selection", { fontSize: -2 });

    expect(zeroResult).toMatchObject({
      ok: false,
      error: 'Argument "fontSize" must be at least 1.'
    });
    expect(negativeResult).toMatchObject({
      ok: false,
      error: 'Argument "fontSize" must be at least 1.'
    });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);

    const validResult = await run("format_selection", { fontSize: 12.5 });

    expect(validResult).toMatchObject({ ok: true, action: "format_selection", changedProperties: 1 });
    expect(mockWordContext.selection.font.size).toBe(12.5);
  });

  it("sets only supplied formatting properties", async () => {
    const result = await run("format_selection", { italic: true });
    expect(mockWordContext.selection.font.italic).toBe(true);
    expect(mockWordContext.selection.font.bold).toBeUndefined();
    expect(result.changedProperties).toBe(1);
  });

  it("rejects an invalid highlightColor for format_selection before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("format_selection", { bold: true, highlightColor: "chartreuse" });

    expect(result).toMatchObject({ ok: false });
    expect(result.error).toContain("invalid highlightColor");
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.selection.font.bold).toBeUndefined();
    expect(mockWordContext.selection.font.highlightColor).toBeUndefined();
  });

  it("normalizes valid highlightColor in format_selection before applying font", async () => {
    const result = await run("format_selection", { italic: true, highlightColor: "#ff00ff" });

    expect(result).toMatchObject({ ok: true, action: "format_selection" });
    expect(mockWordContext.selection.font.highlightColor).toBe("#FF00FF");
  });

  it("rejects an invalid location for insert_text before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("insert_text", { text: "hello", location: "invalid" });

    expect(result).toMatchObject({ ok: false });
    expect(result.error).toContain("invalid location");
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.body.insertText).not.toHaveBeenCalled();
  });

  it("inserts at document end by default", async () => {
    await run("insert_text", { text: "hello" });
    expect(mockWordContext.body.insertText).toHaveBeenCalledWith("hello", "End");
  });

  it("inserts at document start", async () => {
    await run("insert_text", { text: "hello", location: "start" });
    expect(mockWordContext.body.insertText).toHaveBeenCalledWith("hello", "Start");
  });

  it("returns structured errors for unknown tools", async () => {
    expect(await run("nonexistent_tool")).toMatchObject({ ok: false });
  });

  it("returns structured errors when Word.run rejects", async () => {
    wordRun.mockRejectedValueOnce(new Error("Office unavailable"));
    expect(await run("read_document_text")).toMatchObject({ ok: false, error: "word_operation_failed", phase: "read", mutationState: "none" });
  });
});

describe("planned Word parity behaviors", () => {
  it("forwards matchWholeWord to find_replace", async () => {
    const result = await run("find_replace", { find: "draft", replace: "final", matchWholeWord: true });
    expect(result.ok).toBe(true);
    expect(mockWordContext.body.search).toHaveBeenCalledWith("draft", { matchCase: false, matchWholeWord: true, matchWildcards: false, ignorePunct: false, ignoreSpace: false, matchPrefix: false, matchSuffix: false });
  });

  it("counts paragraphs via the Word API", async () => {
    mockWordContext.body.text = "one line without separators";
    const result = await run("get_document_overview");
    expect(result.paragraphCount).toBe(3);
    expect(mockWordContext.body.paragraphs.load).toHaveBeenCalledWith("items");
  });

  it("returns bounded paragraph structure previews", async () => {
    const result = await run("get_document_structure", { maxParagraphs: 2 });
    expect(result).toMatchObject({ ok: true, totalParagraphs: 3 });
    expect(result.paragraphPreviews).toHaveLength(2);
    expect(result.paragraphPreviews[0].text).toBe("Sample paragraph.");
  });

  it("rejects non-positive, fractional, and over-limit paragraph previews before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;

    const zeroResult = await run("get_document_structure", { maxParagraphs: 0 });
    const fractionalResult = await run("get_document_structure", { maxParagraphs: 1.5 });
    const overLimitResult = await run("get_document_structure", { maxParagraphs: 101 });

    expect(zeroResult).toMatchObject({
      ok: false,
      error: 'Argument "maxParagraphs" must be an integer between 1 and 100.'
    });
    expect(fractionalResult).toMatchObject({
      ok: false,
      error: 'Argument "maxParagraphs" must be an integer between 1 and 100.'
    });
    expect(overLimitResult).toMatchObject({
      ok: false,
      error: 'Argument "maxParagraphs" must be an integer between 1 and 100.'
    });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);

    const validResult = await run("get_document_structure", { maxParagraphs: 2 });
    expect(validResult).toMatchObject({ ok: true, totalParagraphs: 3 });
    expect(validResult.paragraphPreviews).toHaveLength(2);
  });

  it("avoids a second sync when the document has no paragraph previews", async () => {
    mockWordContext.body.paragraphs.items = [];
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const result = await run("get_document_structure", { maxParagraphs: 10 });

    expect(result).toMatchObject({ ok: true, totalParagraphs: 0, paragraphPreviews: [] });
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore + 1);
  });

  it("applies a named style to every selected paragraph", async () => {
    const result = await run("apply_style", { style: "Heading 1" });
    expect(result).toMatchObject({ ok: true, styledParagraphs: 3 });
    expect(mockWordContext.selection.paragraphs.items.every((p) => p.style === "Heading 1")).toBe(true);
  });

  it("avoids a second sync when the selection has no paragraphs to style", async () => {
    mockWordContext.selection.paragraphs.items = [];
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const result = await run("apply_style", { style: "Heading 1" });

    expect(result).toMatchObject({ ok: true, styledParagraphs: 0 });
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore + 1);
  });

  it("rejects empty or whitespace-only styles before calling Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const emptyResult = await run("apply_style", { style: "" });
    const whitespaceResult = await run("apply_style", { style: "   \n\t" });

    expect(emptyResult).toMatchObject({ ok: false, error: 'Argument "style" must not be empty.' });
    expect(whitespaceResult).toMatchObject({ ok: false, error: 'Argument "style" must not be empty.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
    expect(mockWordContext.selection.paragraphs.items.every((paragraph) => paragraph.style === "Normal")).toBe(true);
  });

  it("inserts OOXML at the selected range", async () => {
    const result = await run("insert_ooxml", { ooxml: "<w:p/>", location: "selection" });
    expect(result.ok).toBe(true);
    expect(mockWordContext.selection.insertOoxml).toHaveBeenCalledWith("<w:p/>", "Replace");
  });

  it("inserts OOXML at the document end by default", async () => {
    await run("insert_ooxml", { ooxml: "<w:p/>" });
    expect(mockWordContext.body.insertOoxml).toHaveBeenCalledWith("<w:p/>", "End");
  });

  it("rejects empty or whitespace-only OOXML before calling Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const emptyResult = await run("insert_ooxml", { ooxml: "" });
    const whitespaceResult = await run("insert_ooxml", { ooxml: "   \n\t" });

    expect(emptyResult).toMatchObject({ ok: false, error: 'Argument "ooxml" must not be empty.' });
    expect(whitespaceResult).toMatchObject({ ok: false, error: 'Argument "ooxml" must not be empty.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
    expect(mockWordContext.body.insertOoxml).not.toHaveBeenCalled();
    expect(mockWordContext.selection.insertOoxml).not.toHaveBeenCalled();
  });

  it("rejects an invalid location for insert_ooxml before Word.run", async () => {
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;
    const runCallsBefore = wordRun.mock.calls.length;

    const result = await run("insert_ooxml", { ooxml: "<w:p/>", location: "outside" });

    expect(result).toMatchObject({ ok: false });
    expect(result.error).toContain("invalid location");
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.body.insertOoxml).not.toHaveBeenCalled();
    expect(mockWordContext.selection.insertOoxml).not.toHaveBeenCalled();
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
  });

  it("inserts a comment on the current selection", async () => {
    const result = await run("insert_comment", { text: "Review this." });
    expect(result.ok).toBe(true);
    expect(result).toMatchObject({ ok: true, action: "insert_comment" });
    expect(mockWordContext.selection.insertComment).toHaveBeenCalledWith("Review this.");
  });

  it("rejects empty or whitespace-only comment text before Word.run", async () => {
    const runCallsBefore = wordRun.mock.calls.length;
    const syncCallsBefore = mockWordContext.sync.mock.calls.length;

    const emptyResult = await run("insert_comment", { text: "" });
    const whitespaceResult = await run("insert_comment", { text: "   " });

    expect(emptyResult).toMatchObject({ ok: false, error: 'Argument "text" must not be empty.' });
    expect(whitespaceResult).toMatchObject({ ok: false, error: 'Argument "text" must not be empty.' });
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore);
    expect(mockWordContext.sync).toHaveBeenCalledTimes(syncCallsBefore);
    expect(mockWordContext.selection.insertComment).not.toHaveBeenCalled();
  });

  it("rejects insert_comment deterministically when WordApi 1.4 is unavailable", async () => {
    isSetSupported.mockReturnValue(false);

    const result = await run("insert_comment", { text: "Review this." });

    expect(result).toEqual({
      ok: false,
      error: "insert_comment requires WordApi 1.4, which this Word host does not support.",
      phase: "validation", mutationState: "none",
      category: "office_host"
    });
    expect(isSetSupported).toHaveBeenCalledWith("WordApi", "1.4");
    expect(wordRun).not.toHaveBeenCalled();
    expect(mockWordContext.selection.insertComment).not.toHaveBeenCalled();
  });
});

describe("document snapshot", () => {
  it("loads document body, paragraphs, and selection through one Word.run call", async () => {
    const runCountBefore = wordRun.mock.calls.length;

    const snapshot = await getDocumentSnapshot(150);

    expect(wordRun).toHaveBeenCalledTimes(runCountBefore + 1);
    expect(mockWordContext.body.load).toHaveBeenCalledWith("text");
    expect(mockWordContext.body.paragraphs.load).toHaveBeenCalledWith("items");
    expect(mockWordContext.selection.load).toHaveBeenCalledWith(["text", "start", "end"]);
    expect(snapshot.overview.charCount).toBe(mockWordContext.body.text.length);
    expect(snapshot.overview.selection).toHaveProperty("text");
  });

  it("loads selection text but not desktop offsets on unsupported Word hosts", async () => {
    isSetSupported.mockImplementation((set: string, version: string) =>
      !(set === "WordApiDesktop" && version === "1.4"));

    const snapshot = await getDocumentSnapshot();

    expect(isSetSupported).toHaveBeenCalledWith("WordApiDesktop", "1.4");
    expect(mockWordContext.selection.load).toHaveBeenCalledWith("text");
    expect(mockWordContext.selection.load).not.toHaveBeenCalledWith(["text", "start", "end"]);
    expect(snapshot.overview.selection).toMatchObject({ text: "selected content here", start: null, end: null });
  });

  it("uses desktop offsets only when WordApiDesktop 1.4 is supported", async () => {
    isSetSupported.mockImplementation((set: string, version: string) =>
      set === "WordApiDesktop" && version === "1.4");

    const snapshot = await getDocumentSnapshot();

    expect(isSetSupported).toHaveBeenCalledWith("WordApiDesktop", "1.4");
    expect(mockWordContext.selection.load).toHaveBeenCalledWith(["text", "start", "end"]);
    expect(snapshot.overview.selection).toMatchObject({ start: 0, end: 21 });
  });

  it("degrades overview and read_selection offsets independently of WordApi 1.4", async () => {
    isSetSupported.mockImplementation((set: string, version: string) =>
      !(set === "WordApiDesktop" && version === "1.4"));

    const overview = await run("get_document_overview");
    const selection = await run("read_selection");

    expect(mockWordContext.selection.load).toHaveBeenCalledWith("text");
    expect(mockWordContext.selection.load).not.toHaveBeenCalledWith(["text", "start", "end"]);
    expect(overview.selection).toMatchObject({ text: "selected content here", start: null, end: null });
    expect(selection).toMatchObject({ text: "selected content here", start: null, end: null });
    expect(typeof overview.selection.text).toBe("string");
    expect(typeof selection.text).toBe("string");
  });

  it("handles empty document text in snapshot", async () => {
    mockWordContext.body.text = "";
    const snapshot = await getDocumentSnapshot(20);

    expect(snapshot.documentText).toMatchObject({
      text: "",
      truncated: false,
      originalLength: 0,
      returnedLength: 0
    });
  });

  it("truncates long document text in snapshot and keeps metadata", async () => {
    mockWordContext.body.text = "A".repeat(30);
    const snapshot = await getDocumentSnapshot(10);

    expect(snapshot.documentText.text).toBe("A".repeat(10));
    expect(snapshot.documentText.truncated).toBe(true);
    expect(snapshot.documentText.originalLength).toBe(30);
    expect(snapshot.documentText.returnedLength).toBe(10);
    expect(snapshot.overview.selection).toMatchObject({
      start: mockWordContext.selection.start,
      end: mockWordContext.selection.end
    });
    expect(snapshot.overview.selection.text).toBe(truncateText(mockWordContext.selection.text, 300).text);
  });

  it("throws a useful error when the Word snapshot cannot be read", async () => {
    wordRun.mockRejectedValueOnce(new Error("Office unavailable"));
    await expect(getDocumentSnapshot(100)).rejects.toThrow("word_operation_failed");
  });
});

describe("apply_edit_plan execution", () => {
  function makeMatch(count: number, relation = "After") {
    return {
      insertText: vi.fn().mockReturnValue({
        font: { highlightColor: undefined as string | undefined }
      }),
      load: vi.fn(),
      items: Array.from({ length: count }, () => ({
        insertText: vi.fn().mockReturnValue({
          font: { highlightColor: undefined as string | undefined }
        }),
        load: vi.fn(),
        compareLocationWith: vi.fn().mockReturnValue({ value: relation })
      } as { insertText: ReturnType<typeof vi.fn>; load: ReturnType<typeof vi.fn>; compareLocationWith: ReturnType<typeof vi.fn> }))
    };
  }

  it("applies only unique matches in one batch and reports outcomes", async () => {
    const firstMatch = makeMatch(1);
    const secondMatch = makeMatch(1);
    mockWordContext.body.search = vi.fn()
      .mockReturnValueOnce(firstMatch as unknown as { items: unknown[]; load: () => void })
      .mockReturnValueOnce(secondMatch as unknown as { items: unknown[]; load: () => void });

    const result = await run("apply_edit_plan", {
      edits: [
        { find: "hello", replace: "world", matchCase: false, matchWholeWord: false },
        { find: "greeting", replace: "hi", highlightColor: "yellow" }
      ]
    });

    expect(result.ok).toBe(true);
    expect(result.action).toBe("apply_edit_plan");
    expect(result.totalRequested).toBe(2);
    expect(result.totalApplied).toBe(2);
    expect(result.totalSkipped).toBe(0);
    expect(result.edits).toHaveLength(2);
    expect(result.edits[0].status).toBe("applied");
    expect(mockWordContext.body.search).toHaveBeenCalledWith("hello", { matchCase: false, matchWholeWord: false, matchWildcards: false, ignorePunct: false, ignoreSpace: false, matchPrefix: false, matchSuffix: false });
    expect(mockWordContext.body.search).toHaveBeenCalledWith("greeting", { matchCase: false, matchWholeWord: false, matchWildcards: false, ignorePunct: false, ignoreSpace: false, matchPrefix: false, matchSuffix: false });
  });

  it("preserves leading and trailing spaces in exact edit-plan source matches", async () => {
    const spacedMatch = makeMatch(1);
    mockWordContext.body.search = vi.fn().mockReturnValue(spacedMatch as unknown as { items: unknown[]; load: () => void });

    const result = await run("apply_edit_plan", {
      edits: [{ find: " draft ", replace: "final" }]
    });

    expect(mockWordContext.body.search).toHaveBeenCalledWith(" draft ", { matchCase: false, matchWholeWord: false, matchWildcards: false, ignorePunct: false, ignoreSpace: false, matchPrefix: false, matchSuffix: false });
    expect(result.edits[0]).toMatchObject({ find: " draft ", replace: "final", status: "applied" });
  });

  it("rejects a batch whose distinct unique matches overlap before mutating Word ranges", async () => {
    const firstMatch = makeMatch(1, "OverlapsBefore");
    const overlappingMatch = makeMatch(1);
    mockWordContext.body.search = vi.fn()
      .mockReturnValueOnce(firstMatch as unknown as { items: unknown[]; load: () => void })
      .mockReturnValueOnce(overlappingMatch as unknown as { items: unknown[]; load: () => void });

    const result = await run("apply_edit_plan", {
      edits: [
        { find: "hello", replace: "world" },
        { find: "lowo", replace: "text" }
      ]
    });

    expect(result).toMatchObject({ ok: true, totalRequested: 2, totalApplied: 0, totalSkipped: 2 });
    expect(result.edits.map((edit: { status: string }) => edit.status)).toEqual(["overlapping", "overlapping"]);
    expect(firstMatch.items[0].insertText).not.toHaveBeenCalled();
    expect(overlappingMatch.items[0].insertText).not.toHaveBeenCalled();
  });

  it("skips edits with zero or multiple matches and does not mutate those ranges", async () => {
    const missing = { items: [], load: vi.fn() };
    const ambiguous = { items: [{}, {}] as { insertText: ReturnType<typeof vi.fn>; load?: ReturnType<typeof vi.fn> }[], load: vi.fn() };

    mockWordContext.body.search = vi.fn()
      .mockReturnValueOnce(missing as unknown as { items: unknown[]; load: () => void })
      .mockReturnValueOnce(ambiguous as unknown as { items: unknown[]; load: () => void });

    const result = await run("apply_edit_plan", {
      edits: [
        { find: "missing", replace: "x" },
        { find: "ambiguous", replace: "y" }
      ]
    });

    expect(result.totalApplied).toBe(0);
    expect(result.totalSkipped).toBe(2);
    expect(result.edits[0]).toMatchObject({ status: "not_found", matchCount: 0 });
    expect(result.edits[1]).toMatchObject({ status: "ambiguous", matchCount: 2 });
  });

  it("rejects apply_edit_plan deterministically when WordApi 1.3 is unavailable", async () => {
    isSetSupported.mockReturnValue(false);

    const callsBefore = wordRun.mock.calls.length;
    const result = await run("apply_edit_plan", {
      edits: [{ find: "draft", replace: "final" }]
    });

    expect(result).toEqual({
      ok: false,
      error: "apply_edit_plan requires WordApi 1.3, which this Word host does not support.",
      phase: "validation", mutationState: "none",
      category: "office_host"
    });
    expect(isSetSupported).toHaveBeenCalledWith("WordApi", "1.3");
    expect(wordRun).toHaveBeenCalledTimes(callsBefore);
    expect(mockWordContext.body.search).not.toHaveBeenCalled();
  });

  it("rejects duplicate source matches before calling Word.run", async () => {
    const callsBefore = wordRun.mock.calls.length;
    const result = await run("apply_edit_plan", {
      edits: [
        { find: "draft", replace: "final" },
        { find: "draft", replace: "approved" }
      ]
    });

    expect(result).toEqual({
      ok: false,
      error: "Edit 1 duplicates edit 0 for the same search text and match options.",
      phase: "validation", mutationState: "none",
      category: "validation"
    });
    expect(wordRun).toHaveBeenCalledTimes(callsBefore);
  });

  it("returns validator errors without calling Word.run", async () => {
    const callsBefore = wordRun.mock.calls.length;
    const result = await run("apply_edit_plan", { edits: [] });
    expect(result.ok).toBe(false);
    expect(wordRun).toHaveBeenCalledTimes(callsBefore);
  });
});

async function run(name: string, args: Record<string, unknown> = {}) {
  return JSON.parse(await executeWordTool(name, JSON.stringify(args)));
}
