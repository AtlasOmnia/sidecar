import { describe, expect, it } from "vitest";
import { toolAliases, toolSchemas, truncateText } from "../wordTools";

describe("word tool schemas", () => {
  it("registers the required document tool names", () => {
    const names = toolSchemas.map((tool) => tool.function.name);
    expect(names).toEqual(expect.arrayContaining([
      "get_document_overview",
      "read_selection",
      "read_document_text",
      "insert_text",
      "replace_selection",
      "find_replace",
      "format_selection",
      "get_document_structure",
      "apply_style",
      "insert_ooxml",
      "insert_comment",
      "apply_edit_plan"
    ]));
  });

  it("publishes schemas with stable tool arguments", () => {
    const schemaMap = new Map(toolSchemas.map((tool) => [tool.function.name, tool.function.parameters]));

    const findSchema = schemaMap.get("find_replace");
    expect(findSchema).toBeDefined();
    const properties = findSchema?.properties as Record<string, unknown> | undefined;
    expect(properties).toHaveProperty("find");
    expect(properties?.find).toMatchObject({ type: "string", minLength: 1 });

    const insertSchema = schemaMap.get("insert_text");
    const insertProperties = insertSchema?.properties as Record<string, unknown> | undefined;
    expect(insertProperties?.text).toMatchObject({ type: "string", minLength: 1 });
    expect(properties).toHaveProperty("replace");
    expect(properties).toHaveProperty("highlightColor");
    expect(properties).toHaveProperty("matchWholeWord");
    expect(findSchema).toBeDefined();
    expect(findSchema?.required).toEqual(["find", "replace"]);

    const readSchema = schemaMap.get("read_document_text");
    const readProperties = readSchema?.properties as Record<string, unknown> | undefined;
    expect(readProperties?.maxLength).toMatchObject({ type: "integer", minimum: 1, maximum: 20000 });

    const commentSchema = schemaMap.get("insert_comment");
    const commentProperties = commentSchema?.properties as Record<string, unknown> | undefined;
    expect(commentProperties?.text).toMatchObject({ type: "string", minLength: 1 });

    const structureSchema = schemaMap.get("get_document_structure");
    const structureProperties = structureSchema?.properties as Record<string, unknown> | undefined;
    expect(structureProperties?.maxParagraphs).toMatchObject({ type: "integer", minimum: 1, maximum: 100 });

    const styleSchema = schemaMap.get("apply_style");
    const styleProperties = styleSchema?.properties as Record<string, unknown> | undefined;
    expect(styleProperties?.style).toMatchObject({ type: "string", minLength: 1 });

    const ooxmlSchema = schemaMap.get("insert_ooxml");
    const ooxmlProperties = ooxmlSchema?.properties as Record<string, unknown> | undefined;
    expect(ooxmlProperties?.ooxml).toMatchObject({ type: "string", minLength: 1 });
  });

  it("defines apply_edit_plan schema with bounded unique, non-overlapping edits", () => {
    const schemaMap = new Map(toolSchemas.map((tool) => [tool.function.name, tool.function.parameters]));
    const applyPlanTool = toolSchemas.find((tool) => tool.function.name === "apply_edit_plan");
    const applyPlanSchema = schemaMap.get("apply_edit_plan");
    const properties = applyPlanSchema?.properties as Record<string, unknown> | undefined;

    expect(applyPlanTool?.function.description).toContain("non-overlapping");
    expect(applyPlanTool?.function.description).toContain("source search must be unique");
    expect(applyPlanSchema).toEqual(
      expect.objectContaining({
        type: "object",
        required: ["edits"],
        additionalProperties: false
      })
    );
    expect(properties?.edits).toHaveProperty("type", "array");
    expect(applyPlanSchema).toEqual(expect.objectContaining({ required: ["edits"] }));
  });

  it("requires at least one property for format_selection", () => {
    const schemaMap = new Map(toolSchemas.map((tool) => [tool.function.name, tool.function.parameters]));
    expect(schemaMap.get("format_selection")?.minProperties).toBe(1);
  });

  it("allows replacement tools to highlight changed text", () => {
    const schemaMap = new Map(toolSchemas.map((tool) => [tool.function.name, tool.function.parameters]));
    const findProperties = schemaMap.get("find_replace")?.properties as Record<string, unknown> | undefined;
    const selectionProperties = schemaMap.get("replace_selection")?.properties as Record<string, unknown> | undefined;

    expect(findProperties).toHaveProperty("highlightColor");
    expect(selectionProperties).toHaveProperty("highlightColor");

    const selectionText = selectionProperties?.text as Record<string, unknown> | undefined;
    expect(selectionText).toMatchObject({ type: "string", minLength: 1 });
  });

  it("keeps compatibility aliases for stale document read tool names", () => {
    expect(toolAliases.get_document_text).toBe("read_document_text");
    expect(toolAliases.read_document_range).toBe("read_document_text");
  });
});

describe("text truncation helper", () => {
  it("truncates long reads and reports when truncated", () => {
    const result = truncateText("ABCDEFGHIJKLMNOPQRSTUVWXYZ", 10);
    expect(result.text).toBe("ABCDEFGHIJ");
    expect(result.truncated).toBe(true);
    expect(result.returnedLength).toBe(10);
    expect(result.originalLength).toBe(26);
  });

  it("preserves short text unchanged", () => {
    const result = truncateText("hello", 10);
    expect(result.text).toBe("hello");
    expect(result.truncated).toBe(false);
    expect(result.returnedLength).toBe(5);
    expect(result.originalLength).toBe(5);
  });
});
