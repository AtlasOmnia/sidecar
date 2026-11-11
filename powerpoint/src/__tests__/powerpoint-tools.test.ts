import { describe, expect, it } from "vitest";
import {
  buildSlideBodyText,
  defaultTextBoxOptions,
  executePowerPointTool,
  notSupportedResult,
  truncateText,
  toolSchemas,
  checkRequirement,
  getCapabilities,
  replaceLiteralAll,
  validateHtmlColor,
  setRequirementOverride,
} from "../powerPointTools";

describe("PowerPoint tool schemas", () => {
  it("registers the required presentation tool names", () => {
    const names = toolSchemas.map((tool) => tool.function.name);
    expect(names).toEqual(expect.arrayContaining([
      "get_presentation_overview",
      "read_selected_slides",
      "read_all_slides",
      "create_slide",
      "create_text_slide",
      "add_text_box",
      "update_slide_text",
      "add_speaker_notes",
      "add_shape",
      "find_and_replace_text",
      "list_slide_layouts",
      "set_slide_layout",
      "format_text",
      "add_image"
    ]));
  });

  it("publishes stable schema arguments for create/update operations", () => {
    const schemaMap = new Map(toolSchemas.map((tool) => [tool.function.name, tool.function.parameters]));

    const addTextBoxSchema = schemaMap.get("add_text_box");
    expect(addTextBoxSchema).toBeDefined();
    const addTextBoxProperties = addTextBoxSchema?.properties as Record<string, unknown> | undefined;
    expect(addTextBoxProperties).toHaveProperty("text");
    expect(addTextBoxProperties).toHaveProperty("slideIndex");
    expect(addTextBoxProperties?.text).toBeTypeOf("object");
    expect(addTextBoxSchema?.required).toEqual(["text"]);

    const createTextSlideSchema = schemaMap.get("create_text_slide");
    expect(createTextSlideSchema).toBeDefined();
    const createTextSlideProperties = createTextSlideSchema?.properties as Record<string, unknown> | undefined;
    expect(createTextSlideProperties).toHaveProperty("title");
    expect(createTextSlideProperties).toHaveProperty("bullets");
    expect(createTextSlideSchema?.required).toEqual(["title"]);

    const addShapeSchema = schemaMap.get("add_shape");
    expect(addShapeSchema).toBeDefined();
    const addShapeProperties = addShapeSchema?.properties as Record<string, unknown> | undefined;
    expect(addShapeProperties).toHaveProperty("shapeType");
    expect(addShapeProperties?.shapeType).toBeTypeOf("object");
    expect(addShapeSchema?.required).toEqual(["shapeType"]);

    const updateSlideTextSchema = schemaMap.get("update_slide_text");
    expect(updateSlideTextSchema?.required).toEqual(["text"]);
  });

  it("marks add_speaker_notes as supported-by-result shape by schema", () => {
    const schemaMap = new Map(toolSchemas.map((tool) => [tool.function.name, tool.function.parameters]));
    const notesSchema = schemaMap.get("add_speaker_notes");
    expect(notesSchema).toBeDefined();
    const properties = notesSchema?.properties as Record<string, unknown> | undefined;
    expect(properties).toHaveProperty("text");
    expect(notesSchema?.required).toEqual(["text"]);
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

describe("safe slide layout helpers", () => {
  it("adds default body placement when add_text_box has no coordinates", () => {
    expect(defaultTextBoxOptions({})).toEqual({ left: 70, top: 125, width: 820, height: 330 });
  });

  it("preserves explicit placement when provided", () => {
    expect(defaultTextBoxOptions({ left: 10, top: 20, width: 300, height: 90 })).toEqual({
      left: 10,
      top: 20,
      width: 300,
      height: 90
    });
  });

  it("formats text-slide body from body plus bullets", () => {
    expect(buildSlideBodyText("Intro", ["One", "Two"])).toBe("Intro\n\n• One\n• Two");
    expect(buildSlideBodyText(undefined, ["One"])).toBe("• One");
  });
});

describe("not-supported helper", () => {
  it("returns the explicit PowerPoint not-supported result shape", () => {
    const raw = notSupportedResult("add_speaker_notes");
    const parsed = JSON.parse(raw) as Record<string, unknown>;
    expect(parsed.ok).toBe(false);
    expect(parsed.action).toBe("add_speaker_notes");
    expect(parsed.error).toBe("Not supported by current PowerPoint Office.js API");
  });
});

describe("safelyPowerPoint contract", () => {
  it("executePowerPointTool returns {ok:false, error:string} for unknown tool", async () => {
    const result = await executePowerPointTool("nonexistent_tool", "{}");
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain("Unknown tool");
  });

  it("executePowerPointTool handles malformed JSON args without crashing", async () => {
    const result = await executePowerPointTool("get_presentation_overview", "not json");
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
  });

  it("notSupportedResult always produces the correct error shape", () => {
    const raw = notSupportedResult("test_op", { slideIndex: 0 });
    const parsed = JSON.parse(raw);
    expect(parsed).toEqual({
      ok: false,
      action: "test_op",
      error: "Not supported by current PowerPoint Office.js API",
      category: "office_host",
      details: { slideIndex: 0 }
    });
  });

  it("executePowerPointTool returns stringified error for a tool with bad arg types", async () => {
    // create_text_slide requires a string 'title' arg; pass a number to trigger the type guard
    const result = await executePowerPointTool("create_text_slide", '{"title": 123}');
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toMatch(/string/i);
  });
});

describe("read_slide_content", () => {
  it("has a schema registered with required slideIndex arg", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("read_slide_content");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("slideIndex");
    expect(props?.slideIndex).toBeTypeOf("object");
    expect(schema?.required).toContain("slideIndex");
  });

  it("has schema with optional maxLength parameter", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("read_slide_content");
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("maxLength");
  });

  it("returns ok false error when passed invalid args", async () => {
    const result = await executePowerPointTool("read_slide_content", "{}");
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toMatch(/required|string|invalid/i);
  });
});

describe("delete_slide", () => {
  it("has a schema registered with optional slideIndex arg", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("delete_slide");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("slideIndex");
    expect(props?.slideIndex).toBeTypeOf("object");
    expect(schema?.required).toBeUndefined();
  });

  it("has a description mentioning last-slide guard behavior", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function]));
    const schema = schemaMap.get("delete_slide");
    expect(schema?.description.toLowerCase()).toContain("last remaining");
    expect(schema?.description.toLowerCase()).toContain("guard");
  });
});

describe("duplicate_slide", () => {
  it("has a schema registered with required slideIndex arg", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("duplicate_slide");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("slideIndex");
    expect(schema?.required).toContain("slideIndex");
  });

  it("returns not-supported result with required-api and reason details", async () => {
    const result = await executePowerPointTool("duplicate_slide", '{"slideIndex": 0}');
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.action).toBe("duplicate_slide");
    expect(parsed.error).toContain("duplicate");
    expect(parsed.details).toBeDefined();
    expect(parsed.details.requiredApi).toBeDefined();
    expect(parsed.details.reason).toBeDefined();
  });
});

describe("reorder_slide", () => {
  it("has a schema registered with required fromIndex and toIndex args", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("reorder_slide");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("fromIndex");
    expect(props).toHaveProperty("toIndex");
    expect(schema?.required).toContain("fromIndex");
    expect(schema?.required).toContain("toIndex");
  });

  it("returns ok false error when missing required arg", async () => {
    const result = await executePowerPointTool("reorder_slide", "{}");
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toMatch(/required|string|invalid/i);
  });
});

describe("checkRequirement", () => {
  it("detects requirement set names from the schema", () => {
    const result = checkRequirement("PowerPointApi", "1.2");
    expect(result).toHaveProperty("name");
    expect(result).toHaveProperty("minVersion");
    expect(result).toHaveProperty("available");
    expect(typeof result.available).toBe("boolean");
  });

  it("returns available=true for stable PowerPointApi 1.2", () => {
    const result = checkRequirement("PowerPointApi", "1.2");
    expect(result.minVersion).toBe("1.2");
    expect(result.available).toBe(true);
  });

  it("reports unknown requirement sets as available=false with clear reason", () => {
    const result = checkRequirement("PowerPointApi", "9.9");
    expect(result.available).toBe(false);
    expect(result.reason).toBeDefined();
  });
});

describe("find_and_replace_text", () => {
  it("treats regex metacharacters and replacement tokens as literal text", () => {
    expect(replaceLiteralAll("a.b a*b [x] a.b", "a.b", "X")).toBe("X a*b [x] X");
    expect(replaceLiteralAll("$1 $1", "$1", "cash")).toBe("cash cash");
  });

  it("has a schema registered with required findText and replaceText args", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("find_and_replace_text");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("findText");
    expect(props).toHaveProperty("replaceText");
    expect(schema?.required).toContain("findText");
    expect(schema?.required).toContain("replaceText");
  });

  it("has optional slideIndex and shapeId parameters", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("find_and_replace_text");
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("slideIndex");
    expect(props).toHaveProperty("shapeId");
  });

  it("returns ok false error when missing required args", async () => {
    const result = await executePowerPointTool("find_and_replace_text", "{}");
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toMatch(/required|string|invalid/i);
  });

  it("returns ok false error when findText is empty string", async () => {
    const result = await executePowerPointTool("find_and_replace_text", '{"findText": "", "replaceText": "bar"}');
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain("empty");
  });
});

describe("list_slide_layouts", () => {
  it("has a schema registered", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("list_slide_layouts");
    expect(schema).toBeDefined();
  });

  it("has no required parameters", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("list_slide_layouts");
    expect(schema?.required).toBeUndefined();
  });
});

describe("set_slide_layout", () => {
  it("has a schema registered", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("set_slide_layout");
    expect(schema).toBeDefined();
  });

  it("has required slideIndex and layoutId parameters", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("set_slide_layout");
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("slideIndex");
    expect(props).toHaveProperty("layoutId");
    expect(schema?.required).toContain("slideIndex");
    expect(schema?.required).toContain("layoutId");
  });

  it("returns ok false error when missing required args", async () => {
    const result = await executePowerPointTool("set_slide_layout", "{}");
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toMatch(/required|string|invalid/i);
  });
});

describe("format_text", () => {
  it("accepts supported HTML colors and rejects malformed colors", () => {
    expect(validateHtmlColor("#FF0000")).toBe(true);
    expect(validateHtmlColor("#abc")).toBe(true);
    expect(validateHtmlColor("red")).toBe(false);
    expect(validateHtmlColor("#12GG00")).toBe(false);
    expect(validateHtmlColor("#12345")).toBe(false);
  });

  it("rejects an invalid font color before entering Office.js", async () => {
    const result = await executePowerPointTool("format_text", '{"slideIndex":0,"shapeId":"shape-1","font":{"color":"red"}}');
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain("#RRGGBB");
  });

  it("has a schema registered with required slideIndex, shapeId", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("format_text");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("slideIndex");
    expect(props).toHaveProperty("shapeId");
    expect(props).toHaveProperty("font");
    expect(schema?.required).toContain("slideIndex");
    expect(schema?.required).toContain("shapeId");
  });

  it("has optional font formatting properties", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("format_text");
    const props = schema?.properties as Record<string, unknown> | undefined;
    const fontProps = (props?.font as Record<string, unknown>)?.properties as Record<string, unknown> | undefined;
    expect(fontProps).toHaveProperty("bold");
    expect(fontProps).toHaveProperty("italic");
    expect(fontProps).toHaveProperty("color");
    expect(fontProps).toHaveProperty("size");
    expect(fontProps).toHaveProperty("name");
  });

  it("returns ok false error when missing required slideIndex", async () => {
    const result = await executePowerPointTool("format_text", '{"shapeId": "test"}');
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toMatch(/required|string|invalid/i);
  });
});

describe("add_image", () => {
  it("has a schema registered with required base64Image and slideIndex", () => {
    const schemaMap = new Map(toolSchemas.map((t) => [t.function.name, t.function.parameters]));
    const schema = schemaMap.get("add_image");
    expect(schema).toBeDefined();
    const props = schema?.properties as Record<string, unknown> | undefined;
    expect(props).toHaveProperty("base64Image");
    expect(props).toHaveProperty("slideIndex");
    expect(schema?.required).toContain("base64Image");
    expect(schema?.required).toContain("slideIndex");
  });

  it("returns capability-gated unsupported result", async () => {
    const result = await executePowerPointTool("add_image", '{"base64Image": "dGVzdA==", "slideIndex": 0}');
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.action).toBe("add_image");
    expect(parsed.details).toBeDefined();
    expect(parsed.details.available).toBe(false);
    expect(parsed.details.requiredApi).toContain("addImage");
  });
});

describe("getCapabilities", () => {
  it("returns a capabilities map with read_slide_content listed", () => {
    const caps = getCapabilities();
    expect(caps).toHaveProperty("read_slide_content");
    expect(caps).toHaveProperty("delete_slide");
    expect(caps).toHaveProperty("duplicate_slide");
    expect(caps).toHaveProperty("reorder_slide");
    expect(caps).toHaveProperty("add_speaker_notes");
  });

  it("each capability has name, available, and requiredApi fields", () => {
    const caps = getCapabilities();
    for (const [_, cap] of Object.entries(caps)) {
      expect(cap).toHaveProperty("name");
      expect(cap).toHaveProperty("available");
      expect(cap).toHaveProperty("requiredApi");
    }
  });

  it("marks stable operations as available", () => {
    const caps = getCapabilities();
    expect(caps.read_slide_content.available).toBe(true);
    expect(caps.delete_slide.available).toBe(true);
    expect(caps.reorder_slide.available).toBe(true);
  });

  it("marks unsupported operations as unavailable with reason", () => {
    const caps = getCapabilities();
    expect(caps.duplicate_slide.available).toBe(false);
    expect(caps.duplicate_slide.reason).toBeDefined();
    expect(caps.add_speaker_notes.available).toBe(false);
    expect(caps.add_speaker_notes.reason).toBeDefined();
  });

  it("keeps speaker notes unavailable regardless of runtime requirement overrides", () => {
    setRequirementOverride("PowerPointApi", "1.2", true);
    const capability = getCapabilities().add_speaker_notes;
    expect(capability.available).toBe(false);
    expect(capability.reason).toContain("@types/office-js");
    expect(capability.reason).toContain("v1.0.599");
    setRequirementOverride("PowerPointApi", "1.2", null);
  });

  it("includes new P2 capability entries", () => {
    const caps = getCapabilities();
    expect(caps).toHaveProperty("find_and_replace_text");
    expect(caps).toHaveProperty("list_slide_layouts");
    expect(caps).toHaveProperty("set_slide_layout");
    expect(caps).toHaveProperty("format_text");
    expect(caps).toHaveProperty("add_image");
  });

  it("marks add_image as unavailable with reason", () => {
    const caps = getCapabilities();
    expect(caps.add_image.available).toBe(false);
    expect(caps.add_image.reason).toBeDefined();
  });

  it("marks list_slide_layouts and set_slide_layout as available", () => {
    const caps = getCapabilities();
    expect(caps.list_slide_layouts.available).toBe(true);
    expect(caps.set_slide_layout.available).toBe(true);
  });
});
