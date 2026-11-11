import { OpenAITool } from "./types";
import { errorToToolError, stringifyToolError, ToolErrorCategory, ToolErrorCategoryName, validationError } from "./toolErrors";
import { validateToolCall, JsonObject } from "./toolValidation";

type JsonRecord = JsonObject;

type PowerPointTool = {
  schema: OpenAITool;
  execute: (args: JsonRecord) => Promise<string>;
};

export interface ReadResult {
  text: string;
  truncated: boolean;
  originalLength: number;
  returnedLength: number;
}

export interface TextBoxLayout {
  left: number;
  top: number;
  width: number;
  height: number;
}

const MAX_READ_CHARS = 5000;
const MAX_OUTPUT_CHARS = 20000;
const NOT_SUPPORTED_MESSAGE = "Not supported by current PowerPoint Office.js API";
const TITLE_BOX: TextBoxLayout = { left: 60, top: 35, width: 840, height: 70 };
const BODY_BOX: TextBoxLayout = { left: 70, top: 125, width: 820, height: 330 };

/**
 * Known PowerPoint requirement sets and whether this built type snapshot supports them.
 * Since we can only evaluate against installed @types/office-js, not the runtime host,
 * we report available=true for requirement sets whose APIs exist in the types and whose
 * minVersion is ≤ the version we can confirm from the type declarations.
 */
interface RequirementResult {
  name: string;
  minVersion: string;
  available: boolean;
  reason?: string;
}

interface CapabilityEntry {
  name: string;
  available: boolean;
  requiredApi: string;
  reason?: string;
}

/**
 * Check whether a specific Office.js requirement set / version is available.
 * At build-time we can only verify against installed types — runtime `Office.context.requirements.isSetSupported`
 * is only callable inside a live PowerPoint session.
 *
 * For testing and deterministic reset, the internal table can be cleared by calling resetRequirementCache().
 */
let _requirementOverrides: Map<string, boolean> | null = null;

export function resetRequirementCache(): void {
  _requirementOverrides = new Map();
}

function _isSetSupported(name: string, minVersion: string): boolean {
  // If overrides exist, check them first
  if (_requirementOverrides !== null) {
    const key = `${name}:${minVersion}`;
    if (_requirementOverrides.has(key)) {
      return _requirementOverrides.get(key) ?? false;
    }
  }
  // Fall back to runtime check (works in PowerPoint) or type-level default
  try {
    if (typeof Office !== "undefined" && Office.context?.requirements?.isSetSupported) {
      return Office.context.requirements.isSetSupported(name, minVersion);
    }
  } catch {
    // Not in Office runtime
  }
  // Default: assume unknown versions are unavailable, known stable ones are available
  const stableVersions: Record<string, string[]> = {
    PowerPointApi: ["1.0", "1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"],
  };
  const versions = stableVersions[name];
  if (!versions) return false;
  return versions.includes(minVersion);
}

export function checkRequirement(name: string, minVersion: string): RequirementResult {
  const available = _isSetSupported(name, minVersion);
  const result: RequirementResult = { name, minVersion, available };
  if (!available) {
    result.reason = `Requirement set "${name}" version "${minVersion}" is not supported by the current Office.js types or runtime.`;
  }
  return result;
}

/**
 * Override a requirement check result for testing. Pass null for available to remove an override.
 * This mutates the shared cache; resetRequirementCache() clears all overrides.
 */
export function setRequirementOverride(name: string, minVersion: string, available: boolean | null): void {
  if (_requirementOverrides === null) {
    _requirementOverrides = new Map();
  }
  const key = `${name}:${minVersion}`;
  if (available === null) {
    _requirementOverrides.delete(key);
  } else {
    _requirementOverrides.set(key, available);
  }
}

export function getCapabilities(): Record<string, CapabilityEntry> {
  const capAddSpeakerNotes = _buildSpeakerNotesCapability();
  return {
    read_slide_content: {
      name: "read_slide_content",
      available: true,
      requiredApi: "PowerPointApi 1.4",
    },
    delete_slide: {
      name: "delete_slide",
      available: true,
      requiredApi: "PowerPointApi 1.2",
    },
    duplicate_slide: {
      name: "duplicate_slide",
      available: false,
      requiredApi: "PowerPointApi (no duplicate API in stable types)",
      reason: "No Slide.duplicate() or SlideCollection.duplicate() API exists in the installed @types/office-js (v1.0.599). The Office.js PowerPoint API surface does not expose a native slide-duplication method — use insertSlidesFromBase64 with an exported slide as a workaround.",
    },
    reorder_slide: {
      name: "reorder_slide",
      available: true,
      requiredApi: "PowerPointApi 1.8",
    },
    add_speaker_notes: capAddSpeakerNotes,
    find_and_replace_text: {
      name: "find_and_replace_text",
      available: true,
      requiredApi: "PowerPointApi 1.4 (text read + write via TextRange.text)",
    },
    list_slide_layouts: {
      name: "list_slide_layouts",
      available: true,
      requiredApi: "PowerPointApi 1.3 (SlideLayout.name, SlideLayout.id, SlideMaster.name)",
    },
    set_slide_layout: {
      name: "set_slide_layout",
      available: true,
      requiredApi: "PowerPointApi 1.8 (Slide.applyLayout)",
    },
    format_text: {
      name: "format_text",
      available: true,
      requiredApi: "PowerPointApi 1.4 (TextRange.font bold/italic/color/size/name)",
    },
    add_image: {
      name: "add_image",
      available: false,
      requiredApi: "PowerPointApi (no addImage/addPicture API in stable types)",
      reason: "PowerPoint.ShapeCollection in the installed @types/office-js (v1.0.599) does not expose addImage() or addPicture(). Only addTextBox, addGeometricShape, addLine, addTable, and addGroup exist. Image insertion is not available through the current Office.js PowerPoint API surface. Use Slide.getImageAsBase64() for export only.",
    },
  };
}

function _buildSpeakerNotesCapability(): CapabilityEntry {
  // Deterministic deployment contract: enabling this requires a deliberate
  // @types/office-js upgrade plus a reviewed implementation change.
  return {
    name: "add_speaker_notes",
    available: false,
    requiredApi: "PowerPointApi (no speakerNotes property on Slide in installed types)",
    reason: "The installed @types/office-js (v1.0.599) does not expose a speakerNotes property on PowerPoint.Slide. The Office.js PowerPoint speakerNotes API is not available in the current type definitions. The add_speaker_notes tool exists as an advertised schema returning an actionable unsupported response.",
  };
}

export function replaceLiteralAll(source: string, findText: string, replaceText: string): string {
  return source.split(findText).join(replaceText);
}

export function validateHtmlColor(color: string): boolean {
  return /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(color);
}

export const powerPointTools: PowerPointTool[] = [
  {
    schema: {
      type: "function",
      function: {
        name: "get_presentation_overview",
        description: "Get high-level details about the active PowerPoint presentation.",
        parameters: {
          type: "object",
          properties: {
            maxLength: {
              type: "integer",
              minimum: 1,
              maximum: MAX_OUTPUT_CHARS,
              description: "Maximum number of text characters to include in any returned text payload."
            }
          },
          additionalProperties: false
        }
      }
    },
    execute: (args) => {
      const maxLength = validatePresentationOverviewMaxLength(args.maxLength);
      if (maxLength.valid === false) return Promise.resolve(stringifyToolError(validationError(maxLength.error)));
      return getPresentationOverview(maxLength.value);
    }
  },
  {
    schema: {
      type: "function",
      function: {
        name: "read_selected_slides",
        description: "Read the currently selected slides from the active presentation.",
        parameters: {
          type: "object",
          properties: {
            maxLength: {
              type: "number",
              minimum: 1,
              description: "Maximum number of text characters to return."
            }
          },
          additionalProperties: false
        }
      }
    },
    execute: (args) => readSelectedSlides(optionalNumberArg(args, "maxLength") ?? MAX_READ_CHARS)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "read_all_slides",
        description: "Read a summary of all slides in the active presentation.",
        parameters: {
          type: "object",
          properties: {
            maxLength: {
              type: "number",
              minimum: 1,
              description: "Maximum number of text characters to return."
            }
          },
          additionalProperties: false
        }
      }
    },
    execute: (args) => readAllSlides(optionalNumberArg(args, "maxLength") ?? MAX_READ_CHARS)
  },
  {
    schema: {
      type: "function",
      function: {
        name: "create_slide",
        description: "Create a new slide at the end of the presentation.",
        parameters: {
          type: "object",
          properties: {
            layoutId: { type: "string", description: "Optional slide layout ID." },
            slideMasterId: { type: "string", description: "Optional slide master ID." }
          },
          additionalProperties: false
        }
      }
    },
    execute: (args) => createSlide({
      layoutId: optionalStringArg(args, "layoutId"),
      slideMasterId: optionalStringArg(args, "slideMasterId")
    })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "create_text_slide",
        description: "Create one new slide with a title and optional body/bullets using safe non-overlapping default layout. Prefer this over multiple add_text_box calls when drafting deck content.",
        parameters: {
          type: "object",
          properties: {
            title: { type: "string", description: "Slide title." },
            body: { type: "string", description: "Optional body paragraph text." },
            bullets: {
              type: "array",
              items: { type: "string" },
              description: "Optional bullet items. They are inserted into the body text box."
            },
            layoutId: { type: "string", description: "Optional slide layout ID." },
            slideMasterId: { type: "string", description: "Optional slide master ID." }
          },
          required: ["title"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => createTextSlide({
      title: textArg(args, "title"),
      body: optionalStringArg(args, "body"),
      bullets: optionalStringArrayArg(args, "bullets"),
      layoutId: optionalStringArg(args, "layoutId"),
      slideMasterId: optionalStringArg(args, "slideMasterId")
    })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "add_text_box",
        description: "Add a text box to a specific slide.",
        parameters: {
          type: "object",
          properties: {
            text: { type: "string" },
            slideIndex: { type: "integer", minimum: 0 },
            left: { type: "number", minimum: 0 },
            top: { type: "number", minimum: 0 },
            width: { type: "number", minimum: 1 },
            height: { type: "number", minimum: 1 }
          },
          required: ["text"],
          additionalProperties: false
        }
      }
    },
    execute: (args) =>
      addTextBox({
        text: textArg(args, "text"),
        slideIndex: optionalNumberArg(args, "slideIndex"),
        shapeOptions: parseShapeOptions(args)
      })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "update_slide_text",
        description: "Replace the currently selected text in the presentation.",
        parameters: {
          type: "object",
          properties: {
            text: { type: "string" },
            slideIndex: { type: "integer", minimum: 0 },
            shapeId: { type: "string" }
          },
          required: ["text"],
          additionalProperties: false
        }
      }
    },
    execute: (args) =>
      updateSlideText({
        text: textArg(args, "text"),
        slideIndex: optionalNumberArg(args, "slideIndex"),
        shapeId: optionalStringArg(args, "shapeId")
      })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "add_speaker_notes",
        description: "Add speaker notes to a slide (not supported in this PowerPoint Office.js snapshot).",
        parameters: {
          type: "object",
          properties: {
            slideIndex: { type: "integer", minimum: 0 },
            text: { type: "string" }
          },
          required: ["text"],
          additionalProperties: false
        }
      }
    },
  execute: (args) => {
    const slideIdx = optionalNumberArg(args, "slideIndex");
    const text = textArg(args, "text");
    // Use capability-gated unsupported with actionable details
    const caps = getCapabilities();
    const snCap = caps.add_speaker_notes;
    return Promise.resolve(JSON.stringify({
      ok: false,
      action: "add_speaker_notes",
      category: ToolErrorCategory.OfficeHost,
      error: snCap.reason ?? NOT_SUPPORTED_MESSAGE,
      details: {
        slideIndex: slideIdx ?? null,
        textLength: text.length,
        available: snCap.available,
        requiredApi: snCap.requiredApi
      }
    }));
  }
  },
  {
    schema: {
      type: "function",
      function: {
        name: "add_shape",
        description: "Add a geometric shape to a specific slide.",
        parameters: {
          type: "object",
          properties: {
            shapeType: { type: "string", description: "PowerPoint geometric shape name." },
            slideIndex: { type: "integer", minimum: 0 },
            left: { type: "number", minimum: 0 },
            top: { type: "number", minimum: 0 },
            width: { type: "number", minimum: 1 },
            height: { type: "number", minimum: 1 }
          },
          required: ["shapeType"],
          additionalProperties: false
        }
      }
    },
    execute: (args) =>
      addShape({
        shapeType: textArg(args, "shapeType"),
        slideIndex: optionalNumberArg(args, "slideIndex"),
        shapeOptions: parseShapeOptions(args)
      })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "read_slide_content",
        description: "Read text content from shapes on a specific slide by index. Returns text from all text-bearing shapes (text boxes, rectangles with text, placeholders) on the slide.",
        parameters: {
          type: "object",
          properties: {
            slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide to read." },
            maxLength: { type: "number", minimum: 1, description: "Maximum text characters to return per shape." }
          },
          required: ["slideIndex"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => readSlideContent({
      slideIndex: numberArg(args, "slideIndex"),
      maxLength: optionalNumberArg(args, "maxLength") ?? MAX_READ_CHARS
    })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "delete_slide",
        description: "Delete a slide by index. Has a built-in guard against deleting the last remaining slide — if the presentation has only one slide, returns an error instead.",
        parameters: {
          type: "object",
          properties: {
            slideIndex: { type: "integer", minimum: 0, description: "Optional zero-based index of the slide to delete. Defaults to the selected slide." }
          },
          additionalProperties: false
        }
      }
    },
    execute: (args) => deleteSlide({ slideIndex: optionalNumberArg(args, "slideIndex") })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "duplicate_slide",
        description: "Duplicate a slide by index. NOT SUPPORTED — the current Office.js PowerPoint API surface (@types/office-js v1.0.599) has no slide duplication API. This tool exists to provide an actionable diagnostics response; use insertSlidesFromBase64 as an alternative workaround.",
        parameters: {
          type: "object",
          properties: {
            slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide to duplicate." }
          },
          required: ["slideIndex"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => duplicateSlide({ slideIndex: optionalNumberArg(args, "slideIndex") })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "reorder_slide",
        description: "Move a slide from its current position to a new position within the presentation using Slide.moveTo (PowerPointApi 1.8).",
        parameters: {
          type: "object",
          properties: {
            fromIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide to move." },
            toIndex: { type: "integer", minimum: 0, description: "Zero-based target position for the slide." }
          },
          required: ["fromIndex", "toIndex"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => reorderSlide({
      fromIndex: numberArg(args, "fromIndex"),
      toIndex: numberArg(args, "toIndex")
    })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "find_and_replace_text",
        description: "Find and replace text on a slide. Searches shape text content using JavaScript-side string matching and replaces all occurrences. If shapeId is provided, only that shape is affected; otherwise searches all text-bearing shapes on the slide.",
        parameters: {
          type: "object",
          properties: {
            findText: { type: "string", description: "The text string to search for." },
            replaceText: { type: "string", description: "The replacement text string." },
            slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide. Defaults to selected slide." },
            shapeId: { type: "string", description: "Optional shape ID to limit the search to one shape." }
          },
          required: ["findText", "replaceText"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => findAndReplaceText({
      findText: textArg(args, "findText"),
      replaceText: textArg(args, "replaceText"),
      slideIndex: optionalNumberArg(args, "slideIndex"),
      shapeId: optionalStringArg(args, "shapeId")
    })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "list_slide_layouts",
        description: "List all slide masters and their available layouts. Returns slide master names/id and the layout names/types for each master.",
        parameters: {
          type: "object",
          properties: {},
          additionalProperties: false
        }
      }
    },
    execute: () => listSlideLayouts()
  },
  {
    schema: {
      type: "function",
      function: {
        name: "set_slide_layout",
        description: "Change the layout of a specific slide by applying a layout ID. Uses Slide.applyLayout() (PowerPointApi 1.8).",
        parameters: {
          type: "object",
          properties: {
            slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide to modify." },
            layoutId: { type: "string", description: "The ID of the layout to apply (obtained from list_slide_layouts)." }
          },
          required: ["slideIndex", "layoutId"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => setSlideLayout({
      slideIndex: numberArg(args, "slideIndex"),
      layoutId: textArg(args, "layoutId")
    })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "format_text",
        description: "Format text in a shape by setting font properties (bold, italic, color, size, name). Uses TextRange.font (PowerPointApi 1.4). Only provided properties are applied.",
        parameters: {
          type: "object",
          properties: {
            slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide." },
            shapeId: { type: "string", description: "ID of the shape containing the text to format." },
            font: {
              type: "object",
              properties: {
                bold: { type: "boolean", description: "Bold state (true/false)." },
                italic: { type: "boolean", description: "Italic state (true/false)." },
                color: { type: "string", description: "HTML color code (e.g., '#FF0000' for red)." },
                size: { type: "number", description: "Font size in points (e.g., 18)." },
                name: { type: "string", description: "Font family name (e.g., 'Calibri')." }
              },
              additionalProperties: false
            }
          },
          required: ["slideIndex", "shapeId"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => formatText({
      slideIndex: numberArg(args, "slideIndex"),
      shapeId: textArg(args, "shapeId"),
      font: args.font as Record<string, unknown> || {}
    })
  },
  {
    schema: {
      type: "function",
      function: {
        name: "add_image",
        description: "Add an image from a base64-encoded PNG or JPEG to a slide. NOT SUPPORTED — the current Office.js PowerPoint ShapeCollection API has no addImage method. This tool provides an actionable diagnostics response.",
        parameters: {
          type: "object",
          properties: {
            base64Image: { type: "string", description: "Base64-encoded image data (PNG or JPEG)." },
            slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the target slide." },
            left: { type: "number", minimum: 0, description: "Left position in points." },
            top: { type: "number", minimum: 0, description: "Top position in points." },
            width: { type: "number", minimum: 1, description: "Width in points." },
            height: { type: "number", minimum: 1, description: "Height in points." }
          },
          required: ["base64Image", "slideIndex"],
          additionalProperties: false
        }
      }
    },
    execute: (args) => {
      const base64 = textArg(args, "base64Image");
      const slideIdx = numberArg(args, "slideIndex");
      const caps = getCapabilities();
      const imgCap = caps.add_image;
      return Promise.resolve(JSON.stringify({
        ok: false,
        action: "add_image",
        category: ToolErrorCategory.OfficeHost,
        error: imgCap.reason ?? NOT_SUPPORTED_MESSAGE,
        details: {
          slideIndex: slideIdx,
          imageLength: base64.length,
          available: imgCap.available,
          requiredApi: imgCap.requiredApi,
          reason: imgCap.reason
        }
      }));
    }
  }
];

export const toolSchemas = powerPointTools.map((tool) => tool.schema);

export async function executePowerPointToolParsed(name: string, parsedArgs: JsonRecord): Promise<string> {
  const tool = powerPointTools.find((entry) => entry.schema.function.name === name);
  if (!tool) return stringifyToolError({ ok: false, error: `Unknown tool: ${name}`, category: ToolErrorCategory.Validation });
  try {
    return normalizeSerializedToolResult(await tool.execute(parsedArgs), ToolErrorCategory.OfficeHost);
  } catch (error) {
    return normalizeSerializedToolResult(stringifyError(error), ToolErrorCategory.OfficeHost);
  }
}

const TOOL_ERROR_CATEGORIES = new Set<string>(Object.values(ToolErrorCategory));

function normalizeSerializedToolResult(serialized: string, fallbackCategory: ToolErrorCategoryName): string {
  try {
    const parsed: unknown = JSON.parse(serialized);
    if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) return serialized;
    const result = parsed as Record<string, unknown>;
    if (result.ok !== false) return serialized;
    const category = typeof result.category === "string" && TOOL_ERROR_CATEGORIES.has(result.category)
      ? result.category
      : fallbackCategory;
    return JSON.stringify({ ...result, category });
  } catch {
    return serialized;
  }
}

export async function executePowerPointTool(name: string, rawArgs: string): Promise<string> {
  const validation = validateToolCall(name, rawArgs);
  if (!validation.valid) return stringifyToolError(validation);
  return executePowerPointToolParsed(name, validation.parsedArgs);
}

export async function getPresentationContext(): Promise<string> {
  return getPresentationOverview();
}

export function notSupportedResult(action: string, details?: Record<string, unknown>): string {
  return JSON.stringify({
    ok: false,
    action,
    error: NOT_SUPPORTED_MESSAGE,
    category: ToolErrorCategory.OfficeHost,
    ...(details ? { details } : {})
  });
}

async function getPresentationOverview(maxLength = MAX_READ_CHARS): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const presentation = context.presentation;
    presentation.load("title");

    const selectedSlides = presentation.getSelectedSlides();
    selectedSlides.load("items/id,index");

    const selectedTextRange = presentation.getSelectedTextRangeOrNullObject();
    selectedTextRange.load("text");

    const slideCountResult = presentation.slides.getCount();
    await context.sync();

    const selected = selectedSlides.items.map((slide) => ({
      id: slide.id,
      index: slide.index
    }));

    const text = selectedTextRange.isNullObject ? "" : selectedTextRange.text || "";
    const selectedText = truncateText(text, maxLength);

    const resolvedTitle = (presentation.title || "").trim() || "Untitled Presentation";

    return {
      ok: true,
      action: "get_presentation_overview",
      title: resolvedTitle,
      slideCount: slideCountResult.value,
      selectedSlides: selected,
      selectedCount: selected.length,
      selectedText: selectedText.text,
      selectedTextTruncated: selectedText.truncated,
      selectedTextOriginalLength: selectedText.originalLength,
      selectedTextReturnedLength: selectedText.returnedLength
    };
  });
}

async function readSelectedSlides(maxLength: number): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const presentation = context.presentation;
    const selectedSlides = presentation.getSelectedSlides();
    selectedSlides.load("items/id,index");

    const selectedTextRange = presentation.getSelectedTextRangeOrNullObject();
    selectedTextRange.load("text");

    await context.sync();

    const selected = selectedSlides.items.map((slide) => ({
      id: slide.id,
      index: slide.index
    }));

    const selectedText = selectedTextRange.isNullObject ? "" : selectedTextRange.text || "";
    const truncated = truncateText(selectedText, normalizeMaxLength(maxLength));
    const summaryText = selected.length
      ? `${selectedText || `Selected ${selected.length} slide(s).`}`
      : truncated.text;

    return {
      ok: true,
      action: "read_selected_slides",
      selectedCount: selected.length,
      selectedSlides: selected,
      truncated: truncated.truncated,
      originalLength: truncated.originalLength,
      returnedLength: truncated.returnedLength,
      text: truncateText(summaryText, normalizeMaxLength(maxLength)).text
    };
  });
}

async function readAllSlides(maxLength: number): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id,index");
    const slideCount = slides.getCount();

    await context.sync();

    const lines = slides.items.map((slide) => `Slide ${slide.index + 1} (id:${slide.id})`);
    const summary = lines.join("\n") || "No slides.";
    const truncated = truncateText(summary, normalizeMaxLength(maxLength));

    return {
      ok: true,
      action: "read_all_slides",
      slideCount: slideCount.value,
      truncated: truncated.truncated,
      originalLength: truncated.originalLength,
      returnedLength: truncated.returnedLength,
      text: truncated.text
    };
  });
}

async function createSlide(options: { layoutId?: string; slideMasterId?: string }): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    const beforeCount = slides.getCount();
    slides.load("items/id,index");
    await context.sync();
    const beforeIds = slides.items.map((slide) => slide.id);

    const addOptions: PowerPoint.AddSlideOptions = {};
    if (options.layoutId) {
      addOptions.layoutId = options.layoutId;
    }
    if (options.slideMasterId) {
      addOptions.slideMasterId = options.slideMasterId;
    }

    const hasOptions = Object.keys(addOptions).length > 0;
    if (hasOptions) {
      slides.add(addOptions);
    } else {
      slides.add();
    }

    await context.sync();
    slides.load("items/id,index");
    const afterCount = slides.getCount();
    await context.sync();
    const created = resolveCreatedSlide(beforeIds, slides.items.map((slide) => ({ id: slide.id, index: slide.index })));

    return {
      ok: true,
      action: "create_slide",
      beforeCount: beforeCount.value,
      afterCount: afterCount.value,
      newSlideIndex: created.index,
      newSlideId: created.id,
      layoutId: options.layoutId ?? null,
      slideMasterId: options.slideMasterId ?? null
    };
  });
}

async function createTextSlide(options: {
  title: string;
  body?: string;
  bullets?: string[];
  layoutId?: string;
  slideMasterId?: string;
}): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    const beforeCount = slides.getCount();
    slides.load("items/id,index");
    await context.sync();
    const beforeIds = slides.items.map((slide) => slide.id);

    const addOptions: PowerPoint.AddSlideOptions = {};
    if (options.layoutId) {
      addOptions.layoutId = options.layoutId;
    }
    if (options.slideMasterId) {
      addOptions.slideMasterId = options.slideMasterId;
    }

    if (Object.keys(addOptions).length > 0) {
      slides.add(addOptions);
    } else {
      slides.add();
    }

    await context.sync();
    slides.load("items/id,index");
    const afterCount = slides.getCount();
    await context.sync();
    const created = resolveCreatedSlide(beforeIds, slides.items.map((slide) => ({ id: slide.id, index: slide.index })));
    const slideIndex = created.index;
    const slide = slides.getItem(created.id);
    const titleShape = slide.shapes.addTextBox(options.title, TITLE_BOX);
    const bodyText = buildSlideBodyText(options.body, options.bullets);
    const bodyShape = bodyText ? slide.shapes.addTextBox(bodyText, BODY_BOX) : undefined;
    titleShape.load("id");
    bodyShape?.load("id");

    await context.sync();

    return {
      ok: true,
      action: "create_text_slide",
      beforeCount: beforeCount.value,
      afterCount: afterCount.value,
      newSlideIndex: slideIndex,
      titleShapeId: titleShape.id,
      bodyShapeId: bodyShape?.id ?? null,
      titleLength: options.title.length,
      bodyLength: bodyText.length
    };
  });
}

async function addTextBox(params: { text: string; slideIndex?: number; shapeOptions: PowerPoint.ShapeAddOptions }): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const slide = await resolveSlide(context, params.slideIndex);
    const textBox = slide.shapes.addTextBox(params.text, defaultTextBoxOptions(params.shapeOptions));
    textBox.load("id");

    await context.sync();

    return {
      ok: true,
      action: "add_text_box",
      slideId: slide.id,
      slideIndex: slide.index,
      shapeId: textBox.id,
      textLength: params.text.length
    };
  });
}

async function updateSlideText(params: { text: string; slideIndex?: number; shapeId?: string }): Promise<string> {
  return safelyPowerPoint(async (context) => {
    if (params.shapeId) {
      const slide = await resolveSlide(context, params.slideIndex);
      const shape = slide.shapes.getItem(params.shapeId);
      const textFrame = shape.getTextFrameOrNullObject();
      textFrame.load("isNullObject");
      await context.sync();

      if (textFrame.isNullObject) {
        return {
          ok: false,
          action: "update_slide_text",
          error: NOT_SUPPORTED_MESSAGE,
          details: {
          shapeId: params.shapeId
          }
        };
      }

      const textRange = textFrame.textRange;
      textRange.load("text");
      await context.sync();
      const beforeLength = textRange.text.length;
      textRange.text = params.text;
      await context.sync();

      return {
        ok: true,
        action: "update_slide_text",
        mode: "shape",
        slideId: slide.id,
        slideIndex: slide.index,
        shapeId: params.shapeId,
        beforeLength,
        afterLength: params.text.length
      };
    }

    const selectedTextRange = context.presentation.getSelectedTextRangeOrNullObject();
    selectedTextRange.load("text");
    await context.sync();

    if (selectedTextRange.isNullObject) {
      return {
        ok: false,
        action: "update_slide_text",
        error: "No selected text to update. Provide slide and shape identifiers or select text first."
      };
    }

    const beforeLength = selectedTextRange.text.length;
    selectedTextRange.text = params.text;
    await context.sync();

    return {
      ok: true,
      action: "update_slide_text",
      mode: "selection",
      beforeLength,
      afterLength: params.text.length
    };
  });
}

async function addShape(params: { shapeType: string; slideIndex?: number; shapeOptions: PowerPoint.ShapeAddOptions }): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const slide = await resolveSlide(context, params.slideIndex);
    const shape = slide.shapes.addGeometricShape(params.shapeType as PowerPoint.GeometricShapeType, params.shapeOptions);
    shape.load("id");
    await context.sync();

    return {
      ok: true,
      action: "add_shape",
      slideId: slide.id,
      slideIndex: slide.index,
      shapeId: shape.id,
      shapeType: params.shapeType
    };
  });
}

/**
 * Read text content from all text-bearing shapes on a slide.
 * Uses Shape.textFrame.textRange.text (PowerPointApi 1.4).
 * Filters to shapes that have a non-null textFrame with hasText.
 */
export interface SlideContentShape {
  shapeId: string;
  text: string;
  truncated: boolean;
  originalLength: number;
  returnedLength: number;
}

export interface SlideContentResult {
  ok: true;
  action: "read_slide_content";
  slideIndex: number;
  slideId: string;
  shapes: SlideContentShape[];
  totalShapes: number;
  textBearingShapes: number;
}

async function readSlideContent(params: { slideIndex: number; maxLength: number }): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    const total = slides.getCount();
    await context.sync();

    if (total.value === 0) {
      return { ok: false, action: "read_slide_content", error: "No slides in presentation." };
    }

    const safeIndex = validateExplicitSlideIndex(params.slideIndex, total.value);
    const slide = slides.getItemAt(safeIndex);
    slide.load("id,index");
    const shapes = slide.shapes;
    shapes.load("items/id");
    await context.sync();

    const shapeItems = shapes.items;
    const textShapes: SlideContentShape[] = [];
    const maxLen = normalizeMaxLength(params.maxLength);

    // Office.js returns a new proxy for each getTextFrameOrNullObject() call.
    // Keep and reuse the exact proxies whose properties are loaded.
    const textFrames = shapeItems.map((shape) => ({ shape, textFrame: shape.getTextFrameOrNullObject() }));
    for (const { textFrame } of textFrames) textFrame.load("isNullObject,hasText");
    await context.sync();

    const loadedTextRanges = textFrames
      .filter(({ textFrame }) => !textFrame.isNullObject && textFrame.hasText)
      .map(({ shape, textFrame }) => ({ shape, textRange: textFrame.textRange }));
    for (const { textRange } of loadedTextRanges) textRange.load("text");
    await context.sync();

    for (const { shape, textRange } of loadedTextRanges) {
      const rawText = textRange.text ?? "";
      const truncated = truncateText(rawText, maxLen);
      textShapes.push({
        shapeId: shape.id,
        text: truncated.text,
        truncated: truncated.truncated,
        originalLength: truncated.originalLength,
        returnedLength: truncated.returnedLength
      });
    }

    return {
      ok: true,
      action: "read_slide_content",
      slideIndex: safeIndex,
      slideId: slide.id,
      shapes: textShapes,
      totalShapes: shapeItems.length,
      textBearingShapes: textShapes.length
    } satisfies SlideContentResult;
  });
}

/**
 * Delete a slide by index. Uses Slide.delete() (PowerPointApi 1.2).
 * Has a last-slide guard: refuses to delete if the presentation has only one slide.
 */
async function deleteSlide(params: { slideIndex?: number }): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    const total = slides.getCount();
    await context.sync();

    if (total.value <= 1) {
      return {
        ok: false,
        action: "delete_slide",
        error: "Cannot delete the last remaining slide. A presentation must have at least one slide."
      };
    }

    const slide = await resolveSlide(context, params.slideIndex);
    const slideIndex = slide.index;
    slide.delete();
    await context.sync();

    return {
      ok: true,
      action: "delete_slide",
      slideId: slide.id,
      slideIndex,
      remainingCount: total.value - 1
    };
  });
}

/**
 * Duplicate a slide. No native duplicate API exists in @types/office-js v1.0.599.
 * Always returns capability-gated not-supported with actionable details.
 */
async function duplicateSlide(params: { slideIndex?: number }): Promise<string> {
  const caps = getCapabilities();
  const dupCap = caps.duplicate_slide;
  return JSON.stringify({
    ok: false,
    action: "duplicate_slide",
    category: ToolErrorCategory.OfficeHost,
    error: dupCap.reason ?? NOT_SUPPORTED_MESSAGE,
    details: {
      slideIndex: params.slideIndex !== undefined ? Math.floor(params.slideIndex) : null,
      available: dupCap.available,
      requiredApi: dupCap.requiredApi,
      reason: dupCap.reason
    }
  });
}

/**
 * Reorder a slide using Slide.moveTo(slideIndex) — PowerPointApi 1.8.
 */
async function reorderSlide(params: { fromIndex: number; toIndex: number }): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    const total = slides.getCount();
    await context.sync();

    if (total.value === 0) {
      return { ok: false, action: "reorder_slide", error: "No slides in presentation." };
    }

    const fromIdx = validateExplicitSlideIndex(params.fromIndex, total.value);
    const toIdx = validateExplicitSlideIndex(params.toIndex, total.value);

    const slide = slides.getItemAt(fromIdx);
    slide.load("id,index");
    await context.sync();

    slide.moveTo(toIdx);
    await context.sync();

    return {
      ok: true,
      action: "reorder_slide",
      slideId: slide.id,
      fromIndex: fromIdx,
      toIndex: toIdx
    };
  });
}

/**
 * Find and replace text on a slide. Uses JavaScript-side string replacement on each shape's
 * TextRange.text (PowerPointApi 1.4). If shapeId is provided, only that shape is updated.
 * If findText is empty, returns an error.
 */
async function findAndReplaceText(params: {
  findText: string;
  replaceText: string;
  slideIndex?: number;
  shapeId?: string;
}): Promise<string> {
  if (!params.findText) {
    return JSON.stringify({
      ok: false,
      action: "find_and_replace_text",
      error: "findText cannot be empty.",
      category: ToolErrorCategory.Validation
    });
  }
  return safelyPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    const total = slides.getCount();
    await context.sync();

    if (total.value === 0) {
      return { ok: false, action: "find_and_replace_text", error: "No slides in presentation." };
    }

    const slide = await resolveSlide(context, params.slideIndex);
    slide.load("id,index");
    await context.sync();

    const shapes = slide.shapes;
    if (params.shapeId) {
      // Target a specific shape
      const shape = shapes.getItem(params.shapeId);
      const textFrame = shape.getTextFrameOrNullObject();
      textFrame.load("isNullObject,hasText");
      await context.sync();

      if (textFrame.isNullObject || !textFrame.hasText) {
        return {
          ok: false,
          action: "find_and_replace_text",
          error: "Specified shape exists but has no text frame.",
          details: { slideIndex: slide.index, shapeId: params.shapeId }
        };
      }

      const textRange = textFrame.textRange;
      textRange.load("text");
      await context.sync();

      const before = textRange.text ?? "";
      const after = replaceLiteralAll(before, params.findText, params.replaceText);
      const replaceCount = (before.split(params.findText).length - 1);

      if (replaceCount === 0) {
        return {
          ok: true,
          action: "find_and_replace_text",
          slideIndex: slide.index,
          slideId: slide.id,
          shapeId: params.shapeId,
          mode: "single-shape",
          replacements: 0,
          message: "Search text not found in the specified shape."
        };
      }

      textRange.text = after;
      await context.sync();

      return {
        ok: true,
        action: "find_and_replace_text",
        slideIndex: slide.index,
        slideId: slide.id,
        shapeId: params.shapeId,
        mode: "single-shape",
        replacements: replaceCount
      };
    }

    // Search all text-bearing shapes on the slide
    shapes.load("items/id");
    await context.sync();

    const shapeItems = shapes.items;

    // First pass: load and retain each exact TextFrame proxy.
    const loadedTextFrames = shapeItems.map((shape) => ({
      shape,
      textFrame: shape.getTextFrameOrNullObject()
    }));
    for (const { textFrame } of loadedTextFrames) {
      textFrame.load("isNullObject,hasText");
    }
    await context.sync();

    // Second pass: retain each exact TextRange proxy whose text is loaded.
    const loadedTextRanges = loadedTextFrames
      .filter(({ textFrame }) => !textFrame.isNullObject && textFrame.hasText)
      .map(({ shape, textFrame }) => ({ shape, textRange: textFrame.textRange }));
    for (const { textRange } of loadedTextRanges) {
      textRange.load("text");
    }
    await context.sync();

    let totalReplacements = 0;
    const updatedShapes: { shapeId: string; replacements: number }[] = [];

    for (const { shape, textRange } of loadedTextRanges) {
      const before = textRange.text ?? "";
      const count = (before.split(params.findText).length - 1);
      if (count > 0) {
        textRange.text = replaceLiteralAll(before, params.findText, params.replaceText);
        totalReplacements += count;
        updatedShapes.push({ shapeId: shape.id, replacements: count });
      }
    }

    if (totalReplacements > 0) {
      await context.sync();
    }

    return {
      ok: true,
      action: "find_and_replace_text",
      slideIndex: slide.index,
      slideId: slide.id,
      mode: "all-shapes",
      totalReplacements,
      updatedShapes,
      shapesWithoutMatch: shapeItems.length - updatedShapes.length
    };
  });
}

/**
 * List all slide masters and their available layouts.
 * Uses SlideMasterCollection and SlideLayoutCollection (PowerPointApi 1.3).
 */
async function listSlideLayouts(): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id,name");
    await context.sync();

    const masterEntries: Record<string, unknown>[] = [];

    for (const master of masters.items) {
      const layouts = master.layouts;
      layouts.load("items/id,name,type");
      await context.sync();

      const layoutEntries = layouts.items.map((layout) => ({
        id: layout.id,
        name: layout.name,
        type: layout.type
      }));

      masterEntries.push({
        masterId: master.id,
        masterName: master.name,
        layouts: layoutEntries
      });
    }

    return {
      ok: true,
      action: "list_slide_layouts",
      masterCount: masters.items.length,
      masters: masterEntries
    };
  });
}

/**
 * Change the layout of a specific slide using Slide.applyLayout() (PowerPointApi 1.8).
 * The layoutId must correspond to a layout obtained from list_slide_layouts.
 */
async function setSlideLayout(params: { slideIndex: number; layoutId: string }): Promise<string> {
  return safelyPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    const total = slides.getCount();
    await context.sync();

    if (total.value === 0) {
      return { ok: false, action: "set_slide_layout", error: "No slides in presentation." };
    }

    const safeIndex = validateExplicitSlideIndex(params.slideIndex, total.value);
    const slide = slides.getItemAt(safeIndex);
    slide.load("id,index");
    await context.sync();

    // Find the layout object by ID across all masters
    const masters = context.presentation.slideMasters;
    masters.load("items/id");
    await context.sync();

    let targetLayout: PowerPoint.SlideLayout | null = null;

    for (const master of masters.items) {
      const layouts = master.layouts;
      layouts.load("items/id");
      await context.sync();

      for (const layout of layouts.items) {
        if (layout.id === params.layoutId) {
          targetLayout = layout;
          break;
        }
      }
      if (targetLayout) break;
    }

    if (!targetLayout) {
      return {
        ok: false,
        action: "set_slide_layout",
        error: `Layout with ID "${params.layoutId}" not found. Use list_slide_layouts to enumerate available layouts.`,
        details: { slideIndex: safeIndex, layoutId: params.layoutId }
      };
    }

    slide.applyLayout(targetLayout);
    await context.sync();

    return {
      ok: true,
      action: "set_slide_layout",
      slideIndex: safeIndex,
      slideId: slide.id,
      layoutId: params.layoutId
    };
  });
}

/**
 * Format text in a shape by setting font properties on TextRange.font.
 * Uses TextRange.font (PowerPointApi 1.4) for bold, italic, color, size, name.
 * Only non-null/defined properties in the font param are applied.
 */
async function formatText(params: {
  slideIndex: number;
  shapeId: string;
  font: Record<string, unknown>;
}): Promise<string> {
  if (typeof params.font?.color === "string" && !validateHtmlColor(params.font.color)) {
    return JSON.stringify({
      ok: false,
      action: "format_text",
      error: "font.color must be a hexadecimal HTML color in #RGB or #RRGGBB format.",
      category: ToolErrorCategory.Validation
    });
  }
  return safelyPowerPoint(async (context) => {
    const slides = context.presentation.slides;
    const total = slides.getCount();
    await context.sync();

    if (total.value === 0) {
      return { ok: false, action: "format_text", error: "No slides in presentation." };
    }

    const safeIndex = validateExplicitSlideIndex(params.slideIndex, total.value);
    const slide = slides.getItemAt(safeIndex);
    slide.load("id,index");
    await context.sync();

    const shape = slide.shapes.getItem(params.shapeId);
    const textFrame = shape.getTextFrameOrNullObject();
    textFrame.load("hasText");
    await context.sync();

    if (!textFrame.hasText) {
      return {
        ok: false,
        action: "format_text",
        error: "Specified shape has no text frame.",
        details: { slideIndex: safeIndex, shapeId: params.shapeId }
      };
    }

    const textRange = textFrame.textRange;
    const font = textRange.font;

    const fontInput = params.font || {};

    // Apply only the properties that are provided
    const applied: string[] = [];

    if (typeof fontInput.bold === "boolean") {
      font.bold = fontInput.bold;
      applied.push("bold");
    }
    if (typeof fontInput.italic === "boolean") {
      font.italic = fontInput.italic;
      applied.push("italic");
    }
    if (typeof fontInput.color === "string" && fontInput.color) {
      font.color = fontInput.color;
      applied.push("color");
    }
    if (typeof fontInput.size === "number" && fontInput.size > 0) {
      font.size = fontInput.size;
      applied.push("size");
    }
    if (typeof fontInput.name === "string" && fontInput.name) {
      font.name = fontInput.name;
      applied.push("name");
    }

    if (applied.length === 0) {
      return {
        ok: false,
        action: "format_text",
        error: "No valid font properties provided. Use bold, italic, color, size, or name.",
        category: ToolErrorCategory.Validation,
        details: { slideIndex: safeIndex, shapeId: params.shapeId }
      };
    }

    await context.sync();

    return {
      ok: true,
      action: "format_text",
      slideIndex: safeIndex,
      slideId: slide.id,
      shapeId: params.shapeId,
      applied
    };
  });
}

async function resolveSlide(context: PowerPoint.RequestContext, slideIndex?: number): Promise<PowerPoint.Slide> {
  const slides = context.presentation.slides;
  const selectedSlides = context.presentation.getSelectedSlides();

  slides.load("items");
  selectedSlides.load("items/index");
  const total = slides.getCount();

  await context.sync();

  if (total.value === 0) {
    throw new Error("No slides available in this presentation.");
  }

  let slide: PowerPoint.Slide;
  if (typeof slideIndex === "number") {
    const safeIndex = validateExplicitSlideIndex(slideIndex, total.value);
    slide = slides.getItemAt(safeIndex);
  } else {
    const firstSelectedIndex = selectedSlides.items[0]?.index;
    slide = slides.getItemAt(typeof firstSelectedIndex === "number" ? firstSelectedIndex : 0);
  }
  slide.load("id,index");
  await context.sync();
  return slide;
}

function parseShapeOptions(args: JsonRecord): PowerPoint.ShapeAddOptions {
  const options: PowerPoint.ShapeAddOptions = {};

  const left = optionalNumberArg(args, "left");
  if (left !== undefined) {
    options.left = left;
  }

  const top = optionalNumberArg(args, "top");
  if (top !== undefined) {
    options.top = top;
  }

  const width = optionalNumberArg(args, "width");
  if (width !== undefined) {
    options.width = width;
  }

  const height = optionalNumberArg(args, "height");
  if (height !== undefined) {
    options.height = height;
  }

  return options;
}

export function defaultTextBoxOptions(options: PowerPoint.ShapeAddOptions): PowerPoint.ShapeAddOptions {
  const hasPlacement = [options.left, options.top, options.width, options.height].some((value) => typeof value === "number");
  return hasPlacement ? options : { ...BODY_BOX };
}

export function buildSlideBodyText(body?: string, bullets?: string[]): string {
  const parts: string[] = [];
  const cleanBody = body?.trim();
  if (cleanBody) {
    parts.push(cleanBody);
  }

  const cleanBullets = (bullets ?? []).map((item) => item.trim()).filter(Boolean);
  if (cleanBullets.length) {
    parts.push(cleanBullets.map((item) => `• ${item}`).join("\n"));
  }

  return parts.join("\n\n");
}

type PresentationOverviewMaxLength =
  | { valid: true; value: number }
  | { valid: false; error: string };

function validatePresentationOverviewMaxLength(raw: unknown): PresentationOverviewMaxLength {
  if (raw === undefined) {
    return { valid: true, value: MAX_READ_CHARS };
  }
  if (typeof raw !== "number" || !Number.isFinite(raw)) {
    return { valid: false, error: "maxLength must be a finite number." };
  }
  if (!Number.isInteger(raw)) {
    return { valid: false, error: "maxLength must be an integer." };
  }
  if (raw < 1) {
    return { valid: false, error: "maxLength is below minimum 1." };
  }
  if (raw > MAX_OUTPUT_CHARS) {
    return { valid: false, error: `maxLength exceeds maximum ${MAX_OUTPUT_CHARS}.` };
  }
  return { valid: true, value: raw };
}

function normalizeMaxLength(raw: number): number {
  if (!Number.isFinite(raw) || raw < 1) {
    return MAX_READ_CHARS;
  }
  return Math.min(Math.max(Math.floor(raw), 1), MAX_OUTPUT_CHARS);
}

export function truncateText(value: string, maxLength: number): ReadResult {
  const original = value ?? "";
  const normalized = normalizeMaxLength(maxLength);
  const text = original.slice(0, normalized);
  const returnedLength = text.length;
  return {
    text,
    truncated: returnedLength < original.length,
    originalLength: original.length,
    returnedLength
  };
}

function textArg(args: JsonRecord, name: string): string {
  const value = args[name];
  if (typeof value !== "string") {
    throw new Error(`Missing or invalid string argument "${name}".`);
  }
  return value;
}

function optionalStringArg(args: JsonRecord, name: string): string | undefined {
  const value = args[name];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "string") {
    throw new Error(`Invalid string argument "${name}".`);
  }
  const trimmed = value.trim();
  return trimmed.length ? trimmed : undefined;
}

function optionalStringArrayArg(args: JsonRecord, name: string): string[] | undefined {
  const value = args[name];
  if (value === undefined) {
    return undefined;
  }
  if (!Array.isArray(value) || !value.every((item) => typeof item === "string")) {
    throw new Error(`Invalid string array argument "${name}".`);
  }
  return value;
}

function numberArg(args: JsonRecord, name: string): number {
  const value = args[name];
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.floor(value);
  }
  throw new Error(`Missing or invalid number argument "${name}".`);
}

function optionalNumberArg(args: JsonRecord, name: string): number | undefined {
  const value = args[name];
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`Invalid number argument "${name}".`);
  }
  return Math.floor(value);
}

async function safelyPowerPoint(work: (context: PowerPoint.RequestContext) => Promise<Record<string, unknown>>): Promise<string> {
  try {
    const result = await PowerPoint.run(work);
    return normalizeSerializedToolResult(JSON.stringify(result), ToolErrorCategory.OfficeHost);
  } catch (error) {
    return normalizeSerializedToolResult(stringifyError(error), ToolErrorCategory.OfficeHost);
  }
}

function stringifyError(error: unknown): string {
  return stringifyToolError(errorToToolError(error));
}

export function validateExplicitSlideIndex(index: number, count: number): number {
  if (!Number.isFinite(index) || !Number.isInteger(index) || index < 0 || index >= count) {
    const error = new Error(`slideIndex must be an integer from 0 through ${Math.max(0, count - 1)}.`) as Error & { category: ToolErrorCategoryName };
    error.category = ToolErrorCategory.Validation;
    throw error;
  }
  return index;
}

export function resolveCreatedSlide(beforeIds: string[], afterSlides: Array<string | { id: string; index: number }>): { id: string; index: number } {
  const previous = new Set(beforeIds);
  const added = afterSlides.map((slide, position) => typeof slide === "string" ? { id: slide, index: position } : slide).filter(({ id }) => !previous.has(id));
  if (added.length !== 1) throw new Error("Expected exactly one new slide after creation.");
  return added[0];
}
