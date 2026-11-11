import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("clearComments Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("derives a strict-A1 response address without loading range metadata", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const inRangeLocation = {
      address: "'Data Sheet'!B2",
      load: vi.fn()
    };
    const outRangeLocation = {
      address: "'Data Sheet'!E7",
      load: vi.fn()
    };

    const inRangeComment = {
      getLocation: vi.fn(() => inRangeLocation),
      delete: vi.fn()
    };
    const outRangeComment = {
      getLocation: vi.fn(() => outRangeLocation),
      delete: vi.fn()
    };

    const range = {
      get address(): string {
        throw new Error("PropertyNotLoaded: address");
      },
      load: vi.fn()
    };

    const comments = {
      items: [inRangeComment, outRangeComment],
      load: vi.fn()
    };

    const sheet = {
      getRange: vi.fn(() => range),
      comments
    };

    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const clearCommentsTool = excelTools.find(({ schema }) =>
      schema.function.name === "clear_comments"
    );
    const result = JSON.parse(await clearCommentsTool!.execute({
      sheet: "Data's Sheet",
      address: "$A$1:$C$3"
    }));

    expect(sync).toHaveBeenCalledTimes(3);
    expect(range.load).not.toHaveBeenCalled();
    expect(comments.load).toHaveBeenCalledWith("items");
    expect(inRangeComment.getLocation).toHaveBeenCalledOnce();
    expect(outRangeComment.getLocation).toHaveBeenCalledOnce();
    expect(inRangeLocation.load).toHaveBeenCalledWith("address");
    expect(outRangeLocation.load).toHaveBeenCalledWith("address");
    expect(inRangeComment.delete).toHaveBeenCalledOnce();
    expect(outRangeComment.delete).toHaveBeenCalledTimes(0);
    expect(result).toEqual({
      ok: true,
      action: "clear_comments",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!A1:C3",
      deletedCount: 1
    });
  });

  it("returns quickly when no comments are found", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const range = {
      get address(): string {
        throw new Error("PropertyNotLoaded: address");
      },
      load: vi.fn()
    };

    const comments = {
      items: [],
      load: vi.fn()
    };

    const sheet = {
      getRange: vi.fn(() => range),
      comments
    };

    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const clearCommentsTool = excelTools.find(({ schema }) =>
      schema.function.name === "clear_comments"
    );
    const result = JSON.parse(await clearCommentsTool!.execute({
      sheet: "Data's Sheet",
      address: "$A$1:$C$3"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).not.toHaveBeenCalled();
    expect(comments.load).toHaveBeenCalledWith("items");
    expect(result).toEqual({
      ok: true,
      action: "clear_comments",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!A1:C3",
      deletedCount: 0
    });
  });

  it("skips the final sync when comments are present but none are in range", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const locationOne = {
      address: "'Data Sheet'!E7",
      load: vi.fn()
    };
    const locationTwo = {
      address: "'Data Sheet'!F8",
      load: vi.fn()
    };

    const inWorkbookComment = {
      getLocation: vi.fn(() => locationOne),
      delete: vi.fn()
    };
    const outWorkbookComment = {
      getLocation: vi.fn(() => locationTwo),
      delete: vi.fn()
    };

    const range = {
      get address(): string {
        throw new Error("PropertyNotLoaded: address");
      },
      load: vi.fn()
    };

    const comments = {
      items: [inWorkbookComment, outWorkbookComment],
      load: vi.fn()
    };

    const sheet = {
      getRange: vi.fn(() => range),
      comments
    };

    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const clearCommentsTool = excelTools.find(({ schema }) =>
      schema.function.name === "clear_comments"
    );
    const result = JSON.parse(await clearCommentsTool!.execute({
      sheet: "Data's Sheet",
      address: "$A$1:$C$3"
    }));

    expect(sync).toHaveBeenCalledTimes(2);
    expect(range.load).not.toHaveBeenCalled();
    expect(comments.load).toHaveBeenCalledWith("items");
    expect(inWorkbookComment.getLocation).toHaveBeenCalledOnce();
    expect(outWorkbookComment.getLocation).toHaveBeenCalledOnce();
    expect(locationOne.load).toHaveBeenCalledWith("address");
    expect(locationTwo.load).toHaveBeenCalledWith("address");
    expect(inWorkbookComment.delete).toHaveBeenCalledTimes(0);
    expect(outWorkbookComment.delete).toHaveBeenCalledTimes(0);
    expect(result).toEqual({
      ok: true,
      action: "clear_comments",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!A1:C3",
      deletedCount: 0
    });
  });

  it("retains host address loading for non-A1 expressions", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);
    const range = {
      address: "Data!NamedRange",
      load: vi.fn()
    };
    const comments = {
      items: [],
      load: vi.fn()
    };
    const sheet = {
      getRange: vi.fn(() => range),
      comments
    };
    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const clearCommentsTool = excelTools.find(({ schema }) =>
      schema.function.name === "clear_comments"
    );
    const result = JSON.parse(await clearCommentsTool!.execute({
      sheet: "Data",
      address: "NamedRange"
    }));

    expect(sync).toHaveBeenCalledTimes(1);
    expect(range.load).toHaveBeenCalledWith("address");
    expect(result).toEqual({
      ok: true,
      action: "clear_comments",
      sheet: "Data",
      address: "Data!NamedRange",
      deletedCount: 0
    });
  });

  it("matches comments correctly when caller supplies lowercase A1 range (regression)", async () => {
    const sync = vi.fn().mockResolvedValue(undefined);

    const inRangeLocation = {
      address: "'Data Sheet'!B2",
      load: vi.fn()
    };
    const outRangeLocation = {
      address: "'Data Sheet'!E7",
      load: vi.fn()
    };

    const inRangeComment = {
      getLocation: vi.fn(() => inRangeLocation),
      delete: vi.fn()
    };
    const outRangeComment = {
      getLocation: vi.fn(() => outRangeLocation),
      delete: vi.fn()
    };

    const range = {
      get address(): string {
        throw new Error("PropertyNotLoaded: address");
      },
      load: vi.fn()
    };

    const comments = {
      items: [inRangeComment, outRangeComment],
      load: vi.fn()
    };

    const sheet = {
      getRange: vi.fn(() => range),
      comments
    };

    const context = {
      sync,
      workbook: {
        worksheets: {
          getItem: vi.fn(() => sheet)
        }
      }
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (callback: (requestContext: unknown) => Promise<unknown>) => callback(context)
    };

    const clearCommentsTool = excelTools.find(({ schema }) =>
      schema.function.name === "clear_comments"
    );
    // Caller passes lowercase A1 range — must still match canonical host addresses
    const result = JSON.parse(await clearCommentsTool!.execute({
      sheet: "Data Sheet",
      address: "a1:c3"
    }));

    expect(sync).toHaveBeenCalledTimes(3);
    expect(range.load).not.toHaveBeenCalled();
    expect(inRangeComment.delete).toHaveBeenCalledOnce();
    expect(outRangeComment.delete).toHaveBeenCalledTimes(0);
    expect(result).toEqual({
      ok: true,
      action: "clear_comments",
      sheet: "Data Sheet",
      address: "'Data Sheet'!a1:c3",
      deletedCount: 1
    });
  });
});