import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { excelTools } from "../excelTools";

describe("addComment Office.js batching", () => {
  let previousExcel: unknown;

  beforeEach(() => {
    previousExcel = (globalThis as unknown as { Excel?: unknown }).Excel;
  });

  afterEach(() => {
    (globalThis as unknown as { Excel?: unknown }).Excel = previousExcel;
  });

  it("loads existing comment then flushes the content update in a second sync", async () => {
    // First sync resolves (comment exists), second sync flushes the update
    const sync = vi.fn().mockResolvedValue(undefined);

    const range = {
      load: vi.fn()
    };

    let updatedContent = "";
    const existingComment = {
      load: vi.fn(),
      set content(value: string) {
        updatedContent = value;
      }
    };

    const comments = {
      getItemByCell: vi.fn(() => existingComment),
      add: vi.fn()
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

    const addCommentTool = excelTools.find(({ schema }) =>
      schema.function.name === "add_comment"
    );
    const result = JSON.parse(await addCommentTool!.execute({
      sheet: "Data's Sheet",
      address: "B2",
      text: "Review this row"
    }));

    // Two syncs: load check + content update flush
    expect(sync).toHaveBeenCalledTimes(2);
    expect(existingComment.load).toHaveBeenCalledWith("content");
    expect(updatedContent).toBe("Review this row");
    expect(comments.add).toHaveBeenCalledTimes(0);
    expect(result).toEqual({
      ok: true,
      action: "add_comment",
      sheet: "Data's Sheet",
      address: "'Data''s Sheet'!B2",
      updated: true
    });
  });

  it("handles ItemNotFound from sync to create a new comment", async () => {
    // First sync throws ItemNotFound (no existing comment), second sync flushes the add
    let callCount = 0;
    const sync = vi.fn().mockImplementation(() => {
      callCount += 1;
      if (callCount === 1) {
        return Promise.reject({
          code: "ItemNotFound",
          message: "The requested comment does not exist.",
          name: "OfficeExtension.Error"
        });
      }
      return Promise.resolve(undefined);
    });

    const range = {
      address: "Data!C5",
      load: vi.fn()
    };

    const commentProxy = {
      load: vi.fn()
    };

    const comments = {
      getItemByCell: vi.fn(() => commentProxy),
      add: vi.fn()
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

    const addCommentTool = excelTools.find(({ schema }) =>
      schema.function.name === "add_comment"
    );
    const result = JSON.parse(await addCommentTool!.execute({
      sheet: "Data",
      address: "C5",
      text: "Needs follow-up"
    }));

    // Two syncs: load check (fails) + add mutation flush
    expect(sync).toHaveBeenCalledTimes(2);
    expect(comments.getItemByCell).toHaveBeenCalledWith(range);
    expect(commentProxy.load).toHaveBeenCalledWith("content");
    expect(comments.add).toHaveBeenCalledWith(range, "Needs follow-up");
    expect(result).toEqual({
      ok: true,
      action: "add_comment",
      sheet: "Data",
      address: "Data!C5",
      updated: false
    });
  });
});
