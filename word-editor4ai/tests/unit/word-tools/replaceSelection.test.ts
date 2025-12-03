/**
 * 文件名: replaceSelection.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 替换选中内容工具的测试文件 | Test file for replace selection tool
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { replaceSelection, replaceTextAtSelection } from "../../../src/word-tools";

describe("replaceSelection - 替换选中内容工具", () => {
  let originalWordRun: any;

  beforeEach(() => {
    // 保存原始的 Word.run / Save original Word.run
    originalWordRun = global.Word.run;
  });

  afterEach(() => {
    // 恢复原始的 Word.run / Restore original Word.run
    global.Word.run = originalWordRun;
    vi.clearAllMocks();
  });

  describe("replaceSelection - 替换内容", () => {
    it("应该在选中位置替换文本 | Should replace text at selection", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        text: "新文本内容",
      });

      expect(mockContext.document.getSelection).toHaveBeenCalled();
      expect(mockContext.selection.insertText).toHaveBeenCalledWith("新文本内容", "Replace");
    });

    it("应该在选中位置替换文本并应用格式 | Should replace text with format at selection", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        text: "新文本内容",
        format: {
          fontName: "Arial",
          fontSize: 14,
          bold: true,
          color: "#FF0000",
        },
      });

      expect(mockContext.document.getSelection).toHaveBeenCalled();
      expect(mockContext.selection.insertText).toHaveBeenCalledWith("新文本内容", "Replace");
      expect(mockContext.insertedRange.font.name).toBe("Arial");
      expect(mockContext.insertedRange.font.size).toBe(14);
      expect(mockContext.insertedRange.font.bold).toBe(true);
      expect(mockContext.insertedRange.font.color).toBe("#FF0000");
    });

    it("应该在空选择时使用原格式 | Should use original format when selection is empty", async () => {
      const mockContext = createMockContextForInsert(true);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        text: "新文本内容",
      });

      expect(mockContext.document.getSelection).toHaveBeenCalled();
      expect(mockContext.selection.font.load).toHaveBeenCalled();
    });

    it("应该支持插入图片 | Should support inserting images", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        text: "文本内容",
        images: [
          {
            base64: "data:image/png;base64,iVBORw0KGgoAAAANS...",
            width: 200,
            height: 150,
            altText: "测试图片",
          },
        ],
      });

      expect(mockContext.document.getSelection).toHaveBeenCalled();
      expect(mockContext.insertedRange.insertInlinePictureFromBase64).toHaveBeenCalled();
    });

    it("应该支持插入多张图片 | Should support inserting multiple images", async () => {
      // 创建一个共享的spy来追踪所有图片插入调用 / Create a shared spy to track all image insertion calls
      const insertPictureSpy = vi.fn();
      const mockContext = createMockContextForInsertWithSpy(false, insertPictureSpy);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        text: "文本内容",
        images: [
          {
            base64: "data:image/png;base64,image1...",
            width: 200,
            height: 150,
          },
          {
            base64: "data:image/png;base64,image2...",
            width: 300,
            height: 200,
          },
        ],
      });

      expect(insertPictureSpy).toHaveBeenCalledTimes(2);
    });

    it("应该支持只插入图片 | Should support inserting only images", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        images: [
          {
            base64: "data:image/png;base64,iVBORw0KGgoAAAANS...",
            width: 200,
            height: 150,
          },
        ],
      });

      expect(mockContext.selection.insertInlinePictureFromBase64).toHaveBeenCalled();
    });

    it("应该在 replaceSelection=false 时插入而不替换 | Should insert without replacing when replaceSelection=false", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        text: "插入的文本",
        replaceSelection: false,
      });

      expect(mockContext.selection.getRange).toHaveBeenCalledWith("End");
    });

    it("应该在没有文本和图片时抛出错误 | Should throw error when no text or images provided", async () => {
      await expect(replaceSelection({})).rejects.toThrow(
        "必须提供文本或图片 / Must provide text or images"
      );
    });

    it("应该正确处理 base64 前缀 | Should handle base64 prefix correctly", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        images: [
          {
            base64: "data:image/png;base64,iVBORw0KGgoAAAANS...",
          },
        ],
      });

      // 验证 base64 前缀被移除 / Verify base64 prefix is removed
      expect(mockContext.selection.insertInlinePictureFromBase64).toHaveBeenCalledWith(
        "iVBORw0KGgoAAAANS...",
        "End"
      );
    });

    it("应该正确处理错误 | Should handle errors correctly", async () => {
      global.Word.run = vi.fn(async () => {
        throw new Error("API Error");
      });

      await expect(
        replaceSelection({
          text: "测试文本",
        })
      ).rejects.toThrow();
    });

    it("应该在图片插入失败时继续插入下一张 | Should continue inserting next image when one fails", async () => {
      const mockContext = createMockContextForInsert(false);

      // 第一张图片插入失败 / First image insertion fails
      let callCount = 0;
      mockContext.insertedRange.insertInlinePictureFromBase64 = vi.fn(() => {
        callCount++;
        if (callCount === 1) {
          throw new Error("Image insert failed");
        }
        return mockContext.inlinePicture;
      });

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        text: "文本",
        images: [{ base64: "image1..." }, { base64: "image2..." }],
      });

      expect(mockContext.insertedRange.insertInlinePictureFromBase64).toHaveBeenCalledTimes(2);
    });
  });

  describe("replaceTextAtSelection - 替换文本（简化版）", () => {
    it("应该插入文本 | Should insert text", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceTextAtSelection("新文本内容");

      expect(mockContext.document.getSelection).toHaveBeenCalled();
      expect(mockContext.selection.insertText).toHaveBeenCalledWith("新文本内容", "Replace");
    });

    it("应该插入文本并应用格式 | Should insert text with format", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceTextAtSelection("新文本内容", {
        fontName: "Arial",
        fontSize: 14,
        bold: true,
      });

      expect(mockContext.insertedRange.font.name).toBe("Arial");
      expect(mockContext.insertedRange.font.size).toBe(14);
      expect(mockContext.insertedRange.font.bold).toBe(true);
    });
  });

  describe("文本格式应用 | Text format application", () => {
    it("应该正确应用所有文本格式属性 | Should apply all text format properties correctly", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        text: "格式化文本",
        format: {
          fontName: "Times New Roman",
          fontSize: 16,
          bold: true,
          italic: true,
          underline: "Single",
          color: "#0000FF",
          highlightColor: "Yellow",
          strikeThrough: true,
          superscript: false,
          subscript: false,
        },
      });

      const font = mockContext.insertedRange.font;
      expect(font.name).toBe("Times New Roman");
      expect(font.size).toBe(16);
      expect(font.bold).toBe(true);
      expect(font.italic).toBe(true);
      expect(font.underline).toBe("Single");
      expect(font.color).toBe("#0000FF");
      expect(font.highlightColor).toBe("Yellow");
      expect(font.strikeThrough).toBe(true);
      expect(font.superscript).toBe(false);
      expect(font.subscript).toBe(false);
    });

    it("应该只应用指定的格式属性 | Should only apply specified format properties", async () => {
      const mockContext = createMockContextForInsert(false);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await replaceSelection({
        text: "部分格式化文本",
        format: {
          bold: true,
          fontSize: 18,
        },
      });

      const font = mockContext.insertedRange.font;
      expect(font.bold).toBe(true);
      expect(font.size).toBe(18);
      // 其他属性不应该被设置 / Other properties should not be set
      expect(font.name).toBeUndefined();
      expect(font.italic).toBeUndefined();
    });
  });
});

/**
 * 创建用于插入操作的模拟上下文（带spy支持）
 * Create mock context for insert operations (with spy support)
 */
function createMockContextForInsertWithSpy(hasSelection: boolean, insertPictureSpy?: any) {
  // 创建一个新的range用于图片插入后的位置 / Create a new range for position after image insertion
  const createImageInsertRange = (): any => ({
    insertInlinePictureFromBase64: vi.fn((_base64, _location) => {
      if (insertPictureSpy) insertPictureSpy(_base64, _location);
      return mockInlinePicture;
    }),
    getRange: vi.fn().mockReturnThis(),
  });

  const mockInlinePicture: any = {
    width: 0,
    height: 0,
    altTextTitle: "",
    getRange: vi.fn((_location) => {
      return createImageInsertRange();
    }),
  };

  const mockFont = {
    name: undefined,
    size: undefined,
    bold: undefined,
    italic: undefined,
    underline: undefined,
    color: undefined,
    highlightColor: undefined,
    strikeThrough: undefined,
    superscript: undefined,
    subscript: undefined,
    load: vi.fn().mockReturnThis(),
  };

  const mockInsertedRange: any = {
    font: mockFont,
    insertInlinePictureFromBase64: vi.fn((_base64, _location) => {
      if (insertPictureSpy) insertPictureSpy(_base64, _location);
      return mockInlinePicture;
    }),
    insertText: vi.fn((_text, _location) => {
      return mockInsertedRange;
    }),
    getRange: vi.fn().mockReturnThis(),
  };

  const mockSelection = {
    isEmpty: !hasSelection,
    text: hasSelection ? "选中的文本" : "",
    font: {
      name: "Calibri",
      size: 11,
      bold: false,
      italic: false,
      underline: "None",
      color: "#000000",
      highlightColor: "None",
      strikeThrough: false,
      superscript: false,
      subscript: false,
      load: vi.fn().mockReturnThis(),
    },
    clear: vi.fn(),
    insertText: vi.fn((_text, _location) => {
      return mockInsertedRange;
    }),
    insertInlinePictureFromBase64: vi.fn((_base64, _location) => {
      if (insertPictureSpy) insertPictureSpy(_base64, _location);
      return mockInlinePicture;
    }),
    getRange: vi.fn((_location) => {
      return mockInsertedRange;
    }),
    load: vi.fn().mockReturnThis(),
  };

  const mockContext = {
    document: {
      getSelection: vi.fn().mockReturnValue(mockSelection),
    },
    sync: vi.fn().mockResolvedValue(undefined),
    selection: mockSelection,
    insertedRange: mockInsertedRange,
    inlinePicture: mockInlinePicture,
  };

  return mockContext;
}

/**
 * 创建用于插入操作的模拟上下文
 * Create mock context for insert operations
 */
function createMockContextForInsert(hasSelection: boolean) {
  // 创建一个新的range用于图片插入后的位置 / Create a new range for position after image insertion
  const createImageInsertRange = () => ({
    insertInlinePictureFromBase64: vi.fn((_base64, _location) => {
      return mockInlinePicture;
    }),
    getRange: vi.fn().mockReturnThis(),
  });

  const mockInlinePicture = {
    width: 0,
    height: 0,
    altTextTitle: "",
    getRange: vi.fn((_location) => {
      return createImageInsertRange();
    }),
  };

  const mockFont = {
    name: undefined,
    size: undefined,
    bold: undefined,
    italic: undefined,
    underline: undefined,
    color: undefined,
    highlightColor: undefined,
    strikeThrough: undefined,
    superscript: undefined,
    subscript: undefined,
    load: vi.fn().mockReturnThis(),
  };

  const mockInsertedRange = {
    font: mockFont,
    insertInlinePictureFromBase64: vi.fn((_base64, _location) => {
      return mockInlinePicture;
    }),
    insertText: vi.fn((_text, _location) => {
      return mockInsertedRange;
    }),
    getRange: vi.fn().mockReturnThis(),
  };

  const mockSelection = {
    isEmpty: !hasSelection,
    text: hasSelection ? "选中的文本" : "",
    font: {
      name: "Calibri",
      size: 11,
      bold: false,
      italic: false,
      underline: "None",
      color: "#000000",
      highlightColor: "None",
      strikeThrough: false,
      superscript: false,
      subscript: false,
      load: vi.fn().mockReturnThis(),
    },
    clear: vi.fn(),
    insertText: vi.fn((_text, _location) => {
      return mockInsertedRange;
    }),
    insertInlinePictureFromBase64: vi.fn((_base64, _location) => {
      return mockInlinePicture;
    }),
    getRange: vi.fn((_location) => {
      return mockInsertedRange;
    }),
    load: vi.fn().mockReturnThis(),
  };

  const mockContext = {
    document: {
      getSelection: vi.fn().mockReturnValue(mockSelection),
    },
    sync: vi.fn().mockResolvedValue(undefined),
    selection: mockSelection,
    insertedRange: mockInsertedRange,
    inlinePicture: mockInlinePicture,
  };

  return mockContext;
}
