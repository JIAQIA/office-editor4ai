/**
 * 文件名: pageContent.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 页面内容获取工具的测试文件 | Test file for page content tool
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import {
  getPageContent,
  getPageText,
  getPageStats,
  type PageInfo,
  type GetPageContentOptions,
} from "../../../src/word-tools";

describe("pageContent - 页面内容获取工具", () => {
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

  describe("getPageContent - 获取页面内容", () => {
    it("应该返回正确的页面信息 | Should return correct page information", async () => {
      const mockContext = createMockContextWithPage(1, "Test page content");

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageContent(1);

      expect(result).toBeDefined();
      expect(result.index).toBe(0); // 0-based index
      expect(result.elements).toBeDefined();
      expect(result.text).toBe("Test page content");
    });

    it("应该正确处理页面编号验证 | Should validate page number correctly", async () => {
      await expect(getPageContent(0)).rejects.toThrow(/页面编号必须大于等于1/);
      await expect(getPageContent(-1)).rejects.toThrow(/页面编号必须大于等于1/);
    });

    it("应该在页面不存在时抛出错误 | Should throw error when page does not exist", async () => {
      const mockContext = createMockContextWithPages(2); // 只有2页

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await expect(getPageContent(5)).rejects.toThrow(/页面 5 不存在/);
    });

    it("应该支持 includeText 选项 | Should support includeText option", async () => {
      const mockContext = createMockContextWithPage(1, "Test content");

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithText = await getPageContent(1, { includeText: true });
      const resultWithoutText = await getPageContent(1, { includeText: false });

      expect(resultWithText.text).toBe("Test content");
      expect(resultWithoutText.text).toBeUndefined();
    });

    it("应该支持 includeImages 选项 | Should support includeImages option", async () => {
      const mockContext = createMockContextWithImages(1, 2);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithImages = await getPageContent(1, { includeImages: true });
      const resultWithoutImages = await getPageContent(1, { includeImages: false });

      // 验证图片元素的存在 / Verify image elements presence
      const imagesInResult = resultWithImages.elements.filter((e) => e.type === "InlinePicture");
      expect(imagesInResult.length).toBeGreaterThan(0);
    });

    it("应该支持 includeTables 选项 | Should support includeTables option", async () => {
      const mockContext = createMockContextWithTables(1, 2);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithTables = await getPageContent(1, { includeTables: true });
      const resultWithoutTables = await getPageContent(1, { includeTables: false });

      // 验证表格元素的存在 / Verify table elements presence
      const tablesInResult = resultWithTables.elements.filter((e) => e.type === "Table");
      expect(tablesInResult.length).toBeGreaterThan(0);
    });

    it("应该支持 includeContentControls 选项 | Should support includeContentControls option", async () => {
      const mockContext = createMockContextWithContentControls(1, 2);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithControls = await getPageContent(1, { includeContentControls: true });
      const resultWithoutControls = await getPageContent(1, { includeContentControls: false });

      // 验证内容控件元素的存在 / Verify content control elements presence
      const controlsInResult = resultWithControls.elements.filter((e) => e.type === "ContentControl");
      expect(controlsInResult.length).toBeGreaterThan(0);
    });

    it("应该支持 detailedMetadata 选项 | Should support detailedMetadata option", async () => {
      const mockContext = createMockContextWithDetailedParagraph(1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithMetadata = await getPageContent(1, { detailedMetadata: true });
      const resultWithoutMetadata = await getPageContent(1, { detailedMetadata: false });

      const paraWithMetadata = resultWithMetadata.elements[0] as any;
      const paraWithoutMetadata = resultWithoutMetadata.elements[0] as any;

      expect(paraWithMetadata.style).toBeDefined();
      expect(paraWithMetadata.alignment).toBeDefined();
      expect(paraWithoutMetadata.style).toBeUndefined();
      expect(paraWithoutMetadata.alignment).toBeUndefined();
    });

    it("应该支持 maxTextLength 选项 | Should support maxTextLength option", async () => {
      const longText = "a".repeat(1000);
      const mockContext = createMockContextWithPage(1, longText);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageContent(1, { maxTextLength: 100 });

      const paragraph = result.elements[0] as any;
      expect(paragraph.text?.length).toBeLessThanOrEqual(103); // 100 + "..."
      expect(paragraph.text).toContain("...");
    });

    it("应该正确处理多个段落 | Should handle multiple paragraphs correctly", async () => {
      const mockContext = createMockContextWithMultipleParagraphs(1, 5);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageContent(1);

      const paragraphs = result.elements.filter((e) => e.type === "Paragraph");
      expect(paragraphs.length).toBe(5);
    });

    it("应该正确处理表格单元格 | Should handle table cells correctly", async () => {
      const mockContext = createMockContextWithTableCells(1, 3, 4);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageContent(1, { includeTables: true, detailedMetadata: true });

      const table = result.elements.find((e) => e.type === "Table") as any;
      expect(table).toBeDefined();
      expect(table.rowCount).toBe(3);
      expect(table.columnCount).toBe(4);
      expect(table.cells).toBeDefined();
      expect(table.cells.length).toBe(3);
    });

    it("应该正确处理错误 | Should handle errors correctly", async () => {
      global.Word.run = vi.fn(async (callback) => {
        const mockContext = {
          document: {
            body: {
              getRange: vi.fn(() => {
                throw new Error("API Error");
              }),
            },
          },
          sync: vi.fn().mockResolvedValue(undefined),
        };
        return callback(mockContext);
      });

      await expect(getPageContent(1)).rejects.toThrow();
    });
  });

  describe("getPageText - 获取页面文本", () => {
    it("应该返回页面的纯文本内容 | Should return plain text content of page", async () => {
      const mockContext = createMockContextWithPage(1, "Test page text");

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageText(1);

      expect(result).toBe("Test page text");
      expect(typeof result).toBe("string");
    });

    it("应该在页面为空时返回空字符串 | Should return empty string for empty page", async () => {
      const mockContext = createMockContextWithPage(1, "");

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageText(1);

      expect(result).toBe("");
    });

    it("应该正确处理错误 | Should handle errors correctly", async () => {
      global.Word.run = vi.fn(async () => {
        throw new Error("API Error");
      });

      await expect(getPageText(1)).rejects.toThrow();
    });
  });

  describe("getPageStats - 获取页面统计信息", () => {
    it("应该返回正确的统计信息 | Should return correct statistics", async () => {
      const mockContext = createMockContextWithMixedContent(1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageStats(1);

      expect(result).toBeDefined();
      expect(result.pageIndex).toBe(0); // 0-based
      expect(result.elementCount).toBeGreaterThan(0);
      expect(result.characterCount).toBeGreaterThan(0);
      expect(result.paragraphCount).toBeGreaterThan(0);
    });

    it("应该正确统计各类元素 | Should count different element types correctly", async () => {
      const mockContext = createMockContextWithMixedContent(1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageStats(1);

      expect(result).toHaveProperty("paragraphCount");
      expect(result).toHaveProperty("tableCount");
      expect(result).toHaveProperty("imageCount");
      expect(result).toHaveProperty("contentControlCount");
      expect(result.paragraphCount).toBeGreaterThanOrEqual(0);
      expect(result.tableCount).toBeGreaterThanOrEqual(0);
      expect(result.imageCount).toBeGreaterThanOrEqual(0);
      expect(result.contentControlCount).toBeGreaterThanOrEqual(0);
    });

    it("应该正确统计字符数 | Should count characters correctly", async () => {
      const testText = "Hello World 你好世界";
      const mockContext = createMockContextWithPage(1, testText);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageStats(1);

      expect(result.characterCount).toBe(testText.length);
    });

    it("应该正确处理空页面 | Should handle empty page correctly", async () => {
      const mockContext = createMockContextWithEmptyPage(1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageStats(1);

      expect(result.elementCount).toBe(0);
      expect(result.characterCount).toBe(0);
      expect(result.paragraphCount).toBe(0);
    });

    it("应该正确处理错误 | Should handle errors correctly", async () => {
      global.Word.run = vi.fn(async () => {
        throw new Error("API Error");
      });

      await expect(getPageStats(1)).rejects.toThrow();
    });
  });

  describe("PageInfo 类型验证 | PageInfo type validation", () => {
    it("应该包含所有必需的字段 | Should contain all required fields", async () => {
      const mockContext = createMockContextWithPage(1, "Test");

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getPageContent(1);

      expect(result).toHaveProperty("index");
      expect(result).toHaveProperty("elements");
      expect(result).toHaveProperty("text");
      expect(Array.isArray(result.elements)).toBe(true);
    });
  });
});

/**
 * 创建包含单个页面的模拟上下文
 * Create mock context with a single page
 */
function createMockContextWithPage(pageNumber: number, text: string) {
  const mockParagraph = {
    text,
    style: "Normal",
    alignment: "Left",
    firstLineIndent: 0,
    leftIndent: 0,
    rightIndent: 0,
    lineSpacing: 1.5,
    spaceAfter: 10,
    spaceBefore: 0,
    isListItem: false,
    inlinePictures: {
      items: [],
      load: vi.fn().mockReturnThis(),
    },
    load: vi.fn().mockReturnThis(),
  };

  const mockPage = {
    index: pageNumber, // Word API uses 1-based index
    getRange: vi.fn().mockReturnValue({
      text,
      paragraphs: {
        items: [mockParagraph],
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含多个页面的模拟上下文
 * Create mock context with multiple pages
 */
function createMockContextWithPages(pageCount: number) {
  const mockPages = Array.from({ length: pageCount }, (_, i) => ({
    index: i + 1, // 1-based
    getRange: vi.fn().mockReturnValue({
      text: `Page ${i + 1} content`,
      paragraphs: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  }));

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: mockPages,
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含图片的模拟上下文
 * Create mock context with images
 */
function createMockContextWithImages(pageNumber: number, imageCount: number) {
  const mockImages = Array.from({ length: imageCount }, (_, i) => ({
    width: 200 + i * 10,
    height: 150 + i * 10,
    altTextTitle: `Image ${i + 1}`,
    altTextDescription: `Description ${i + 1}`,
    hyperlink: `https://example.com/image${i + 1}`,
    load: vi.fn().mockReturnThis(),
  }));

  const mockParagraph = {
    text: "Paragraph with images",
    style: "Normal",
    alignment: "Left",
    firstLineIndent: 0,
    leftIndent: 0,
    rightIndent: 0,
    lineSpacing: 1.5,
    spaceAfter: 10,
    spaceBefore: 0,
    isListItem: false,
    inlinePictures: {
      items: mockImages,
      load: vi.fn().mockReturnThis(),
    },
    load: vi.fn().mockReturnThis(),
  };

  const mockPage = {
    index: pageNumber,
    getRange: vi.fn().mockReturnValue({
      text: "Paragraph with images",
      paragraphs: {
        items: [mockParagraph],
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含表格的模拟上下文
 * Create mock context with tables
 */
function createMockContextWithTables(pageNumber: number, tableCount: number) {
  const mockTables = Array.from({ length: tableCount }, (_, i) => ({
    rowCount: 3 + i,
    columns: {
      items: Array(4 + i).fill({}),
      load: vi.fn().mockReturnThis(),
    },
    rows: {
      items: [],
      load: vi.fn().mockReturnThis(),
    },
    load: vi.fn().mockReturnThis(),
  }));

  const mockPage = {
    index: pageNumber,
    getRange: vi.fn().mockReturnValue({
      text: "Page with tables",
      paragraphs: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: mockTables,
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含内容控件的模拟上下文
 * Create mock context with content controls
 */
function createMockContextWithContentControls(pageNumber: number, controlCount: number) {
  const mockControls = Array.from({ length: controlCount }, (_, i) => ({
    text: `Control ${i + 1} content`,
    title: `Control ${i + 1}`,
    tag: `tag-${i + 1}`,
    type: "RichText",
    cannotDelete: false,
    cannotEdit: false,
    placeholderText: `Placeholder ${i + 1}`,
    load: vi.fn().mockReturnThis(),
  }));

  const mockPage = {
    index: pageNumber,
    getRange: vi.fn().mockReturnValue({
      text: "Page with controls",
      paragraphs: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: mockControls,
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含详细段落信息的模拟上下文
 * Create mock context with detailed paragraph
 */
function createMockContextWithDetailedParagraph(pageNumber: number) {
  const mockParagraph = {
    text: "Detailed paragraph",
    style: "Heading 1",
    alignment: "Centered",
    firstLineIndent: 36,
    leftIndent: 72,
    rightIndent: 72,
    lineSpacing: 2.0,
    spaceAfter: 12,
    spaceBefore: 6,
    isListItem: true,
    inlinePictures: {
      items: [],
      load: vi.fn().mockReturnThis(),
    },
    load: vi.fn().mockReturnThis(),
  };

  const mockPage = {
    index: pageNumber,
    getRange: vi.fn().mockReturnValue({
      text: "Detailed paragraph",
      paragraphs: {
        items: [mockParagraph],
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含多个段落的模拟上下文
 * Create mock context with multiple paragraphs
 */
function createMockContextWithMultipleParagraphs(pageNumber: number, paragraphCount: number) {
  const mockParagraphs = Array.from({ length: paragraphCount }, (_, i) => ({
    text: `Paragraph ${i + 1}`,
    style: "Normal",
    alignment: "Left",
    firstLineIndent: 0,
    leftIndent: 0,
    rightIndent: 0,
    lineSpacing: 1.5,
    spaceAfter: 10,
    spaceBefore: 0,
    isListItem: false,
    inlinePictures: {
      items: [],
      load: vi.fn().mockReturnThis(),
    },
    load: vi.fn().mockReturnThis(),
  }));

  const mockPage = {
    index: pageNumber,
    getRange: vi.fn().mockReturnValue({
      text: mockParagraphs.map((p) => p.text).join("\n"),
      paragraphs: {
        items: mockParagraphs,
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含表格单元格的模拟上下文
 * Create mock context with table cells
 */
function createMockContextWithTableCells(pageNumber: number, rowCount: number, colCount: number) {
  const mockRows = Array.from({ length: rowCount }, (_, rowIndex) => ({
    cells: {
      items: Array.from({ length: colCount }, (_, colIndex) => ({
        value: `Cell ${rowIndex}-${colIndex}`,
        width: 100,
        load: vi.fn().mockReturnThis(),
      })),
      load: vi.fn().mockReturnThis(),
    },
  }));

  const mockTable = {
    rowCount,
    columns: {
      items: Array(colCount).fill({}),
      load: vi.fn().mockReturnThis(),
    },
    rows: {
      items: mockRows,
      load: vi.fn().mockReturnThis(),
    },
    load: vi.fn().mockReturnThis(),
  };

  const mockPage = {
    index: pageNumber,
    getRange: vi.fn().mockReturnValue({
      text: "Page with table",
      paragraphs: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: [mockTable],
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建空页面的模拟上下文
 * Create mock context with empty page
 */
function createMockContextWithEmptyPage(pageNumber: number) {
  const mockPage = {
    index: pageNumber,
    getRange: vi.fn().mockReturnValue({
      text: "",
      paragraphs: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: [],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含混合内容的模拟上下文
 * Create mock context with mixed content
 */
function createMockContextWithMixedContent(pageNumber: number) {
  const mockParagraph = {
    text: "Test paragraph",
    style: "Normal",
    alignment: "Left",
    firstLineIndent: 0,
    leftIndent: 0,
    rightIndent: 0,
    lineSpacing: 1.5,
    spaceAfter: 10,
    spaceBefore: 0,
    isListItem: false,
    inlinePictures: {
      items: [
        {
          width: 200,
          height: 150,
          altTextTitle: "Test Image",
          altTextDescription: "Test",
          hyperlink: "",
          load: vi.fn().mockReturnThis(),
        },
      ],
      load: vi.fn().mockReturnThis(),
    },
    load: vi.fn().mockReturnThis(),
  };

  const mockTable = {
    rowCount: 2,
    columns: {
      items: [{}, {}],
      load: vi.fn().mockReturnThis(),
    },
    rows: {
      items: [],
      load: vi.fn().mockReturnThis(),
    },
    load: vi.fn().mockReturnThis(),
  };

  const mockControl = {
    text: "Control content",
    title: "Test Control",
    tag: "test-tag",
    type: "RichText",
    cannotDelete: false,
    cannotEdit: false,
    placeholderText: "",
    load: vi.fn().mockReturnThis(),
  };

  const mockPage = {
    index: pageNumber,
    getRange: vi.fn().mockReturnValue({
      text: "Test paragraph",
      paragraphs: {
        items: [mockParagraph],
        load: vi.fn().mockReturnThis(),
      },
      tables: {
        items: [mockTable],
        load: vi.fn().mockReturnThis(),
      },
      contentControls: {
        items: [mockControl],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      body: {
        getRange: vi.fn().mockReturnValue({
          pages: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        }),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}
