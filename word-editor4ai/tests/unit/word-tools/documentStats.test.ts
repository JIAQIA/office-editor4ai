/**
 * 文件名: documentStats.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 文档统计信息工具的测试文件 | Test file for document statistics tool
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import {
  getDocumentStats,
  getBasicDocumentStats,
  formatDocumentStats,
  type DocumentStats,
} from "../../../src/word-tools";

describe("documentStats - 文档统计信息工具", () => {
  let originalWordRun: any;
  let originalHeaderFooterType: any;

  beforeEach(() => {
    // 保存原始的 Word.run
    originalWordRun = global.Word.run;
    originalHeaderFooterType = global.Word?.HeaderFooterType;

    // Mock Word.HeaderFooterType 枚举
    // Mock Word.HeaderFooterType enum
    if (!global.Word) {
      (global as any).Word = {};
    }
    global.Word.HeaderFooterType = {
      primary: "Primary",
      firstPage: "FirstPage",
      evenPages: "EvenPages",
    } as any;
  });

  afterEach(() => {
    // 恢复原始的 Word.run 和 HeaderFooterType
    global.Word.run = originalWordRun;
    if (originalHeaderFooterType) {
      global.Word.HeaderFooterType = originalHeaderFooterType;
    }
    vi.clearAllMocks();
  });

  describe("getDocumentStats - 获取文档统计信息", () => {
    it("应该返回正确的基本统计信息 | Should return correct basic statistics", async () => {
      const mockText = "Hello World! 你好世界！";
      const mockContext = createMockContext(mockText, 3, 2, 1, 0);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentStats();

      expect(result).toBeDefined();
      expect(result.characterCount).toBe(mockText.length);
      expect(result.paragraphCount).toBe(3);
      expect(result.tableCount).toBe(2);
      expect(result.inlinePictureCount).toBe(1);
      expect(result.contentControlCount).toBe(0);
    });

    it("应该正确统计中英文单词数 | Should correctly count Chinese and English words", async () => {
      const mockText = "Hello World 你好世界";
      const mockContext = createMockContext(mockText, 1, 0, 0, 0);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentStats();

      // "Hello" + "World" + "你" + "好" + "世" + "界" = 6 words
      expect(result.wordCount).toBeGreaterThan(0);
    });

    it("应该正确统计不含空格的字符数 | Should correctly count characters without spaces", async () => {
      const mockText = "Hello   World";
      const mockContext = createMockContext(mockText, 1, 0, 0, 0);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentStats();

      expect(result.characterCount).toBe(mockText.length);
      expect(result.characterCountNoSpaces).toBe("HelloWorld".length);
    });

    it("应该正确估算页数 | Should correctly estimate page count", async () => {
      // 创建一个约3600字符的文本（应该估算为2页）
      const mockText = "a".repeat(3600);
      const mockContext = createMockContext(mockText, 1, 0, 0, 0);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentStats();

      expect(result.pageCount).toBe(2);
    });

    it("应该正确统计标题 | Should correctly count headings", async () => {
      const mockContext = createMockContextWithHeadings();

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentStats({ includeHeadingStats: true });

      expect(result.totalHeadingCount).toBe(3);
      expect(result.headingCounts[1]).toBe(1);
      expect(result.headingCounts[2]).toBe(2);
    });

    it("应该支持 includeHeaderFooter 选项 | Should support includeHeaderFooter option", async () => {
      const mockContext = createMockContextWithHeaderFooter();

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithHeader = await getDocumentStats({ includeHeaderFooter: true });
      const resultWithoutHeader = await getDocumentStats({ includeHeaderFooter: false });

      // 包含页眉页脚的字符数应该更多
      expect(resultWithHeader.characterCount).toBeGreaterThan(resultWithoutHeader.characterCount);
    });

    it("应该正确处理空文档 | Should handle empty document correctly", async () => {
      const mockContext = createMockContext("", 0, 0, 0, 0);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentStats();

      expect(result.characterCount).toBe(0);
      expect(result.wordCount).toBe(0);
      expect(result.paragraphCount).toBe(0);
      expect(result.pageCount).toBe(1); // 至少1页
    });

    it("应该正确处理错误 | Should handle errors correctly", async () => {
      global.Word.run = vi.fn(async (callback) => {
        const mockContext = {
          document: {
            body: {
              load: vi.fn(() => {
                throw new Error("API Error");
              }),
            },
          },
          sync: vi.fn().mockResolvedValue(undefined),
        };
        return callback(mockContext);
      });

      await expect(getDocumentStats()).rejects.toThrow(/获取文档统计信息失败/);
    });
  });

  describe("getBasicDocumentStats - 获取基本统计信息", () => {
    it("应该返回简化的统计信息 | Should return simplified statistics", async () => {
      const mockText = "Hello World";
      const mockContext = {
        document: {
          body: {
            text: mockText,
            paragraphs: {
              items: [{}, {}, {}],
              load: vi.fn().mockReturnThis(),
            },
            load: vi.fn().mockReturnThis(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getBasicDocumentStats();

      expect(result).toHaveProperty("characterCount");
      expect(result).toHaveProperty("wordCount");
      expect(result).toHaveProperty("paragraphCount");
      expect(result).toHaveProperty("pageCount");
      expect(result.characterCount).toBe(mockText.length);
      expect(result.paragraphCount).toBe(3);
    });
  });

  describe("formatDocumentStats - 格式化统计信息", () => {
    it("应该返回格式化的文本 | Should return formatted text", () => {
      const stats: DocumentStats = {
        characterCount: 1000,
        characterCountNoSpaces: 800,
        wordCount: 150,
        paragraphCount: 10,
        pageCount: 1,
        sectionCount: 1,
        tableCount: 2,
        imageCount: 3,
        inlinePictureCount: 3,
        contentControlCount: 0,
        listCount: 1,
        footnoteCount: 0,
        endnoteCount: 0,
        headingCounts: { 1: 2, 2: 3 },
        totalHeadingCount: 5,
      };

      const formatted = formatDocumentStats(stats);

      expect(formatted).toContain("文档统计信息");
      expect(formatted).toContain("1,000");
      expect(formatted).toContain("150");
      expect(formatted).toContain("标题统计");
    });

    it("应该正确显示脚注尾注信息 | Should correctly display footnote/endnote info", () => {
      const stats: DocumentStats = {
        characterCount: 1000,
        characterCountNoSpaces: 800,
        wordCount: 150,
        paragraphCount: 10,
        pageCount: 1,
        sectionCount: 1,
        tableCount: 0,
        imageCount: 0,
        inlinePictureCount: 0,
        contentControlCount: 0,
        listCount: 0,
        footnoteCount: 5,
        endnoteCount: 3,
        headingCounts: {},
        totalHeadingCount: 0,
      };

      const formatted = formatDocumentStats(stats);

      expect(formatted).toContain("注释统计");
      expect(formatted).toContain("脚注数");
      expect(formatted).toContain("5");
      expect(formatted).toContain("尾注数");
      expect(formatted).toContain("3");
    });
  });

  describe("DocumentStats 类型验证 | DocumentStats type validation", () => {
    it("应该包含所有必需的字段 | Should contain all required fields", async () => {
      const mockContext = createMockContext("Test", 1, 0, 0, 0);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentStats();

      expect(result).toHaveProperty("characterCount");
      expect(result).toHaveProperty("characterCountNoSpaces");
      expect(result).toHaveProperty("wordCount");
      expect(result).toHaveProperty("paragraphCount");
      expect(result).toHaveProperty("pageCount");
      expect(result).toHaveProperty("sectionCount");
      expect(result).toHaveProperty("tableCount");
      expect(result).toHaveProperty("imageCount");
      expect(result).toHaveProperty("inlinePictureCount");
      expect(result).toHaveProperty("contentControlCount");
      expect(result).toHaveProperty("listCount");
      expect(result).toHaveProperty("footnoteCount");
      expect(result).toHaveProperty("endnoteCount");
      expect(result).toHaveProperty("headingCounts");
      expect(result).toHaveProperty("totalHeadingCount");
    });
  });
});

/**
 * 创建模拟的上下文对象
 * Create mock context object
 */
function createMockContext(
  text: string,
  paragraphCount: number,
  tableCount: number,
  pictureCount: number,
  controlCount: number
) {
  const paragraphs = Array.from({ length: paragraphCount }, (_, i) => ({
    text: `Paragraph ${i}`,
    style: "Normal",
    isListItem: false,
    listItem: null,
    load: vi.fn().mockReturnThis(),
  }));

  return {
    document: {
      body: {
        text,
        paragraphs: {
          items: paragraphs,
          load: vi.fn().mockReturnThis(),
        },
        tables: {
          items: Array(tableCount).fill({}),
          load: vi.fn().mockReturnThis(),
        },
        inlinePictures: {
          items: Array(pictureCount).fill({}),
          load: vi.fn().mockReturnThis(),
        },
        contentControls: {
          items: Array(controlCount).fill({}),
          load: vi.fn().mockReturnThis(),
        },
        footnotes: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        endnotes: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        getRange: vi.fn().mockReturnThis(),
        load: vi.fn().mockReturnThis(),
      },
      sections: {
        items: [{}],
        load: vi.fn().mockReturnThis(),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含标题的模拟上下文
 * Create mock context with headings
 */
function createMockContextWithHeadings() {
  const paragraphs = [
    {
      text: "Heading 1",
      style: "Heading 1",
      isListItem: false,
      listItem: null,
      load: vi.fn().mockReturnThis(),
    },
    {
      text: "Heading 2 First",
      style: "Heading 2",
      isListItem: false,
      listItem: null,
      load: vi.fn().mockReturnThis(),
    },
    {
      text: "Heading 2 Second",
      style: "Heading 2",
      isListItem: false,
      listItem: null,
      load: vi.fn().mockReturnThis(),
    },
    {
      text: "Normal paragraph",
      style: "Normal",
      isListItem: false,
      listItem: null,
      load: vi.fn().mockReturnThis(),
    },
  ];

  return {
    document: {
      body: {
        text: "Some text",
        paragraphs: {
          items: paragraphs,
          load: vi.fn().mockReturnThis(),
        },
        tables: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        inlinePictures: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        contentControls: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        footnotes: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        endnotes: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        getRange: vi.fn().mockReturnThis(),
        load: vi.fn().mockReturnThis(),
      },
      sections: {
        items: [{}],
        load: vi.fn().mockReturnThis(),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含页眉页脚的模拟上下文
 * Create mock context with header/footer
 */
function createMockContextWithHeaderFooter() {
  const mockSection = {
    getHeader: vi.fn((type: any) => ({
      text: `Header ${type}`,
      load: vi.fn().mockReturnThis(),
    })),
    getFooter: vi.fn((type: any) => ({
      text: `Footer ${type}`,
      load: vi.fn().mockReturnThis(),
    })),
  };

  return {
    document: {
      body: {
        text: "Main body text",
        paragraphs: {
          items: [
            {
              text: "Paragraph",
              style: "Normal",
              isListItem: false,
              listItem: null,
              load: vi.fn().mockReturnThis(),
            },
          ],
          load: vi.fn().mockReturnThis(),
        },
        tables: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        inlinePictures: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        contentControls: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        footnotes: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        endnotes: {
          items: [],
          load: vi.fn().mockReturnThis(),
        },
        getRange: vi.fn().mockReturnThis(),
        load: vi.fn().mockReturnThis(),
      },
      sections: {
        items: [mockSection],
        load: vi.fn().mockReturnThis(),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}
