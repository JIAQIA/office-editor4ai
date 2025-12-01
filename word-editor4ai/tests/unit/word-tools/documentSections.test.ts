/**
 * 文件名: documentSections.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 文档节信息工具的测试文件 | Test file for document sections tool
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { getDocumentSections } from "../../../src/word-tools";

describe("documentSections - 文档节信息工具", () => {
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

  describe("getDocumentSections - 获取文档节信息", () => {
    it("应该返回正确的节信息列表 | Should return correct sections list", async () => {
      // 模拟包含2个节的文档
      const mockContext = {
        document: {
          sections: {
            items: [createMockSection(0), createMockSection(1)],
            load: vi.fn().mockReturnThis(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentSections();

      expect(result).toBeDefined();
      expect(result).toHaveLength(2);
      expect(result[0].index).toBe(0);
      expect(result[1].index).toBe(1);
    });

    it("应该返回正确的页眉页脚信息 | Should return correct header footer info", async () => {
      const mockContext = {
        document: {
          sections: {
            items: [createMockSection(0)],
            load: vi.fn().mockReturnThis(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentSections({ includeContent: true });

      expect(result).toHaveLength(1);
      expect(result[0].headers).toBeDefined();
      expect(result[0].headers).toHaveLength(3);
      expect(result[0].footers).toBeDefined();
      expect(result[0].footers).toHaveLength(3);
    });

    it("应该返回正确的页面设置信息 | Should return correct page setup info", async () => {
      const mockContext = {
        document: {
          sections: {
            items: [createMockSection(0)],
            load: vi.fn().mockReturnThis(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentSections({ includePageSetup: true });

      expect(result).toHaveLength(1);
      expect(result[0].pageSetup).toBeDefined();
      expect(result[0].pageSetup.pageWidth).toBe(612);
      expect(result[0].pageSetup.pageHeight).toBe(792);
      expect(result[0].pageSetup.orientation).toBe("portrait");
    });

    it("应该正确处理空文档 | Should handle empty document correctly", async () => {
      const mockContext = {
        document: {
          sections: {
            items: [],
            load: vi.fn().mockReturnThis(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentSections();

      expect(result).toBeDefined();
      expect(result).toHaveLength(0);
    });

    it("应该正确处理错误 | Should handle errors correctly", async () => {
      // 模拟 Word.run 内部抛出错误
      // Mock Word.run throwing error internally
      global.Word.run = vi.fn(async (callback) => {
        const mockContext = {
          document: {
            sections: {
              items: [],
              load: vi.fn(() => {
                throw new Error("API Error");
              }),
            },
          },
          sync: vi.fn().mockResolvedValue(undefined),
        };
        return callback(mockContext);
      });

      await expect(getDocumentSections()).rejects.toThrow(/获取文档节信息失败/);
    });

    it("应该支持 includeContent 选项 | Should support includeContent option", async () => {
      const mockContext = {
        document: {
          sections: {
            items: [createMockSection(0)],
            load: vi.fn().mockReturnThis(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithContent = await getDocumentSections({ includeContent: true });
      const resultWithoutContent = await getDocumentSections({ includeContent: false });

      expect(resultWithContent).toHaveLength(1);
      expect(resultWithoutContent).toHaveLength(1);
    });
  });

  describe("SectionInfo 类型验证 | SectionInfo type validation", () => {
    it("应该包含所有必需的字段 | Should contain all required fields", async () => {
      const mockContext = {
        document: {
          sections: {
            items: [createMockSection(0)],
            load: vi.fn().mockReturnThis(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentSections();
      const section = result[0];

      expect(section).toHaveProperty("index");
      expect(section).toHaveProperty("headers");
      expect(section).toHaveProperty("footers");
      expect(section).toHaveProperty("pageSetup");
      expect(section).toHaveProperty("sectionType");
      expect(section).toHaveProperty("differentFirstPage");
      expect(section).toHaveProperty("differentOddAndEven");
      expect(section).toHaveProperty("columnCount");
    });

    it("页面设置应该包含所有必需的字段 | Page setup should contain all required fields", async () => {
      const mockContext = {
        document: {
          sections: {
            items: [createMockSection(0)],
            load: vi.fn().mockReturnThis(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentSections();
      const pageSetup = result[0].pageSetup;

      expect(pageSetup).toHaveProperty("pageWidth");
      expect(pageSetup).toHaveProperty("pageHeight");
      expect(pageSetup).toHaveProperty("topMargin");
      expect(pageSetup).toHaveProperty("bottomMargin");
      expect(pageSetup).toHaveProperty("leftMargin");
      expect(pageSetup).toHaveProperty("rightMargin");
      expect(pageSetup).toHaveProperty("orientation");
    });
  });
});

/**
 * 创建模拟的节对象
 * Create mock section object
 */
function createMockSection(_index: number) {
  const mockPageSetup = {
    pageWidth: 612,
    pageHeight: 792,
    topMargin: 72,
    bottomMargin: 72,
    leftMargin: 72,
    rightMargin: 72,
    orientation: "Portrait",
    sectionStart: "NewPage",
    differentFirstPageHeaderFooter: false,
    oddAndEvenPagesHeaderFooter: false,
    load: vi.fn().mockReturnThis(),
  };

  const section = {
    body: {
      style: "Normal",
      type: "Normal",
      load: vi.fn().mockReturnThis(),
      parentSection: null as any, // 稍后设置 | Set later
    },
    getHeader: vi.fn((type: any) => createMockBody(`Header ${type}`)),
    getFooter: vi.fn((type: any) => createMockBody(`Footer ${type}`)),
    load: vi.fn().mockReturnThis(),
    pageSetup: mockPageSetup, // 添加 pageSetup 到 section | Add pageSetup to section
  };

  // 确保 body.parentSection 指向包含此 body 的 section
  // Ensure body.parentSection points to the section containing this body
  section.body.parentSection = section as any;

  return section;
}

/**
 * 创建模拟的 Body 对象（用于页眉页脚）
 * Create mock Body object (for headers/footers)
 */
function createMockBody(text: string) {
  return {
    text: text,
    load: vi.fn().mockReturnThis(),
  };
}
