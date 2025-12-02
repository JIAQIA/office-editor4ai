/**
 * 文件名: rangeContent.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: rangeContent 工具的单元测试
 */

import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { getRangeContent } from "../../../src/word-tools";
import type { ContentInfo } from "../../../src/word-tools/types";

describe("rangeContent", () => {
  let originalWordRun: any;

  beforeEach(() => {
    // 保存原始的 Word.run / Save original Word.run
    originalWordRun = global.Word?.run;
  });

  afterEach(() => {
    // 恢复原始的 Word.run / Restore original Word.run
    if (originalWordRun) {
      global.Word.run = originalWordRun;
    }
    vi.clearAllMocks();
  });

  describe("getRangeContent - 段落索引定位 / Paragraph Index Locator", () => {
    it("应该能够通过段落索引获取单个段落内容 / Should get single paragraph content by index", async () => {
      const mockContext = createMockContext();
      const mockPara = createMockParagraph("这是第一个段落", "Normal");
      mockPara.inlinePictures = {
        items: [],
        load: vi.fn(),
      } as any;

      const mockRange = createMockRange("这是第一个段落");
      mockRange.paragraphs = {
        items: [mockPara],
        load: vi.fn(),
      } as any;
      mockRange.tables = {
        items: [],
        load: vi.fn(),
      } as any;
      mockRange.contentControls = {
        items: [],
        load: vi.fn(),
      } as any;

      const mockParagraphs = {
        items: [mockPara, createMockParagraph("第二个段落", "Normal")],
        load: vi.fn(),
      };

      mockContext.document.body.paragraphs = mockParagraphs;

      // 模拟 getRange 方法 / Mock getRange method
      mockPara.getRange = vi.fn((location?: any) => {
        if (location === "start" || location === "end") {
          return {
            expandTo: vi.fn().mockReturnValue(mockRange),
          };
        }
        return mockRange;
      });

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
        RangeLocation: {
          start: "start",
          end: "end",
        },
      } as any;

      const result = await getRangeContent(
        { type: "paragraph", startIndex: 0 },
        { includeText: true }
      );

      expect(result.text).toBe("这是第一个段落");
      expect(result.elements).toHaveLength(1);
      expect(result.metadata?.locatorType).toBe("paragraph");
    });

    it("应该能够通过段落索引获取范围内容 / Should get range content by paragraph indices", async () => {
      const mockContext = createMockContext();
      const mockPara1 = createMockParagraph("第一个段落", "Normal");
      const mockPara2 = createMockParagraph("第二个段落", "Normal");

      mockPara1.inlinePictures = { items: [], load: vi.fn() } as any;
      mockPara2.inlinePictures = { items: [], load: vi.fn() } as any;

      const mockRange = createMockRange("第一个段落\n第二个段落");
      mockRange.paragraphs = {
        items: [mockPara1, mockPara2],
        load: vi.fn(),
      } as any;
      mockRange.tables = { items: [], load: vi.fn() } as any;
      mockRange.contentControls = { items: [], load: vi.fn() } as any;

      const mockParagraphs = {
        items: [mockPara1, mockPara2, createMockParagraph("第三个段落", "Normal")],
        load: vi.fn(),
      };

      mockContext.document.body.paragraphs = mockParagraphs;

      mockPara1.getRange = vi.fn((location?: any) => {
        if (location === "start") {
          return { expandTo: vi.fn().mockReturnValue(mockRange) };
        }
        return mockRange;
      });

      mockPara2.getRange = vi.fn((location?: any) => {
        if (location === "end") {
          return mockRange;
        }
        return mockRange;
      });

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
        RangeLocation: {
          start: "start",
          end: "end",
        },
      } as any;

      const result = await getRangeContent(
        { type: "paragraph", startIndex: 0, endIndex: 1 },
        { includeText: true }
      );

      expect(result.elements).toHaveLength(2);
      expect(result.metadata?.paragraphCount).toBe(2);
    });
  });

  describe("getRangeContent - 标题定位 / Heading Locator", () => {
    it("应该能够通过标题文本获取内容 / Should get content by heading text", async () => {
      const mockContext = createMockContext();
      const mockHeading = createMockParagraph("第一章 引言", "Heading1");
      mockHeading.styleBuiltIn = 1; // Heading1
      mockHeading.inlinePictures = { items: [], load: vi.fn() } as any;

      const mockRange = createMockRange("第一章 引言");
      mockRange.paragraphs = {
        items: [mockHeading],
        load: vi.fn(),
      } as any;
      mockRange.tables = { items: [], load: vi.fn() } as any;
      mockRange.contentControls = { items: [], load: vi.fn() } as any;

      const mockParagraphs = {
        items: [mockHeading, createMockParagraph("正文内容", "Normal")],
        load: vi.fn(),
      };

      mockContext.document.body.paragraphs = mockParagraphs;
      mockHeading.getRange = vi.fn().mockReturnValue(mockRange);

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
        BuiltInStyleName: {
          heading1: 1,
          heading9: 9,
        },
      } as any;

      const result = await getRangeContent(
        { type: "heading", text: "第一章" },
        { includeText: true }
      );

      expect(result.text).toBe("第一章 引言");
      expect(result.metadata?.locatorType).toBe("heading");
    });
  });

  describe("getRangeContent - 节定位 / Section Locator", () => {
    it("应该能够通过节索引获取内容 / Should get content by section index", async () => {
      const mockContext = createMockContext();
      const mockPara = createMockParagraph("节内容", "Normal");
      mockPara.inlinePictures = { items: [], load: vi.fn() } as any;

      const mockRange = createMockRange("节内容");
      mockRange.paragraphs = {
        items: [mockPara],
        load: vi.fn(),
      } as any;
      mockRange.tables = { items: [], load: vi.fn() } as any;
      mockRange.contentControls = { items: [], load: vi.fn() } as any;

      const mockSection = {
        body: {
          getRange: vi.fn().mockReturnValue(mockRange),
        },
      };

      mockContext.document.sections = {
        items: [mockSection],
        load: vi.fn(),
      };

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      const result = await getRangeContent(
        { type: "section", index: 0 },
        { includeText: true }
      );

      expect(result.text).toBe("节内容");
      expect(result.metadata?.locatorType).toBe("section");
    });
  });

  describe("getRangeContent - 内容控件定位 / Content Control Locator", () => {
    it("应该能够通过控件标题获取内容 / Should get content by control title", async () => {
      const mockContext = createMockContext();
      const mockPara = createMockParagraph("控件内容", "Normal");
      mockPara.inlinePictures = { items: [], load: vi.fn() } as any;

      const mockRange = createMockRange("控件内容");
      mockRange.paragraphs = {
        items: [mockPara],
        load: vi.fn(),
      } as any;
      mockRange.tables = { items: [], load: vi.fn() } as any;
      mockRange.contentControls = { items: [], load: vi.fn() } as any;

      const mockControl = {
        title: "测试控件",
        tag: "test-tag",
        load: vi.fn(),
        getRange: vi.fn().mockReturnValue(mockRange),
      };

      mockContext.document.body.contentControls = {
        items: [mockControl],
        load: vi.fn(),
      };

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      const result = await getRangeContent(
        { type: "contentControl", title: "测试控件" },
        { includeText: true }
      );

      expect(result.text).toBe("控件内容");
      expect(result.metadata?.locatorType).toBe("contentControl");
    });
  });

  describe("getRangeContent - 错误处理 / Error Handling", () => {
    it("应该在段落索引超出范围时抛出错误 / Should throw error when paragraph index out of range", async () => {
      const mockContext = createMockContext();
      mockContext.document.body.paragraphs = {
        items: [createMockParagraph("段落1", "Normal")],
        load: vi.fn(),
      };

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      await expect(
        getRangeContent({ type: "paragraph", startIndex: 10 })
      ).rejects.toThrow("起始段落索引超出范围");
    });

    it("应该在找不到标题时抛出错误 / Should throw error when heading not found", async () => {
      const mockContext = createMockContext();
      const mockPara = createMockParagraph("正文", "Normal");
      mockPara.styleBuiltIn = 0;
      mockPara.inlinePictures = { items: [], load: vi.fn() } as any;

      mockContext.document.body.paragraphs = {
        items: [mockPara],
        load: vi.fn(),
      };

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
        BuiltInStyleName: {
          heading1: 1,
          heading9: 9,
        },
      } as any;

      await expect(
        getRangeContent({ type: "heading", text: "不存在的标题" })
      ).rejects.toThrow("找不到匹配的标题");
    });

    it("应该在找不到内容控件时抛出错误 / Should throw error when content control not found", async () => {
      const mockContext = createMockContext();
      mockContext.document.body.contentControls = {
        items: [],
        load: vi.fn(),
      };

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      await expect(
        getRangeContent({ type: "contentControl", title: "不存在的控件" })
      ).rejects.toThrow("找不到匹配的内容控件");
    });

    it("应该能够处理 API 调用失败 / Should handle API call failures", async () => {
      global.Word = {
        run: vi.fn().mockRejectedValue(new Error("API 调用失败")),
      } as any;

      await expect(
        getRangeContent({ type: "paragraph", startIndex: 0 })
      ).rejects.toThrow("API 调用失败");
    });
  });

  describe("getRangeContent - 选项配置 / Options Configuration", () => {
    it("应该能够排除文本内容 / Should exclude text content", async () => {
      const mockContext = createMockContext();
      const mockPara = createMockParagraph("段落内容", "Normal");
      mockPara.inlinePictures = { items: [], load: vi.fn() } as any;

      const mockRange = createMockRange("段落内容");
      mockRange.paragraphs = {
        items: [mockPara],
        load: vi.fn(),
      } as any;
      mockRange.tables = { items: [], load: vi.fn() } as any;
      mockRange.contentControls = { items: [], load: vi.fn() } as any;

      mockContext.document.body.paragraphs = {
        items: [mockPara],
        load: vi.fn(),
      };

      mockPara.getRange = vi.fn((location?: any) => {
        if (location === "start" || location === "end") {
          return { expandTo: vi.fn().mockReturnValue(mockRange) };
        }
        return mockRange;
      });

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
        RangeLocation: {
          start: "start",
          end: "end",
        },
      } as any;

      const result = await getRangeContent(
        { type: "paragraph", startIndex: 0 },
        { includeText: false }
      );

      expect(result.text).toBe("");
      expect(result.elements[0].text).toBeUndefined();
    });
  });
});

/**
 * 创建模拟的 Word 上下文
 * Create mock Word context
 */
function createMockContext() {
  return {
    document: {
      body: {
        paragraphs: {
          items: [],
          load: vi.fn(),
        },
        contentControls: {
          items: [],
          load: vi.fn(),
        },
        search: vi.fn().mockReturnValue({
          items: [],
          load: vi.fn(),
        }),
      },
      sections: {
        items: [],
        load: vi.fn(),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建模拟的范围对象
 * Create mock range object
 */
function createMockRange(text: string) {
  return {
    text,
    isEmpty: false,
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
  };
}

/**
 * 创建模拟的段落对象
 * Create mock paragraph object
 */
function createMockParagraph(text: string, style: string) {
  return {
    text,
    style,
    styleBuiltIn: 0,
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
      load: vi.fn(),
    },
    load: vi.fn().mockReturnThis(),
    getRange: vi.fn(),
  };
}
