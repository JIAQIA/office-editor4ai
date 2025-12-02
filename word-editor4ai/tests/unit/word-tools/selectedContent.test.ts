/**
 * 文件名: selectedContent.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: selectedContent 工具的单元测试
 */

import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { getSelectedContent } from "../../../src/word-tools";
import type { ContentInfo } from "../../../src/word-tools/selectedContent";

describe("selectedContent", () => {
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

  describe("getSelectedContent", () => {
    it("应该能够获取选中的文本内容 / Should get selected text content", async () => {
      const mockContext = createMockContext();
      const mockSelection = createMockRange("这是选中的文本内容");

      // 模拟空的集合 / Mock empty collections
      mockSelection.paragraphs = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.tables = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.contentControls = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.isEmpty = false;

      mockContext.document.getSelection = vi.fn().mockReturnValue(mockSelection);

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      const result = await getSelectedContent({
        includeText: true,
        includeImages: false,
        includeTables: false,
        includeContentControls: false,
      });

      expect(result.text).toBe("这是选中的文本内容");
      expect(result.metadata?.isEmpty).toBe(false);
      expect(result.metadata?.characterCount).toBe(9);
    });

    it("应该能够检测空选择 / Should detect empty selection", async () => {
      const mockContext = createMockContext();
      const mockSelection = createMockRange("");

      mockSelection.isEmpty = true;
      mockContext.document.getSelection = vi.fn().mockReturnValue(mockSelection);

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      const result = await getSelectedContent();

      expect(result.text).toBe("");
      expect(result.elements).toHaveLength(0);
      expect(result.metadata?.isEmpty).toBe(true);
      expect(result.metadata?.characterCount).toBe(0);
    });

    it("应该能够获取选中的段落 / Should get selected paragraphs", async () => {
      const mockContext = createMockContext();
      const mockSelection = createMockRange("段落1\n段落2");

      const mockPara1 = createMockParagraph("段落1", "Normal");
      const mockPara2 = createMockParagraph("段落2", "Normal");

      mockPara1.inlinePictures = {
        items: [],
        load: vi.fn(),
      } as any;
      mockPara2.inlinePictures = {
        items: [],
        load: vi.fn(),
      } as any;

      mockSelection.paragraphs = {
        items: [mockPara1, mockPara2],
        load: vi.fn(),
      } as any;
      mockSelection.tables = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.contentControls = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.isEmpty = false;

      mockContext.document.getSelection = vi.fn().mockReturnValue(mockSelection);

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      const result = await getSelectedContent({
        includeText: true,
        includeImages: false,
        includeTables: false,
        includeContentControls: false,
      });

      expect(result.elements).toHaveLength(2);
      expect(result.elements[0].type).toBe("Paragraph");
      expect(result.elements[0].text).toBe("段落1");
      expect(result.elements[1].type).toBe("Paragraph");
      expect(result.elements[1].text).toBe("段落2");
      expect(result.metadata?.paragraphCount).toBe(2);
    });

    it("应该能够获取选中的表格 / Should get selected tables", async () => {
      const mockContext = createMockContext();
      const mockSelection = createMockRange("表格内容");

      const mockTable = {
        rowCount: 2,
        columns: {
          items: [{}, {}],
          load: vi.fn(),
        },
        rows: {
          items: [],
          load: vi.fn(),
        },
        load: vi.fn(),
      };

      mockSelection.paragraphs = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.tables = {
        items: [mockTable],
        load: vi.fn(),
      } as any;
      mockSelection.contentControls = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.isEmpty = false;

      mockContext.document.getSelection = vi.fn().mockReturnValue(mockSelection);

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      const result = await getSelectedContent({
        includeText: true,
        includeImages: false,
        includeTables: true,
        includeContentControls: false,
      });

      expect(result.elements).toHaveLength(1);
      expect(result.elements[0].type).toBe("Table");
      expect((result.elements[0] as any).rowCount).toBe(2);
      expect((result.elements[0] as any).columnCount).toBe(2);
      expect(result.metadata?.tableCount).toBe(1);
    });

    it("应该能够限制文本长度 / Should limit text length", async () => {
      const mockContext = createMockContext();
      const longText = "这是一段很长的文本".repeat(100); // 创建一个很长的文本
      const mockSelection = createMockRange(longText);

      mockSelection.paragraphs = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.tables = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.contentControls = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.isEmpty = false;

      mockContext.document.getSelection = vi.fn().mockReturnValue(mockSelection);

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      const result = await getSelectedContent({
        includeText: true,
        maxTextLength: 100,
      });

      expect(result.text.length).toBeLessThanOrEqual(104); // 100 + "..." (3 characters)
      expect(result.text).toContain("...");
    });

    it("应该能够获取详细的元数据 / Should get detailed metadata", async () => {
      const mockContext = createMockContext();
      const mockSelection = createMockRange("段落内容");

      const mockPara = createMockParagraph("段落内容", "Heading1");
      mockPara.alignment = "Left";
      mockPara.firstLineIndent = 0;
      mockPara.leftIndent = 0;
      mockPara.rightIndent = 0;
      mockPara.lineSpacing = 1.5;
      mockPara.spaceAfter = 10;
      mockPara.spaceBefore = 0;
      mockPara.isListItem = false;
      mockPara.inlinePictures = {
        items: [],
        load: vi.fn(),
      } as any;

      mockSelection.paragraphs = {
        items: [mockPara],
        load: vi.fn(),
      } as any;
      mockSelection.tables = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.contentControls = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.isEmpty = false;

      mockContext.document.getSelection = vi.fn().mockReturnValue(mockSelection);

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      const result = await getSelectedContent({
        includeText: true,
        detailedMetadata: true,
      });

      expect(result.elements).toHaveLength(1);
      const para = result.elements[0] as any;
      expect(para.style).toBe("Heading1");
      expect(para.alignment).toBe("Left");
      expect(para.lineSpacing).toBe(1.5);
    });

    it("应该能够处理包含内联图片的选中内容 / Should handle selection with inline pictures", async () => {
      const mockContext = createMockContext();
      const mockSelection = createMockRange("段落与图片");

      const mockPicture = {
        width: 100,
        height: 100,
        altTextTitle: "测试图片",
        altTextDescription: "这是一张测试图片",
        hyperlink: "",
        load: vi.fn(),
      };

      const mockPara = createMockParagraph("段落与图片", "Normal");
      mockPara.inlinePictures = {
        items: [mockPicture],
        load: vi.fn(),
      } as any;

      mockSelection.paragraphs = {
        items: [mockPara],
        load: vi.fn(),
      } as any;
      mockSelection.tables = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.contentControls = {
        items: [],
        load: vi.fn(),
      } as any;
      mockSelection.isEmpty = false;

      mockContext.document.getSelection = vi.fn().mockReturnValue(mockSelection);

      global.Word = {
        run: vi.fn(async (callback) => callback(mockContext)),
      } as any;

      const result = await getSelectedContent({
        includeText: true,
        includeImages: true,
      });

      expect(result.elements).toHaveLength(2); // 1 段落 + 1 图片
      expect(result.elements[0].type).toBe("Paragraph");
      expect(result.elements[1].type).toBe("InlinePicture");
      expect((result.elements[1] as any).width).toBe(100);
      expect((result.elements[1] as any).altText).toBe("测试图片");
      expect(result.metadata?.imageCount).toBe(1);
    });

    it("应该能够处理错误情况 / Should handle errors", async () => {
      global.Word = {
        run: vi.fn().mockRejectedValue(new Error("API 调用失败")),
      } as any;

      await expect(getSelectedContent()).rejects.toThrow("API 调用失败");
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
      getSelection: vi.fn(),
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
}
