/**
 * 文件名: textBoxContent.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 文本框内容获取工具的测试文件 | Test file for text box content tool
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { getTextBoxes } from "../../../src/word-tools";

describe("textBoxContent - 文本框内容获取工具", () => {
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

  describe("getTextBoxes - 获取文本框内容", () => {
    it("应该在没有选择时返回可见区域的文本框 | Should return text boxes in visible area when no selection", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 2);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getTextBoxes();

      expect(result).toBeDefined();
      expect(Array.isArray(result)).toBe(true);
      expect(result.length).toBe(2);
    });

    it("应该在有选择时返回选择范围内的文本框 | Should return text boxes in selection when selection exists", async () => {
      const mockContext = createMockContextWithTextBoxes(true, 3);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getTextBoxes();

      expect(result).toBeDefined();
      expect(result.length).toBe(3);
    });

    it("应该在没有文本框时返回空数组 | Should return empty array when no text boxes found", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 0);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getTextBoxes();

      expect(result).toBeDefined();
      expect(result.length).toBe(0);
    });

    it("应该支持 includeText 选项 | Should support includeText option", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithText = await getTextBoxes({ includeText: true });
      const resultWithoutText = await getTextBoxes({ includeText: false });

      expect(resultWithText[0].text).toBeDefined();
      expect(resultWithoutText[0].text).toBeUndefined();
    });

    it("应该支持 includeParagraphs 选项 | Should support includeParagraphs option", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithParagraphs = await getTextBoxes({ includeParagraphs: true });
      const resultWithoutParagraphs = await getTextBoxes({ includeParagraphs: false });

      expect(resultWithParagraphs[0].paragraphs).toBeDefined();
      expect(resultWithParagraphs[0].paragraphs!.length).toBeGreaterThan(0);
      expect(resultWithoutParagraphs[0].paragraphs).toBeUndefined();
    });

    it("应该支持 detailedMetadata 选项 | Should support detailedMetadata option", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithMetadata = await getTextBoxes({ detailedMetadata: true });
      const resultWithoutMetadata = await getTextBoxes({ detailedMetadata: false });

      expect(resultWithMetadata[0].name).toBeDefined();
      expect(resultWithMetadata[0].width).toBeDefined();
      expect(resultWithMetadata[0].height).toBeDefined();
      expect(resultWithoutMetadata[0].name).toBeUndefined();
      expect(resultWithoutMetadata[0].width).toBeUndefined();
    });

    it("应该支持 maxTextLength 选项 | Should support maxTextLength option", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getTextBoxes({ includeText: true, maxTextLength: 10 });

      expect(result[0].text).toBeDefined();
      expect(result[0].text!.length).toBeLessThanOrEqual(13); // 10 + "..."
    });

    it("应该正确处理文本框属性 | Should handle text box properties correctly", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getTextBoxes({ detailedMetadata: true });

      expect(result[0]).toHaveProperty("id");
      expect(result[0]).toHaveProperty("name");
      expect(result[0]).toHaveProperty("width");
      expect(result[0]).toHaveProperty("height");
      expect(result[0]).toHaveProperty("left");
      expect(result[0]).toHaveProperty("top");
      expect(result[0]).toHaveProperty("rotation");
      expect(result[0]).toHaveProperty("visible");
      expect(result[0]).toHaveProperty("lockAspectRatio");
    });

    it("应该正确处理段落详情 | Should handle paragraph details correctly", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getTextBoxes({
        includeParagraphs: true,
        detailedMetadata: true,
      });

      expect(result[0].paragraphs).toBeDefined();
      expect(result[0].paragraphs!.length).toBeGreaterThan(0);

      const paragraph = result[0].paragraphs![0];
      expect(paragraph).toHaveProperty("id");
      expect(paragraph).toHaveProperty("type");
      expect(paragraph).toHaveProperty("text");
      expect(paragraph.type).toBe("Paragraph");
    });

    it("应该正确处理错误 | Should handle errors correctly", async () => {
      global.Word.run = vi.fn(async () => {
        throw new Error("API Error");
      });

      await expect(getTextBoxes()).rejects.toThrow();
    });

    it("应该正确处理多个文本框 | Should handle multiple text boxes correctly", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 5);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getTextBoxes({ includeText: true });

      expect(result.length).toBe(5);
      result.forEach((textBox, index) => {
        expect(textBox.id).toBe(`textbox-${index + 1}`);
        expect(textBox.text).toBeDefined();
      });
    });
  });

  describe("TextBoxInfo 类型验证 | TextBoxInfo type validation", () => {
    it("应该包含所有必需的字段 | Should contain all required fields", async () => {
      const mockContext = createMockContextWithTextBoxes(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getTextBoxes({ detailedMetadata: true, includeText: true });

      expect(result[0]).toHaveProperty("id");
      expect(typeof result[0].id).toBe("string");
    });
  });
});

/**
 * 创建包含文本框的模拟上下文
 * Create mock context with text boxes
 */
function createMockContextWithTextBoxes(hasSelection: boolean, textBoxCount: number) {
  // 创建模拟的文本框形状 / Create mock text box shapes
  const mockTextBoxes = Array.from({ length: textBoxCount }, (_, i) => {
    const mockParagraph = {
      text: `Text box ${i + 1} paragraph content`,
      style: "Normal",
      alignment: "Left",
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
      lineSpacing: 1.5,
      spaceAfter: 10,
      spaceBefore: 0,
      isListItem: false,
      load: vi.fn().mockReturnThis(),
    };

    return {
      id: i + 1,
      name: `TextBox${i + 1}`,
      type: Word.ShapeType.textBox,
      width: 200,
      height: 100,
      left: 50 + i * 20,
      top: 50 + i * 20,
      rotation: 0,
      visible: true,
      lockAspectRatio: false,
      body: {
        text: `Text box ${i + 1} content`,
        paragraphs: {
          items: [mockParagraph],
          load: vi.fn().mockReturnThis(),
        },
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    };
  });

  // 创建模拟的选择范围 / Create mock selection range
  const mockSelection = {
    text: hasSelection ? "some text" : "",
    isEmpty: !hasSelection,
    shapes: {
      items: hasSelection ? mockTextBoxes : [],
      load: vi.fn().mockReturnThis(),
    },
    load: vi.fn().mockReturnThis(),
    getRange: vi.fn().mockReturnValue({
      getRange: vi.fn().mockReturnValue({
        body: {
          shapes: {
            items: mockTextBoxes,
            load: vi.fn().mockReturnThis(),
          },
        },
      }),
    }),
  };

  // 创建模拟的可见页面 / Create mock visible pages
  const mockPage = {
    getRange: vi.fn().mockReturnValue({
      getRange: vi.fn().mockReturnValue({
        getRange: vi.fn().mockReturnValue({
          body: {
            shapes: {
              items: mockTextBoxes,
              load: vi.fn().mockReturnThis(),
            },
          },
        }),
      }),
    }),
    load: vi.fn().mockReturnThis(),
  };

  return {
    document: {
      getSelection: vi.fn().mockReturnValue(mockSelection),
      body: {
        shapes: {
          items: mockTextBoxes,
          load: vi.fn().mockReturnThis(),
        },
      },
      activeWindow: {
        activePane: {
          pagesEnclosingViewport: {
            items: [mockPage],
            load: vi.fn().mockReturnThis(),
          },
        },
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}
