/**
 * 文件名: insertTextBox.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: insertTextBox 工具的单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { insertTextBox, insertTextBoxes } from "../../../src/word-tools";
import type { TextBoxOptions } from "../../../src/word-tools";

// Mock Word API
const mockContext = {
  document: {
    body: {
      getRange: vi.fn(),
    },
    getSelection: vi.fn(),
  },
  sync: vi.fn().mockResolvedValue(undefined),
};

const mockRange = {
  insertTextBox: vi.fn(),
};

const mockTextBox = {
  name: "",
  lockAspectRatio: false,
  visible: true,
  left: 0,
  top: 0,
  rotation: 0,
  id: "test-textbox-id",
  body: {
    getRange: vi.fn(),
  },
  load: vi.fn(),
};

const mockTextRange = {
  font: {
    name: "",
    size: 12,
    bold: false,
    italic: false,
    underline: "None",
    color: "#000000",
    highlightColor: "",
    strikeThrough: false,
    superscript: false,
    subscript: false,
  },
  load: vi.fn(),
};

// Mock Word.run
global.Word = {
  run: vi.fn((callback) => callback(mockContext)),
  UnderlineType: {
    none: "None",
    single: "Single",
    double: "Double",
  },
} as any;

describe("insertTextBox", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockContext.document.body.getRange.mockReturnValue(mockRange);
    mockContext.document.getSelection.mockReturnValue(mockRange);
    mockRange.insertTextBox.mockReturnValue(mockTextBox);
    mockTextBox.body.getRange.mockReturnValue(mockTextRange);

    // 重置 mockTextBox 属性 / Reset mockTextBox properties
    mockTextBox.name = "";
    mockTextBox.lockAspectRatio = false;
    mockTextBox.visible = true;
    mockTextBox.left = 0;
    mockTextBox.top = 0;
    mockTextBox.rotation = 0;

    // 重置 mockTextRange.font 属性 / Reset mockTextRange.font properties
    mockTextRange.font.name = "";
    mockTextRange.font.size = 12;
    mockTextRange.font.bold = false;
    mockTextRange.font.italic = false;
    mockTextRange.font.underline = "None";
    mockTextRange.font.color = "#000000";
    mockTextRange.font.highlightColor = "";
    mockTextRange.font.strikeThrough = false;
    mockTextRange.font.superscript = false;
    mockTextRange.font.subscript = false;
  });

  describe("基本功能 / Basic functionality", () => {
    it("应该成功插入简单文本框 / Should successfully insert simple text box", async () => {
      const result = await insertTextBox("Hello World", "End");

      expect(result.success).toBe(true);
      expect(result.textBoxId).toBe("textbox-test-textbox-id");
      expect(result.error).toBeUndefined();
      expect(mockRange.insertTextBox).toHaveBeenCalledWith("Hello World", {
        width: 150,
        height: 100,
      });
    });

    it("应该在不同位置插入文本框 / Should insert text box at different locations", async () => {
      const locations: Array<"Start" | "End" | "Before" | "After" | "Replace"> = [
        "Start",
        "End",
        "Before",
        "After",
        "Replace",
      ];

      for (const location of locations) {
        await insertTextBox("Test", location);
        expect(mockRange.insertTextBox).toHaveBeenCalledWith("Test", {
          width: 150,
          height: 100,
        });
      }
    });

    it("应该使用默认位置 End / Should use default location End", async () => {
      await insertTextBox("Test");

      expect(mockRange.insertTextBox).toHaveBeenCalledWith("Test", {
        width: 150,
        height: 100,
      });
    });
  });

  describe("参数验证 / Parameter validation", () => {
    it("应该拒绝空文本 / Should reject empty text", async () => {
      const result = await insertTextBox("");

      expect(result.success).toBe(false);
      expect(result.error).toContain("必须提供文本内容");
      expect(mockRange.insertTextBox).not.toHaveBeenCalled();
    });
  });

  describe("文本框选项 / Text box options", () => {
    it("应该设置自定义宽度和高度 / Should set custom width and height", async () => {
      const options: TextBoxOptions = {
        width: 200,
        height: 150,
      };

      await insertTextBox("Test", "End", options);

      expect(mockRange.insertTextBox).toHaveBeenCalledWith("Test", {
        width: 200,
        height: 150,
      });
    });

    it("应该设置文本框名称 / Should set text box name", async () => {
      const options: TextBoxOptions = {
        name: "MyTextBox",
      };

      await insertTextBox("Test", "End", options);

      expect(mockTextBox.name).toBe("MyTextBox");
    });

    it("应该设置锁定纵横比 / Should set lock aspect ratio", async () => {
      const options: TextBoxOptions = {
        lockAspectRatio: true,
      };

      await insertTextBox("Test", "End", options);

      expect(mockTextBox.lockAspectRatio).toBe(true);
    });

    it("应该设置可见性 / Should set visibility", async () => {
      const options: TextBoxOptions = {
        visible: false,
      };

      await insertTextBox("Test", "End", options);

      expect(mockTextBox.visible).toBe(false);
    });

    it("应该设置位置和旋转 / Should set position and rotation", async () => {
      const options: TextBoxOptions = {
        left: 100,
        top: 200,
        rotation: 45,
      };

      await insertTextBox("Test", "End", options);

      expect(mockRange.insertTextBox).toHaveBeenCalledWith("Test", {
        width: 150,
        height: 100,
        left: 100,
        top: 200,
      });
      expect(mockTextBox.rotation).toBe(45);
    });
  });

  describe("文本格式 / Text format", () => {
    it("应该应用字体格式 / Should apply font format", async () => {
      const options: TextBoxOptions = {
        format: {
          fontName: "Arial",
          fontSize: 14,
          bold: true,
          italic: true,
          color: "#FF0000",
        },
      };

      await insertTextBox("Test", "End", options);

      expect(mockTextRange.font.name).toBe("Arial");
      expect(mockTextRange.font.size).toBe(14);
      expect(mockTextRange.font.bold).toBe(true);
      expect(mockTextRange.font.italic).toBe(true);
      expect(mockTextRange.font.color).toBe("#FF0000");
    });

    it("应该应用下划线和删除线 / Should apply underline and strikethrough", async () => {
      const options: TextBoxOptions = {
        format: {
          underline: "Single",
          strikeThrough: true,
        },
      };

      await insertTextBox("Test", "End", options);

      expect(mockTextRange.font.underline).toBe("Single");
      expect(mockTextRange.font.strikeThrough).toBe(true);
    });

    it("应该应用上标和下标 / Should apply superscript and subscript", async () => {
      const options: TextBoxOptions = {
        format: {
          superscript: true,
        },
      };

      await insertTextBox("Test", "End", options);

      expect(mockTextRange.font.superscript).toBe(true);
    });

    it("应该应用高亮颜色 / Should apply highlight color", async () => {
      const options: TextBoxOptions = {
        format: {
          highlightColor: "#FFFF00",
        },
      };

      await insertTextBox("Test", "End", options);

      expect(mockTextRange.font.highlightColor).toBe("#FFFF00");
    });
  });

  describe("错误处理 / Error handling", () => {
    it("应该处理插入失败 / Should handle insert failure", async () => {
      const error = new Error("Insert failed");
      mockRange.insertTextBox.mockImplementationOnce(() => {
        throw error;
      });

      const result = await insertTextBox("Test", "End");

      expect(result.success).toBe(false);
      expect(result.error).toBe("Insert failed");
    });

    it("应该处理格式应用失败 / Should handle format application failure", async () => {
      const options: TextBoxOptions = {
        format: {
          fontName: "Arial",
        },
      };

      mockTextBox.body.getRange.mockImplementationOnce(() => {
        throw new Error("Format failed");
      });

      // 应该仍然成功插入文本框，只是格式应用失败
      // Should still successfully insert text box, just format application failed
      const result = await insertTextBox("Test", "End", options);

      expect(result.success).toBe(true);
    });
  });

  describe("批量插入 / Batch insert", () => {
    it("应该批量插入多个文本框 / Should batch insert multiple text boxes", async () => {
      const textBoxes = [
        { text: "Box 1", location: "End" as const },
        { text: "Box 2", location: "End" as const, options: { width: 200 } },
        { text: "Box 3", location: "Start" as const },
      ];

      const results = await insertTextBoxes(textBoxes);

      expect(results).toHaveLength(3);
      expect(results.every((r) => r.success)).toBe(true);
      expect(mockRange.insertTextBox).toHaveBeenCalledTimes(3);
    });

    it("应该返回每个文本框的结果 / Should return result for each text box", async () => {
      const textBoxes = [
        { text: "Box 1", location: "End" as const },
        { text: "", location: "End" as const }, // 这个会失败 / This will fail
        { text: "Box 3", location: "End" as const },
      ];

      const results = await insertTextBoxes(textBoxes);

      expect(results).toHaveLength(3);
      expect(results[0].success).toBe(true);
      expect(results[1].success).toBe(false);
      expect(results[2].success).toBe(true);
    });
  });

  describe("完整场景 / Complete scenarios", () => {
    it("应该插入完整配置的文本框 / Should insert fully configured text box", async () => {
      const options: TextBoxOptions = {
        width: 250,
        height: 180,
        name: "CompleteTextBox",
        lockAspectRatio: true,
        visible: true,
        left: 50,
        top: 100,
        rotation: 30,
        format: {
          fontName: "Times New Roman",
          fontSize: 16,
          bold: true,
          italic: false,
          underline: "Double",
          color: "#0000FF",
          highlightColor: "#FFFF00",
          strikeThrough: false,
        },
      };

      const result = await insertTextBox("Complete Text Box", "End", options);

      expect(result.success).toBe(true);
      expect(result.textBoxId).toBe("textbox-test-textbox-id");
      // 验证 insertTextBox 被正确调用，包含位置参数 / Verify insertTextBox is called correctly with position parameters
      expect(mockRange.insertTextBox).toHaveBeenCalledWith("Complete Text Box", {
        width: 250,
        height: 180,
        left: 50,
        top: 100,
      });
      // 验证直接设置的属性 / Verify directly set properties
      expect(mockTextBox.name).toBe("CompleteTextBox");
      expect(mockTextBox.lockAspectRatio).toBe(true);
      expect(mockTextBox.rotation).toBe(30);
      // 注意：left 和 top 是通过 insertShapeOptions 传递的，不会直接设置到 textBox 对象上
      // Note: left and top are passed via insertShapeOptions, not directly set on textBox object
    });
  });
});
