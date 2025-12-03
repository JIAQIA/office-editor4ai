/**
 * 文件名: insertShape.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: insertShape 工具的单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { insertShape, insertShapes } from "../../../src/word-tools";
import type { ShapeOptions } from "../../../src/word-tools";

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
  insertShape: vi.fn(),
  insertGeometricShape: vi.fn(),
};

const mockShape = {
  name: "",
  lockAspectRatio: false,
  visible: true,
  rotation: 0,
  id: "test-shape-id",
  fill: {
    setSolidColor: vi.fn(),
  },
  line: {
    color: "",
    weight: 1,
    style: "Single",
  },
  body: {
    getRange: vi.fn(),
  },
  textWrap: {
    type: "",
  },
  load: vi.fn(),
};

const mockTextRange = {
  insertText: vi.fn(),
};

// Mock Word.run
global.Word = {
  run: vi.fn((callback) => callback(mockContext)),
  ShapeType: {
    rectangle: "Rectangle",
    ellipse: "Ellipse",
  },
  ShapeTextWrapType: {
    inline: "inline",
    square: "square",
    tight: "tight",
    through: "through",
    topBottom: "topBottom",
    behind: "behind",
    front: "front",
  },
  ShapeLineStyle: {
    single: "Single",
    dash: "Dash",
    dot: "Dot",
  },
} as any;

describe("insertShape", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockContext.document.body.getRange.mockReturnValue(mockRange);
    mockContext.document.getSelection.mockReturnValue(mockRange);
    mockRange.insertGeometricShape.mockReturnValue(mockShape);
    mockShape.body.getRange.mockReturnValue(mockTextRange);

    // 重置 mockShape 属性 / Reset mockShape properties
    mockShape.name = "";
    mockShape.lockAspectRatio = false;
    mockShape.visible = true;
    mockShape.rotation = 0;
    mockShape.line.color = "";
    mockShape.line.weight = 1;
    mockShape.line.style = "Single";
  });

  describe("基本功能 / Basic functionality", () => {
    it("应该成功插入简单形状 / Should successfully insert simple shape", async () => {
      const result = await insertShape("Rectangle", "End");

      expect(result.success).toBe(true);
      expect(result.shapeId).toBe("shape-test-shape-id");
      expect(result.error).toBeUndefined();
      expect(mockRange.insertGeometricShape).toHaveBeenCalledWith("Rectangle", {
        width: 100,
        height: 100,
      });
    });

    it("应该在不同位置插入形状 / Should insert shape at different locations", async () => {
      const locations: Array<"Start" | "End" | "Before" | "After" | "Replace"> = [
        "Start",
        "End",
        "Before",
        "After",
        "Replace",
      ];

      for (const location of locations) {
        await insertShape("Ellipse", location);
        expect(mockRange.insertGeometricShape).toHaveBeenCalledWith("Ellipse", {
          width: 100,
          height: 100,
        });
      }
    });

    it("应该使用默认位置 End / Should use default location End", async () => {
      await insertShape("Diamond");

      expect(mockRange.insertGeometricShape).toHaveBeenCalledWith("Diamond", {
        width: 100,
        height: 100,
      });
    });

    it("应该支持不同的形状类型 / Should support different shape types", async () => {
      const shapeTypes = ["Rectangle", "Ellipse", "Diamond", "Triangle", "Star"];

      for (const shapeType of shapeTypes) {
        await insertShape(shapeType, "End");
        expect(mockRange.insertGeometricShape).toHaveBeenCalledWith(shapeType, {
          width: 100,
          height: 100,
        });
      }
    });
  });

  describe("参数验证 / Parameter validation", () => {
    it("应该拒绝空形状类型 / Should reject empty shape type", async () => {
      const result = await insertShape("");

      expect(result.success).toBe(false);
      expect(result.error).toContain("必须提供形状类型");
      expect(mockRange.insertGeometricShape).not.toHaveBeenCalled();
    });
  });

  describe("形状选项 / Shape options", () => {
    it("应该设置自定义宽度和高度 / Should set custom width and height", async () => {
      const options: ShapeOptions = {
        width: 200,
        height: 150,
      };

      await insertShape("Rectangle", "End", options);

      expect(mockRange.insertGeometricShape).toHaveBeenCalledWith("Rectangle", {
        width: 200,
        height: 150,
      });
    });

    it("应该设置形状名称 / Should set shape name", async () => {
      const options: ShapeOptions = {
        name: "MyShape",
      };

      await insertShape("Ellipse", "End", options);

      expect(mockShape.name).toBe("MyShape");
    });

    it("应该设置锁定纵横比 / Should set lock aspect ratio", async () => {
      const options: ShapeOptions = {
        lockAspectRatio: true,
      };

      await insertShape("Rectangle", "End", options);

      expect(mockShape.lockAspectRatio).toBe(true);
    });

    it("应该设置可见性 / Should set visibility", async () => {
      const options: ShapeOptions = {
        visible: false,
      };

      await insertShape("Rectangle", "End", options);

      expect(mockShape.visible).toBe(false);
    });

    it("应该设置位置和旋转 / Should set position and rotation", async () => {
      const options: ShapeOptions = {
        left: 100,
        top: 200,
        rotation: 45,
      };

      await insertShape("Rectangle", "End", options);

      expect(mockRange.insertGeometricShape).toHaveBeenCalledWith("Rectangle", {
        width: 100,
        height: 100,
        left: 100,
        top: 200,
      });
      expect(mockShape.rotation).toBe(45);
    });
  });

  describe("样式选项 / Style options", () => {
    it("应该设置填充颜色 / Should set fill color", async () => {
      const options: ShapeOptions = {
        fillColor: "#FF0000",
      };

      await insertShape("Rectangle", "End", options);

      expect(mockShape.fill.setSolidColor).toHaveBeenCalledWith("#FF0000");
    });

    it("应该忽略线条颜色（API 不支持）/ Should ignore line color (API not supported)", async () => {
      const consoleSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
      const options: ShapeOptions = {
        lineColor: "#0000FF",
      };

      await insertShape("Rectangle", "End", options);

      // 线条颜色不应该被设置 / Line color should not be set
      expect(mockShape.line.color).toBe("");
      // 应该输出警告 / Should output warning
      expect(consoleSpy).toHaveBeenCalledWith(
        expect.stringContaining("线条样式设置暂不支持")
      );
      consoleSpy.mockRestore();
    });

    it("应该忽略线条宽度（API 不支持）/ Should ignore line weight (API not supported)", async () => {
      const consoleSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
      const options: ShapeOptions = {
        lineWeight: 3,
      };

      await insertShape("Rectangle", "End", options);

      // 线条宽度不应该被设置 / Line weight should not be set
      expect(mockShape.line.weight).toBe(1);
      // 应该输出警告 / Should output warning
      expect(consoleSpy).toHaveBeenCalledWith(
        expect.stringContaining("线条样式设置暂不支持")
      );
      consoleSpy.mockRestore();
    });

    it("应该忽略线条样式（API 不支持）/ Should ignore line style (API not supported)", async () => {
      const consoleSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
      const options: ShapeOptions = {
        lineStyle: "Dash",
      };

      await insertShape("Rectangle", "End", options);

      // 线条样式不应该被设置 / Line style should not be set
      expect(mockShape.line.style).toBe("Single");
      // 应该输出警告 / Should output warning
      expect(consoleSpy).toHaveBeenCalledWith(
        expect.stringContaining("线条样式设置暂不支持")
      );
      consoleSpy.mockRestore();
    });

    it("应该设置填充但忽略线条样式 / Should set fill but ignore line styles", async () => {
      const consoleSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
      const options: ShapeOptions = {
        fillColor: "#FF0000",
        lineColor: "#0000FF",
        lineWeight: 2,
        lineStyle: "Dot",
      };

      await insertShape("Ellipse", "End", options);

      // 填充颜色应该被设置 / Fill color should be set
      expect(mockShape.fill.setSolidColor).toHaveBeenCalledWith("#FF0000");
      // 线条样式不应该被设置 / Line styles should not be set
      expect(mockShape.line.color).toBe("");
      expect(mockShape.line.weight).toBe(1);
      expect(mockShape.line.style).toBe("Single");
      // 应该输出警告 / Should output warning
      expect(consoleSpy).toHaveBeenCalledWith(
        expect.stringContaining("线条样式设置暂不支持")
      );
      consoleSpy.mockRestore();
    });
  });

  describe("文本内容 / Text content", () => {
    it("应该添加文本内容 / Should add text content", async () => {
      const options: ShapeOptions = {
        text: "Hello World",
      };

      await insertShape("RoundRectangle", "End", options);

      expect(mockTextRange.insertText).toHaveBeenCalledWith("Hello World", "Replace");
    });

    it("应该处理文本添加失败 / Should handle text addition failure", async () => {
      const options: ShapeOptions = {
        text: "Test",
      };

      mockShape.body.getRange.mockImplementationOnce(() => {
        throw new Error("Text failed");
      });

      // 应该仍然成功插入形状，只是文本添加失败
      // Should still successfully insert shape, just text addition failed
      const result = await insertShape("Rectangle", "End", options);

      expect(result.success).toBe(true);
    });
  });

  describe("错误处理 / Error handling", () => {
    it("应该处理插入失败 / Should handle insert failure", async () => {
      const error = new Error("Insert failed");
      mockRange.insertGeometricShape.mockImplementationOnce(() => {
        throw error;
      });

      const result = await insertShape("Rectangle", "End");

      expect(result.success).toBe(false);
      expect(result.error).toBe("Insert failed");
    });

    it("应该处理样式应用失败 / Should handle style application failure", async () => {
      const options: ShapeOptions = {
        fillColor: "#FF0000",
      };

      mockShape.fill.setSolidColor.mockImplementationOnce(() => {
        throw new Error("Style failed");
      });

      // 应该仍然成功插入形状，只是样式应用失败
      // Should still successfully insert shape, just style application failed
      const result = await insertShape("Rectangle", "End", options);

      expect(result.success).toBe(true);
    });
  });

  describe("批量插入 / Batch insert", () => {
    it("应该批量插入多个形状 / Should batch insert multiple shapes", async () => {
      const shapes = [
        { shapeType: "Rectangle", location: "End" as const },
        { shapeType: "Ellipse", location: "End" as const, options: { width: 200 } },
        { shapeType: "Diamond", location: "Start" as const },
      ];

      const results = await insertShapes(shapes);

      expect(results).toHaveLength(3);
      expect(results.every((r) => r.success)).toBe(true);
      expect(mockRange.insertGeometricShape).toHaveBeenCalledTimes(3);
    });

    it("应该返回每个形状的结果 / Should return result for each shape", async () => {
      const shapes = [
        { shapeType: "Rectangle", location: "End" as const },
        { shapeType: "", location: "End" as const }, // 这个会失败 / This will fail
        { shapeType: "Ellipse", location: "End" as const },
      ];

      const results = await insertShapes(shapes);

      expect(results).toHaveLength(3);
      expect(results[0].success).toBe(true);
      expect(results[1].success).toBe(false);
      expect(results[2].success).toBe(true);
    });
  });

  describe("完整场景 / Complete scenarios", () => {
    it("应该插入完整配置的形状 / Should insert fully configured shape", async () => {
      const options: ShapeOptions = {
        width: 250,
        height: 180,
        name: "CompleteShape",
        lockAspectRatio: true,
        visible: true,
        left: 50,
        top: 100,
        rotation: 30,
        fillColor: "#FF0000",
        lineColor: "#0000FF",
        lineWeight: 2,
        lineStyle: "Dash",
        text: "Complete Shape",
      };

      const result = await insertShape("RoundRectangle", "End", options);

      expect(result.success).toBe(true);
      expect(result.shapeId).toBe("shape-test-shape-id");
      // 验证 insertShape 被正确调用，包含位置参数 / Verify insertShape is called correctly with position parameters
      expect(mockRange.insertGeometricShape).toHaveBeenCalledWith("RoundRectangle", {
        width: 250,
        height: 180,
        left: 50,
        top: 100,
      });
      // 验证直接设置的属性 / Verify directly set properties
      expect(mockShape.name).toBe("CompleteShape");
      expect(mockShape.lockAspectRatio).toBe(true);
      expect(mockShape.rotation).toBe(30);
      // 验证样式 / Verify styles
      expect(mockShape.fill.setSolidColor).toHaveBeenCalledWith("#FF0000");
      // 线条样式不支持，应保持默认值 / Line styles not supported, should keep default values
      expect(mockShape.line.color).toBe("");
      expect(mockShape.line.weight).toBe(1);
      expect(mockShape.line.style).toBe("Single");
      // 验证文本 / Verify text
      expect(mockTextRange.insertText).toHaveBeenCalledWith("Complete Shape", "Replace");
    });

    it("应该插入带样式的圆形 / Should insert styled circle", async () => {
      const options: ShapeOptions = {
        width: 150,
        height: 150,
        fillColor: "#0078D4",
        lineColor: "#000000",
        lineWeight: 3,
      };

      const result = await insertShape("Ellipse", "End", options);

      expect(result.success).toBe(true);
      expect(mockRange.insertGeometricShape).toHaveBeenCalledWith("Ellipse", {
        width: 150,
        height: 150,
      });
      expect(mockShape.fill.setSolidColor).toHaveBeenCalledWith("#0078D4");
      // 线条样式不支持，应保持默认值 / Line styles not supported, should keep default values
      expect(mockShape.line.color).toBe("");
      expect(mockShape.line.weight).toBe(1);
    });

    it("应该插入流程图形状 / Should insert flowchart shape", async () => {
      const options: ShapeOptions = {
        width: 200,
        height: 100,
        text: "Process",
        fillColor: "#E1F5FE",
        lineColor: "#01579B",
        lineWeight: 2,
      };

      const result = await insertShape("FlowChartProcess", "End", options);

      expect(result.success).toBe(true);
      expect(mockRange.insertGeometricShape).toHaveBeenCalledWith("FlowChartProcess", {
        width: 200,
        height: 100,
      });
      expect(mockTextRange.insertText).toHaveBeenCalledWith("Process", "Replace");
      expect(mockShape.fill.setSolidColor).toHaveBeenCalledWith("#E1F5FE");
      // 线条样式不支持，应保持默认值 / Line styles not supported, should keep default values
      expect(mockShape.line.color).toBe("");
    });
  });

  describe("文字环绕 / Text wrapping", () => {
    beforeEach(() => {
      // 重置 textWrap mock / Reset textWrap mock
      mockShape.textWrap.type = "";
    });

    it("应该设置四周型环绕 / Should set square wrapping", async () => {
      const options: ShapeOptions = {
        wrapType: "Square",
      };

      await insertShape("Rectangle", "End", options);

      expect(mockShape.textWrap.type).toBe("square");
    });

    it("应该设置紧密型环绕 / Should set tight wrapping", async () => {
      const options: ShapeOptions = {
        wrapType: "Tight",
      };

      await insertShape("Ellipse", "End", options);

      expect(mockShape.textWrap.type).toBe("tight");
    });

    it("应该设置上下型环绕 / Should set top-bottom wrapping", async () => {
      const options: ShapeOptions = {
        wrapType: "TopBottom",
      };

      await insertShape("Triangle", "End", options);

      expect(mockShape.textWrap.type).toBe("topBottom");
    });

    it("应该设置浮于文字上方 / Should set in front of text", async () => {
      const options: ShapeOptions = {
        wrapType: "Front",
      };

      await insertShape("Rectangle", "End", options);

      expect(mockShape.textWrap.type).toBe("front");
    });

    it("应该设置衬于文字下方 / Should set behind text", async () => {
      const options: ShapeOptions = {
        wrapType: "Behind",
      };

      await insertShape("Rectangle", "End", options);

      expect(mockShape.textWrap.type).toBe("behind");
    });

    it("应该同时设置样式和文字环绕 / Should set both style and text wrapping", async () => {
      const options: ShapeOptions = {
        width: 150,
        height: 100,
        fillColor: "#FF5722",
        wrapType: "Square",
        text: "Wrapped Shape",
      };

      await insertShape("RoundRectangle", "End", options);

      expect(mockShape.fill.setSolidColor).toHaveBeenCalledWith("#FF5722");
      expect(mockShape.textWrap.type).toBe("square");
      expect(mockTextRange.insertText).toHaveBeenCalledWith("Wrapped Shape", "Replace");
    });

    it("未设置 wrapType 时不应修改文字环绕 / Should not modify text wrapping when wrapType is not set", async () => {
      const options: ShapeOptions = {
        width: 100,
        height: 100,
      };

      await insertShape("Rectangle", "End", options);

      expect(mockShape.textWrap.type).toBe("");
    });
  });
});
