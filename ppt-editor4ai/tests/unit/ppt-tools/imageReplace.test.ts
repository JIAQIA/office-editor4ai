/**
 * 文件名: imageReplace.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: imageReplace 工具的单元测试 | imageReplace tool unit tests
 */

import { describe, it, expect, beforeEach } from "vitest";
import { replaceImage, replaceSelectedImage, getImageInfo } from "../../../src/ppt-tools";

type MockShape = {
  id: string;
  type: string;
  name: string;
  left: number;
  top: number;
  width: number;
  height: number;
  placeholderFormat?: {
    type: string;
    load: () => void;
  };
  fill: {
    type: string;
    foregroundColor?: string;
    setSolidColor: (color: string) => void;
    setPictureRelativeToOriginalSize: (scale: number) => void;
    load: () => void;
  };
  delete: () => void;
  load: () => void;
};

type MockData = {
  context: {
    presentation: {
      getSelectedSlides: () => {
        items: Array<{
          shapes: {
            items: MockShape[];
            addImage: (base64: string) => MockShape;
            load: () => void;
          };
          load: () => void;
        }>;
        load: () => void;
      };
      getSelectedShapes: () => {
        items: MockShape[];
        load: () => void;
      };
    };
    sync: () => Promise<void>;
  };
  run: (callback: (context: MockData["context"]) => Promise<void>) => Promise<void>;
  _shapes: MockShape[];
  _selectedShapeIds: string[];
  _deletedShapeIds: string[];
  _addedImages: Array<{ base64: string; shape: MockShape }>;
};

// 创建 mock 图片对象
const createMockShape = (
  id: string,
  type: "Picture" | "Placeholder" | "TextBox",
  name: string,
  placeholderType?: string
): MockShape => {
  const shape: MockShape = {
    id,
    type,
    name,
    left: 100,
    top: 100,
    width: 200,
    height: 150,
    fill: {
      type: "Solid",
      foregroundColor: "#FFFFFF",
      setSolidColor: function (_color: string) {
        this.type = "Solid";
      },
      setPictureRelativeToOriginalSize: function (_scale: number) {
        this.type = "Picture";
      },
      load: function () {},
    },
    delete: function () {
      // Will be handled by mockData
    },
    load: function () {},
  };

  if (type === "Placeholder" && placeholderType) {
    shape.placeholderFormat = {
      type: placeholderType,
      load: function () {},
    };
  }

  return shape;
};

// 创建 mock 数据
const createMockData = (): MockData => {
  const shapes: MockShape[] = [
    createMockShape("shape-1", "Picture", "Image 1"),
    createMockShape("shape-2", "Placeholder", "Placeholder 1", "Picture"),
    createMockShape("shape-3", "TextBox", "Text 1"),
    createMockShape("shape-4", "Picture", "Image 2"),
  ];

  let selectedShapeIds: string[] = ["shape-1"];
  const deletedShapeIds: string[] = [];
  const addedImages: Array<{ base64: string; shape: MockShape }> = [];

  const mockData: MockData = {
    context: {
      presentation: {
        getSelectedSlides: function () {
          return {
            items: [
              {
                shapes: {
                  items: shapes.filter((s) => !deletedShapeIds.includes(s.id)),
                  addImage: function (base64: string) {
                    const newShape = createMockShape(
                      `new-shape-${addedImages.length + 1}`,
                      "Picture",
                      "New Image"
                    );
                    addedImages.push({ base64, shape: newShape });
                    shapes.push(newShape);
                    return newShape;
                  },
                  load: function () {},
                },
                load: function () {},
              },
            ],
            load: function () {},
          };
        },
        getSelectedShapes: function () {
          const selected = shapes.filter((s) => selectedShapeIds.includes(s.id));
          return {
            items: selected,
            load: function () {},
          };
        },
      },
      sync: async function () {
        // Mock sync - do nothing
      },
    },
    run: async function (callback) {
      await callback(this.context);
    },
    _shapes: shapes,
    _selectedShapeIds: selectedShapeIds,
    _deletedShapeIds: deletedShapeIds,
    _addedImages: addedImages,
  };

  // Override delete to track deleted shapes
  shapes.forEach((shape) => {
    shape.delete = function () {
      if (!deletedShapeIds.includes(shape.id)) {
        deletedShapeIds.push(shape.id);
      }
    };
  });

  return mockData;
};

describe("imageReplace", () => {
  let mockData: MockData;

  beforeEach(() => {
    mockData = createMockData();
    // @ts-expect-error - Mock PowerPoint global
    global.PowerPoint = mockData;
  });

  describe("replaceImage", () => {
    it("应该成功替换选中的图片", async () => {
      const result = await replaceImage({
        imageSource: "data:image/png;base64,iVBORw0KGgoAAAANS...",
        keepDimensions: true,
      });

      expect(result.success).toBe(true);
      expect(result.message).toContain("图片替换成功");
      expect(result.elementId).toBeDefined();
      expect(result.elementType).toBe("Picture");
      expect(result.originalDimensions).toMatchObject({
        left: 100,
        top: 100,
        width: 200,
        height: 150,
      });
    });

    it("应该成功替换指定ID的图片", async () => {
      const result = await replaceImage({
        elementId: "shape-4",
        imageSource: "iVBORw0KGgoAAAANS...",
        keepDimensions: true,
      });

      expect(result.success).toBe(true);
      expect(result.message).toContain("图片替换成功");
    });

    it("应该成功替换占位符图片", async () => {
      mockData._selectedShapeIds = ["shape-2"];

      const result = await replaceImage({
        imageSource: "data:image/png;base64,iVBORw0KGgoAAAANS...",
        keepDimensions: true,
      });

      expect(result.success).toBe(true);
      expect(result.message).toContain("图片替换成功");
    });

    it("应该支持自定义尺寸", async () => {
      const result = await replaceImage({
        imageSource: "iVBORw0KGgoAAAANS...",
        keepDimensions: false,
        width: 300,
        height: 250,
      });

      expect(result.success).toBe(true);
      expect(mockData._addedImages[0].shape.width).toBe(300);
      expect(mockData._addedImages[0].shape.height).toBe(250);
    });

    it("应该拒绝空的图片数据", async () => {
      const result = await replaceImage({
        imageSource: "",
      });

      expect(result.success).toBe(false);
      expect(result.message).toContain("图片数据不能为空");
    });

    it("应该拒绝未找到的元素ID", async () => {
      const result = await replaceImage({
        elementId: "non-existent",
        imageSource: "iVBORw0KGgoAAAANS...",
      });

      expect(result.success).toBe(false);
      expect(result.message).toContain("未找到ID为 non-existent 的元素");
    });

    it("应该拒绝非图片元素", async () => {
      mockData._selectedShapeIds = ["shape-3"]; // TextBox

      const result = await replaceImage({
        imageSource: "iVBORw0KGgoAAAANS...",
      });

      expect(result.success).toBe(false);
      expect(result.message).toContain("不支持图片替换");
    });

    it("应该拒绝未选中任何元素", async () => {
      mockData._selectedShapeIds = [];

      const result = await replaceImage({
        imageSource: "iVBORw0KGgoAAAANS...",
      });

      expect(result.success).toBe(false);
      expect(result.message).toContain("请先选中一个图片元素");
    });

    it("应该拒绝选中多个元素", async () => {
      mockData._selectedShapeIds = ["shape-1", "shape-4"];

      const result = await replaceImage({
        imageSource: "iVBORw0KGgoAAAANS...",
      });

      expect(result.success).toBe(false);
      expect(result.message).toContain("请只选中一个图片元素");
    });

    it("应该正确处理带 data URL 前缀的 Base64", async () => {
      const result = await replaceImage({
        imageSource: "data:image/png;base64,iVBORw0KGgoAAAANS...",
      });

      expect(result.success).toBe(true);
      expect(mockData._addedImages[0].base64).toBe("iVBORw0KGgoAAAANS...");
    });

    it("应该正确处理不带 data URL 前缀的 Base64", async () => {
      const result = await replaceImage({
        imageSource: "iVBORw0KGgoAAAANS...",
      });

      expect(result.success).toBe(true);
      expect(mockData._addedImages[0].base64).toBe("iVBORw0KGgoAAAANS...");
    });

    it("应该删除旧图片并添加新图片", async () => {
      const result = await replaceImage({
        imageSource: "iVBORw0KGgoAAAANS...",
      });

      expect(result.success).toBe(true);
      expect(mockData._deletedShapeIds).toContain("shape-1");
      expect(mockData._addedImages).toHaveLength(1);
    });
  });

  describe("replaceSelectedImage", () => {
    it("应该成功替换当前选中的图片", async () => {
      const result = await replaceSelectedImage("iVBORw0KGgoAAAANS...", true);

      expect(result.success).toBe(true);
      expect(result.message).toContain("图片替换成功");
    });

    it("应该支持不保持尺寸", async () => {
      const result = await replaceSelectedImage("iVBORw0KGgoAAAANS...", false);

      expect(result.success).toBe(true);
    });
  });

  describe("getImageInfo", () => {
    it("应该返回选中图片的信息", async () => {
      const info = await getImageInfo();

      expect(info).not.toBeNull();
      expect(info?.elementId).toBe("shape-1");
      expect(info?.elementType).toBe("Picture");
      expect(info?.name).toBe("Image 1");
      expect(info?.left).toBe(100);
      expect(info?.top).toBe(100);
      expect(info?.width).toBe(200);
      expect(info?.height).toBe(150);
      expect(info?.isPlaceholder).toBe(false);
    });

    it("应该返回指定ID图片的信息", async () => {
      const info = await getImageInfo("shape-4");

      expect(info).not.toBeNull();
      expect(info?.elementId).toBe("shape-4");
      expect(info?.elementType).toBe("Picture");
      expect(info?.name).toBe("Image 2");
    });

    it("应该返回占位符图片的信息", async () => {
      mockData._selectedShapeIds = ["shape-2"];

      const info = await getImageInfo();

      expect(info).not.toBeNull();
      expect(info?.elementId).toBe("shape-2");
      expect(info?.elementType).toBe("Placeholder");
      expect(info?.isPlaceholder).toBe(true);
      expect(info?.placeholderType).toBe("Picture");
    });

    it("应该处理未选中任何元素的情况", async () => {
      mockData._selectedShapeIds = [];

      const info = await getImageInfo();

      expect(info).toBeNull();
    });

    it("应该处理未找到指定ID的情况", async () => {
      const info = await getImageInfo("non-existent");

      expect(info).toBeNull();
    });
  });

  describe("边界情况", () => {
    it("应该处理极小的尺寸", async () => {
      const result = await replaceImage({
        imageSource: "iVBORw0KGgoAAAANS...",
        keepDimensions: false,
        width: 1,
        height: 1,
      });

      expect(result.success).toBe(true);
      expect(mockData._addedImages[0].shape.width).toBe(1);
      expect(mockData._addedImages[0].shape.height).toBe(1);
    });

    it("应该处理极大的尺寸", async () => {
      const result = await replaceImage({
        imageSource: "iVBORw0KGgoAAAANS...",
        keepDimensions: false,
        width: 10000,
        height: 10000,
      });

      expect(result.success).toBe(true);
      expect(mockData._addedImages[0].shape.width).toBe(10000);
      expect(mockData._addedImages[0].shape.height).toBe(10000);
    });

    it("应该保持原图片的名称", async () => {
      const result = await replaceImage({
        imageSource: "iVBORw0KGgoAAAANS...",
      });

      expect(result.success).toBe(true);
      expect(mockData._addedImages[0].shape.name).toBe("Image 1");
    });
  });
});
