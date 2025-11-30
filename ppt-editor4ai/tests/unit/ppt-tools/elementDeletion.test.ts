/**
 * 文件名: elementDeletion.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 元素删除功能单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import {
  deleteElementById,
  deleteElementByName,
  deleteElementByIndex,
  deleteElement,
  deleteElementsByIds,
} from "../../../src/ppt-tools/elementDeletion";

// Mock PowerPoint API
const mockPowerPoint = {
  run: vi.fn(),
};

global.PowerPoint = mockPowerPoint as any;

describe("elementDeletion", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("deleteElementById", () => {
    it("应该成功通过ID删除元素", async () => {
      const mockShape1 = {
        id: "shape-123",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShape2 = {
        id: "shape-456",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShapes = {
        items: [mockShape1, mockShape2],
        load: vi.fn(),
      };

      const mockSlide = {
        shapes: mockShapes,
      };

      const mockSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockSelectedSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn().mockReturnValue(mockSelectedSlides),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      mockPowerPoint.run.mockImplementation(async (callback: any) => {
        await callback(mockContext);
      });

      const result = await deleteElementById("shape-123");

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(mockShape1.delete).toHaveBeenCalledTimes(1);
      expect(mockShape2.delete).not.toHaveBeenCalled();
    });

    it("应该在找不到元素时返回失败", async () => {
      const mockShape = {
        id: "shape-123",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShapes = {
        items: [mockShape],
        load: vi.fn(),
      };

      const mockSlide = {
        shapes: mockShapes,
      };

      const mockSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockSelectedSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn().mockReturnValue(mockSelectedSlides),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      mockPowerPoint.run.mockImplementation(async (callback: any) => {
        await callback(mockContext);
      });

      const result = await deleteElementById("non-existent-id");

      expect(result.success).toBe(false);
      expect(result.deletedCount).toBe(0);
      expect(result.message).toContain("未找到");
    });

    it("应该支持指定幻灯片页码", async () => {
      const mockShape = {
        id: "shape-123",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShapes = {
        items: [mockShape],
        load: vi.fn(),
      };

      const mockSlide1 = {
        shapes: mockShapes,
      };

      const mockSlide2 = {
        shapes: { items: [], load: vi.fn() },
      };

      const mockSlides = {
        items: [mockSlide1, mockSlide2],
        load: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      mockPowerPoint.run.mockImplementation(async (callback: any) => {
        await callback(mockContext);
      });

      const result = await deleteElementById("shape-123", 1);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
    });
  });

  describe("deleteElementByName", () => {
    it("应该成功通过名称删除元素", async () => {
      const mockShape1 = {
        name: "TextBox 1",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShape2 = {
        name: "TextBox 2",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShapes = {
        items: [mockShape1, mockShape2],
        load: vi.fn(),
      };

      const mockSlide = {
        shapes: mockShapes,
      };

      const mockSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockSelectedSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn().mockReturnValue(mockSelectedSlides),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      mockPowerPoint.run.mockImplementation(async (callback: any) => {
        await callback(mockContext);
      });

      const result = await deleteElementByName("TextBox 1");

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(mockShape1.delete).toHaveBeenCalledTimes(1);
      expect(mockShape2.delete).not.toHaveBeenCalled();
    });

    it("应该支持删除所有同名元素", async () => {
      const mockShape1 = {
        name: "TextBox",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShape2 = {
        name: "TextBox",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShape3 = {
        name: "Other",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShapes = {
        items: [mockShape1, mockShape2, mockShape3],
        load: vi.fn(),
      };

      const mockSlide = {
        shapes: mockShapes,
      };

      const mockSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockSelectedSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn().mockReturnValue(mockSelectedSlides),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      mockPowerPoint.run.mockImplementation(async (callback: any) => {
        await callback(mockContext);
      });

      const result = await deleteElementByName("TextBox", undefined, true);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(2);
      expect(mockShape1.delete).toHaveBeenCalledTimes(1);
      expect(mockShape2.delete).toHaveBeenCalledTimes(1);
      expect(mockShape3.delete).not.toHaveBeenCalled();
    });
  });

  describe("deleteElementByIndex", () => {
    it("应该成功通过索引删除元素", async () => {
      const mockShape1 = {
        delete: vi.fn(),
      };

      const mockShape2 = {
        delete: vi.fn(),
      };

      const mockShapes = {
        items: [mockShape1, mockShape2],
        load: vi.fn(),
      };

      const mockSlide = {
        shapes: mockShapes,
      };

      const mockSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockSelectedSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn().mockReturnValue(mockSelectedSlides),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      mockPowerPoint.run.mockImplementation(async (callback: any) => {
        await callback(mockContext);
      });

      const result = await deleteElementByIndex(1);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(mockShape1.delete).not.toHaveBeenCalled();
      expect(mockShape2.delete).toHaveBeenCalledTimes(1);
    });

    it("应该在索引超出范围时抛出错误", async () => {
      const mockShapes = {
        items: [],
        load: vi.fn(),
      };

      const mockSlide = {
        shapes: mockShapes,
      };

      const mockSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockSelectedSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn().mockReturnValue(mockSelectedSlides),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      mockPowerPoint.run.mockImplementation(async (callback: any) => {
        await callback(mockContext);
      });

      const result = await deleteElementByIndex(5);

      expect(result.success).toBe(false);
      expect(result.message).toContain("超出范围");
    });
  });

  describe("deleteElement", () => {
    it("应该优先使用ID删除", async () => {
      const mockShape = {
        id: "shape-123",
        name: "TextBox",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShapes = {
        items: [mockShape],
        load: vi.fn(),
      };

      const mockSlide = {
        shapes: mockShapes,
      };

      const mockSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockSelectedSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn().mockReturnValue(mockSelectedSlides),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      mockPowerPoint.run.mockImplementation(async (callback: any) => {
        await callback(mockContext);
      });

      const result = await deleteElement({
        elementId: "shape-123",
        elementName: "TextBox",
        elementIndex: 0,
      });

      expect(result.success).toBe(true);
      expect(mockShape.delete).toHaveBeenCalledTimes(1);
    });

    it("应该在没有提供任何选择器时返回错误", async () => {
      const result = await deleteElement({});

      expect(result.success).toBe(false);
      expect(result.message).toContain("必须提供");
    });
  });

  describe("deleteElementsByIds", () => {
    it("应该批量删除多个元素", async () => {
      const mockShape1 = {
        id: "shape-123",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShape2 = {
        id: "shape-456",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShape3 = {
        id: "shape-789",
        delete: vi.fn(),
        load: vi.fn(),
      };

      const mockShapes = {
        items: [mockShape1, mockShape2, mockShape3],
        load: vi.fn(),
      };

      const mockSlide = {
        shapes: mockShapes,
      };

      const mockSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockSelectedSlides = {
        items: [mockSlide],
        load: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn().mockReturnValue(mockSelectedSlides),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      mockPowerPoint.run.mockImplementation(async (callback: any) => {
        await callback(mockContext);
      });

      const result = await deleteElementsByIds(["shape-123", "shape-789"]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(2);
      expect(mockShape1.delete).toHaveBeenCalledTimes(1);
      expect(mockShape2.delete).not.toHaveBeenCalled();
      expect(mockShape3.delete).toHaveBeenCalledTimes(1);
    });
  });
});
