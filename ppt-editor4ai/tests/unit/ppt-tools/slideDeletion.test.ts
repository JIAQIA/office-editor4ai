/**
 * 文件名: slideDeletion.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: slideDeletion 工具的单元测试
 */

import { beforeEach, describe, expect, it, vi } from "vitest";
import { deleteCurrentSlide, deleteSlides, deleteSlidesByNumbers } from "../../../src/ppt-tools";

// 定义 Mock 幻灯片类型
type MockSlide = {
  id: string;
  _deleted: boolean;
  delete: () => void;
  load: (props?: string) => MockSlide;
};

// 定义 PowerPoint context 类型
type MockContext = {
  presentation: {
    slides: {
      items: MockSlide[];
      load: (props?: string) => void;
    };
    getSelectedSlides?: () => {
      items: MockSlide[];
      load: (props?: string) => void;
    };
  };
  sync: () => Promise<void>;
};

// 创建 mock 幻灯片对象
const createMockSlide = (id: string): MockSlide => {
  const slide: MockSlide = {
    id,
    _deleted: false,
    delete: function () {
      slide._deleted = true;
    },
    load: function (_props?: string) {
      return slide;
    },
  };
  return slide;
};

// 创建 mock PowerPoint 数据
const createMockData = (slideCount: number = 5) => {
  const slides = Array.from({ length: slideCount }, (_, i) => createMockSlide(`slide-${i + 1}`));

  let itemsLoaded = false;

  const slidesObject = {
    items: slides,
    load: function (_props?: string) {
      // Mock load 方法，标记为已加载
      itemsLoaded = true;
    },
  };

  return {
    context: {
      presentation: {
        slides: slidesObject,
        getSelectedSlides: function () {
          return {
            items: [slides[0]],
            load: function (_props?: string) {
              itemsLoaded = true;
            },
          };
        },
      },
      sync: async function () {
        // Mock sync 方法 - 在这里确保 items 可用
        itemsLoaded = true;
      },
    },
    run: async function (callback: (context: MockContext) => Promise<void>) {
      await callback(this.context);
    },
    _getSlides: () => slides,
  };
};

describe("slideDeletion 工具测试 / Slide Deletion Tool Tests", () => {
  beforeEach(() => {
    // 重置 global.PowerPoint
    delete (global as any).PowerPoint;
    // 清除 console 的 mock
    vi.restoreAllMocks();
  });

  describe("deleteSlidesByNumbers", () => {
    it("应该能够删除单个幻灯片 / Should delete a single slide", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([3]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(result.failedCount).toBe(0);
      expect(result.details?.deleted).toEqual([3]);
      expect(result.details?.notFound).toEqual([]);
      expect(result.details?.errors).toEqual([]);

      const slides = (global as any).PowerPoint._getSlides();
      expect(slides[2]._deleted).toBe(true);
    });

    it("应该能够删除多个幻灯片 / Should delete multiple slides", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([1, 3, 5]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(3);
      expect(result.failedCount).toBe(0);
      expect(result.details?.deleted).toEqual([5, 3, 1]); // 按从大到小排序

      const slides = (global as any).PowerPoint._getSlides();
      expect(slides[0]._deleted).toBe(true);
      expect(slides[2]._deleted).toBe(true);
      expect(slides[4]._deleted).toBe(true);
    });

    it("应该按从大到小的顺序删除幻灯片 / Should delete slides in descending order", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([2, 4, 1, 5, 3]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(5);
      // 验证删除顺序是从大到小
      expect(result.details?.deleted).toEqual([5, 4, 3, 2, 1]);
    });

    it("应该去除重复的页码 / Should remove duplicate slide numbers", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([2, 2, 3, 3, 2]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(2); // 只删除 2 个不同的页码
      expect(result.details?.deleted).toEqual([3, 2]);
    });

    it("应该处理不存在的页码 / Should handle non-existent slide numbers", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([3, 10, 15]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(result.failedCount).toBe(2);
      expect(result.details?.deleted).toEqual([3]);
      expect(result.details?.notFound).toEqual([15, 10]); // 按从大到小排序
      expect(result.message).toContain("页码不存在");
    });

    it("应该处理负数页码 / Should handle negative slide numbers", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([-1, 0, 3]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(result.failedCount).toBe(2);
      expect(result.details?.deleted).toEqual([3]);
      expect(result.details?.notFound).toEqual([0, -1]);
    });

    it("应该处理空数组 / Should handle empty array", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([]);

      expect(result.success).toBe(false);
      expect(result.deletedCount).toBe(0);
      expect(result.failedCount).toBe(0);
    });

    it("应该处理全部页码不存在的情况 / Should handle all non-existent slide numbers", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([10, 20, 30]);

      expect(result.success).toBe(false);
      expect(result.deletedCount).toBe(0);
      expect(result.failedCount).toBe(3);
      expect(result.details?.notFound).toEqual([30, 20, 10]);
    });

    it("应该在删除失败时返回错误信息 / Should return error info on deletion failure", async () => {
      (global as any).PowerPoint = createMockData(5);
      // 模拟删除时抛出错误
      (global as any).PowerPoint._getSlides()[2].delete = function () {
        throw new Error("删除失败");
      };

      const result = await deleteSlidesByNumbers([3]);

      expect(result.success).toBe(false);
      expect(result.deletedCount).toBe(0);
      expect(result.failedCount).toBe(1);
      expect(result.details?.errors).toHaveLength(1);
      expect(result.details?.errors[0].slideNumber).toBe(3);
      expect(result.details?.errors[0].error).toBe("删除失败");
    });

    it("应该在 PowerPoint.run 失败时返回错误 / Should return error when PowerPoint.run fails", async () => {
      (global as any).PowerPoint = {
        run: async function () {
          throw new Error("PowerPoint API 错误");
        },
      };

      const result = await deleteSlidesByNumbers([1, 2]);

      expect(result.success).toBe(false);
      expect(result.deletedCount).toBe(0);
      expect(result.failedCount).toBe(2);
      expect(result.message).toContain("PowerPoint API 错误");
    });
  });

  describe("deleteCurrentSlide", () => {
    it("应该能够删除当前选中的幻灯片 / Should delete the currently selected slide", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteCurrentSlide();

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(result.failedCount).toBe(0);
      expect(result.details?.deleted).toEqual([1]); // 默认选中第一页
      expect(result.message).toContain("成功删除第 1 页");

      const slides = (global as any).PowerPoint._getSlides();
      expect(slides[0]._deleted).toBe(true);
    });

    it("应该在没有选中幻灯片时返回错误 / Should return error when no slide is selected", async () => {
      const mockData = createMockData(5);
      mockData.context.presentation.getSelectedSlides = function () {
        return {
          items: [],
          load: function (_props?: string) {
            // Mock load 方法
          },
        };
      };
      (global as any).PowerPoint = mockData;

      const result = await deleteCurrentSlide();

      expect(result.success).toBe(false);
      expect(result.deletedCount).toBe(0);
      expect(result.failedCount).toBe(1);
      expect(result.message).toContain("没有选中的幻灯片");
    });

    it("应该能够删除选中的第三页 / Should delete the selected third slide", async () => {
      const mockData = createMockData(5);
      mockData.context.presentation.getSelectedSlides = function () {
        return {
          items: [this.slides.items[2]], // 选中第三页
          load: function (_props?: string) {
            // Mock load 方法
          },
        };
      };
      (global as any).PowerPoint = mockData;

      const result = await deleteCurrentSlide();

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(result.details?.deleted).toEqual([3]);

      const slides = (global as any).PowerPoint._getSlides();
      expect(slides[2]._deleted).toBe(true);
    });

    it("应该在 PowerPoint.run 失败时返回错误 / Should return error when PowerPoint.run fails", async () => {
      (global as any).PowerPoint = {
        run: async function () {
          throw new Error("API 调用失败");
        },
      };

      const result = await deleteCurrentSlide();

      expect(result.success).toBe(false);
      expect(result.deletedCount).toBe(0);
      expect(result.failedCount).toBe(1);
      expect(result.message).toContain("API 调用失败");
    });
  });

  describe("deleteSlides", () => {
    it("应该在指定页码时调用 deleteSlidesByNumbers / Should call deleteSlidesByNumbers when slide numbers are provided", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlides({ slideNumbers: [2, 4] });

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(2);
      expect(result.details?.deleted).toEqual([4, 2]);
    });

    it("应该在未指定页码时删除当前幻灯片 / Should delete current slide when no slide numbers are provided", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlides({});

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(result.details?.deleted).toEqual([1]);
    });

    it("应该在 deleteCurrentSlide 为 false 且无页码时返回错误 / Should return error when deleteCurrentSlide is false and no slide numbers", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlides({ deleteCurrentSlide: false });

      expect(result.success).toBe(false);
      expect(result.deletedCount).toBe(0);
      expect(result.failedCount).toBe(0);
      expect(result.message).toBe("未指定要删除的幻灯片");
    });

    it("应该优先使用指定的页码而非当前幻灯片 / Should prioritize slide numbers over current slide", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlides({
        slideNumbers: [3, 5],
        deleteCurrentSlide: true,
      });

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(2);
      expect(result.details?.deleted).toEqual([5, 3]);

      // 验证只删除了指定页码，而非当前页
      const slides = (global as any).PowerPoint._getSlides();
      expect(slides[0]._deleted).toBe(false); // 第一页（当前页）未删除
      expect(slides[2]._deleted).toBe(true); // 第三页已删除
      expect(slides[4]._deleted).toBe(true); // 第五页已删除
    });

    it("应该处理空页码数组 / Should handle empty slide numbers array", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlides({ slideNumbers: [] });

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);
      expect(result.details?.deleted).toEqual([1]); // 删除当前页
    });
  });

  describe("边界情况测试 / Edge Cases", () => {
    it("应该能够删除只有一页的演示文稿的幻灯片 / Should delete slide from single-slide presentation", async () => {
      (global as any).PowerPoint = createMockData(1);

      const result = await deleteSlidesByNumbers([1]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);

      const slides = (global as any).PowerPoint._getSlides();
      expect(slides[0]._deleted).toBe(true);
    });

    it("应该能够删除最后一页 / Should delete the last slide", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([5]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);

      const slides = (global as any).PowerPoint._getSlides();
      expect(slides[4]._deleted).toBe(true);
    });

    it("应该能够删除第一页 / Should delete the first slide", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([1]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(1);

      const slides = (global as any).PowerPoint._getSlides();
      expect(slides[0]._deleted).toBe(true);
    });

    it("应该能够删除所有幻灯片 / Should delete all slides", async () => {
      (global as any).PowerPoint = createMockData(3);

      const result = await deleteSlidesByNumbers([1, 2, 3]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(3);

      const slides = (global as any).PowerPoint._getSlides();
      expect(slides.every((slide) => slide._deleted)).toBe(true);
    });

    it("应该处理超大页码 / Should handle very large slide numbers", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([999999]);

      expect(result.success).toBe(false);
      expect(result.deletedCount).toBe(0);
      expect(result.failedCount).toBe(1);
      expect(result.details?.notFound).toEqual([999999]);
    });

    it("应该处理混合的有效和无效页码 / Should handle mixed valid and invalid slide numbers", async () => {
      (global as any).PowerPoint = createMockData(5);

      const result = await deleteSlidesByNumbers([2, 10, 4, -5, 0, 3]);

      expect(result.success).toBe(true);
      expect(result.deletedCount).toBe(3); // 2, 3, 4
      expect(result.failedCount).toBe(3); // 10, -5, 0
      expect(result.details?.deleted).toEqual([4, 3, 2]);
      expect(result.details?.notFound).toEqual([10, 0, -5]);
    });
  });
});
