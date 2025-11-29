/**
 * 文件名: slideScreenshot.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 描述: slideScreenshot 工具的单元测试 | Unit tests for slideScreenshot tool
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import {
  getSlideScreenshot,
  getCurrentSlideScreenshot,
  getSlideScreenshotByPageNumber,
  getAllSlidesScreenshots,
} from "../../../src/ppt-tools";

// Mock PowerPoint context
const createMockPowerPointContext = () => {
  const mockSlide = {
    id: "slide-123",
    index: 0,
    load: vi.fn(),
    getImageAsBase64: vi.fn().mockReturnValue({
      value: "mockBase64ImageData",
    }),
  };

  const mockSlides = {
    items: [mockSlide],
    load: vi.fn(),
    getItemAt: vi.fn().mockReturnValue(mockSlide),
  };

  const mockSelectedSlides = {
    getItemAt: vi.fn().mockReturnValue(mockSlide),
  };

  return {
    presentation: {
      slides: mockSlides,
      getSelectedSlides: vi.fn().mockReturnValue(mockSelectedSlides),
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
};

describe("slideScreenshot 工具测试 | slideScreenshot Tool Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();

    // Mock PowerPoint global object
    if (!global.PowerPoint) {
      global.PowerPoint = {} as typeof PowerPoint;
    }

    // Mock PowerPoint.run
    const mockContext = createMockPowerPointContext();
    global.PowerPoint.run = vi.fn((callback) => {
      return Promise.resolve(callback(mockContext as any));
    }) as any;
  });

  describe("getSlideScreenshot", () => {
    it("应该能够获取当前选中幻灯片的截图 | should get screenshot of current selected slide", async () => {
      const result = await getSlideScreenshot({});

      expect(global.PowerPoint.run).toHaveBeenCalledTimes(1);
      expect(result).toEqual({
        imageBase64: "mockBase64ImageData",
        slideIndex: 0,
        slideId: "slide-123",
        width: undefined,
        height: undefined,
      });
    });

    it("应该能够获取指定索引的幻灯片截图 | should get screenshot of slide by index", async () => {
      const result = await getSlideScreenshot({ slideIndex: 0 });

      const mockContext = createMockPowerPointContext();
      expect(result).toEqual({
        imageBase64: "mockBase64ImageData",
        slideIndex: 0,
        slideId: "slide-123",
        width: undefined,
        height: undefined,
      });
    });

    it("应该能够获取指定尺寸的截图 | should get screenshot with specified dimensions", async () => {
      const mockContext = createMockPowerPointContext();
      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const result = await getSlideScreenshot({
        width: 800,
        height: 600,
      });

      expect(result).toEqual({
        imageBase64: "mockBase64ImageData",
        slideIndex: 0,
        slideId: "slide-123",
        width: 800,
        height: 600,
      });

      // 验证 getImageAsBase64 被调用时传入了正确的选项
      const slide = mockContext.presentation.getSelectedSlides().getItemAt(0);
      expect(slide.getImageAsBase64).toHaveBeenCalledWith({
        width: 800,
        height: 600,
      });
    });

    it("应该能够获取只指定宽度的截图 | should get screenshot with only width specified", async () => {
      const mockContext = createMockPowerPointContext();
      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const result = await getSlideScreenshot({
        width: 1024,
      });

      expect(result.width).toBe(1024);
      expect(result.height).toBeUndefined();

      const slide = mockContext.presentation.getSelectedSlides().getItemAt(0);
      expect(slide.getImageAsBase64).toHaveBeenCalledWith({
        width: 1024,
      });
    });

    it("应该能够获取只指定高度的截图 | should get screenshot with only height specified", async () => {
      const mockContext = createMockPowerPointContext();
      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const result = await getSlideScreenshot({
        height: 768,
      });

      expect(result.width).toBeUndefined();
      expect(result.height).toBe(768);

      const slide = mockContext.presentation.getSelectedSlides().getItemAt(0);
      expect(slide.getImageAsBase64).toHaveBeenCalledWith({
        height: 768,
      });
    });

    it("应该能够获取指定索引和尺寸的截图 | should get screenshot with index and dimensions", async () => {
      const mockContext = createMockPowerPointContext();
      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const result = await getSlideScreenshot({
        slideIndex: 2,
        width: 1920,
        height: 1080,
      });

      expect(result).toEqual({
        imageBase64: "mockBase64ImageData",
        slideIndex: 0,
        slideId: "slide-123",
        width: 1920,
        height: 1080,
      });

      // 验证使用了指定索引
      expect(mockContext.presentation.slides.getItemAt).toHaveBeenCalledWith(2);
    });

    it("应该正确加载幻灯片的 ID 和索引 | should correctly load slide ID and index", async () => {
      const mockContext = createMockPowerPointContext();
      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      await getSlideScreenshot({});

      const slide = mockContext.presentation.getSelectedSlides().getItemAt(0);
      expect(slide.load).toHaveBeenCalledWith("id,index");
    });

    it("应该在获取截图失败时抛出错误 | should throw error when screenshot fails", async () => {
      const mockError = new Error("获取截图失败");
      global.PowerPoint.run = vi.fn().mockRejectedValue(mockError);

      await expect(getSlideScreenshot({})).rejects.toThrow("获取截图失败");
    });
  });

  describe("getCurrentSlideScreenshot", () => {
    it("应该能够获取当前幻灯片的截图（无参数）| should get current slide screenshot without parameters", async () => {
      const result = await getCurrentSlideScreenshot();

      expect(global.PowerPoint.run).toHaveBeenCalledTimes(1);
      expect(result).toEqual({
        imageBase64: "mockBase64ImageData",
        slideIndex: 0,
        slideId: "slide-123",
        width: undefined,
        height: undefined,
      });
    });

    it("应该能够获取当前幻灯片的截图（带尺寸）| should get current slide screenshot with dimensions", async () => {
      const result = await getCurrentSlideScreenshot(640, 480);

      expect(result).toEqual({
        imageBase64: "mockBase64ImageData",
        slideIndex: 0,
        slideId: "slide-123",
        width: 640,
        height: 480,
      });
    });

    it("应该正确调用 getSlideScreenshot | should correctly call getSlideScreenshot", async () => {
      const width = 800;
      const height = 600;

      const result = await getCurrentSlideScreenshot(width, height);

      expect(result.width).toBe(width);
      expect(result.height).toBe(height);
    });
  });

  describe("getSlideScreenshotByPageNumber", () => {
    it("应该能够根据页码获取截图（页码从 1 开始）| should get screenshot by page number (starting from 1)", async () => {
      const result = await getSlideScreenshotByPageNumber(1);

      expect(result).toEqual({
        imageBase64: "mockBase64ImageData",
        slideIndex: 0,
        slideId: "slide-123",
        width: undefined,
        height: undefined,
      });
    });

    it("应该能够根据页码获取截图（带尺寸）| should get screenshot by page number with dimensions", async () => {
      const result = await getSlideScreenshotByPageNumber(3, 1024, 768);

      expect(result).toEqual({
        imageBase64: "mockBase64ImageData",
        slideIndex: 0,
        slideId: "slide-123",
        width: 1024,
        height: 768,
      });
    });

    it("应该正确将页码转换为索引 | should correctly convert page number to index", async () => {
      const mockContext = createMockPowerPointContext();
      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      await getSlideScreenshotByPageNumber(5);

      // 页码 5 应该转换为索引 4
      expect(mockContext.presentation.slides.getItemAt).toHaveBeenCalledWith(4);
    });

    it("应该在页码小于 1 时抛出错误 | should throw error when page number is less than 1", async () => {
      await expect(getSlideScreenshotByPageNumber(0)).rejects.toThrow("页码必须从 1 开始");
      await expect(getSlideScreenshotByPageNumber(-1)).rejects.toThrow("页码必须从 1 开始");
    });
  });

  describe("getAllSlidesScreenshots", () => {
    it("应该能够获取所有幻灯片的截图 | should get screenshots of all slides", async () => {
      const mockSlide1 = {
        id: "slide-1",
        index: 0,
        load: vi.fn(),
        getImageAsBase64: vi.fn().mockReturnValue({ value: "base64-slide-1" }),
      };

      const mockSlide2 = {
        id: "slide-2",
        index: 1,
        load: vi.fn(),
        getImageAsBase64: vi.fn().mockReturnValue({ value: "base64-slide-2" }),
      };

      const mockSlide3 = {
        id: "slide-3",
        index: 2,
        load: vi.fn(),
        getImageAsBase64: vi.fn().mockReturnValue({ value: "base64-slide-3" }),
      };

      const mockSlides = {
        items: [mockSlide1, mockSlide2, mockSlide3],
        load: vi.fn(),
        getItemAt: vi.fn((index: number) => {
          return [mockSlide1, mockSlide2, mockSlide3][index];
        }),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn(),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const results = await getAllSlidesScreenshots();

      expect(results).toHaveLength(3);
      expect(results[0].imageBase64).toBe("base64-slide-1");
      expect(results[1].imageBase64).toBe("base64-slide-2");
      expect(results[2].imageBase64).toBe("base64-slide-3");
    });

    it("应该能够获取所有幻灯片的截图（带尺寸）| should get all slides screenshots with dimensions", async () => {
      const mockSlide1 = {
        id: "slide-1",
        index: 0,
        load: vi.fn(),
        getImageAsBase64: vi.fn().mockReturnValue({ value: "base64-1" }),
      };

      const mockSlide2 = {
        id: "slide-2",
        index: 1,
        load: vi.fn(),
        getImageAsBase64: vi.fn().mockReturnValue({ value: "base64-2" }),
      };

      const mockSlides = {
        items: [mockSlide1, mockSlide2],
        load: vi.fn(),
        getItemAt: vi.fn((index: number) => [mockSlide1, mockSlide2][index]),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn(),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const results = await getAllSlidesScreenshots(800, 600);

      expect(results).toHaveLength(2);
      expect(results[0].width).toBe(800);
      expect(results[0].height).toBe(600);
      expect(results[1].width).toBe(800);
      expect(results[1].height).toBe(600);
    });

    it("应该处理空演示文稿（无幻灯片）| should handle empty presentation (no slides)", async () => {
      const mockSlides = {
        items: [],
        load: vi.fn(),
        getItemAt: vi.fn(),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const results = await getAllSlidesScreenshots();

      expect(results).toHaveLength(0);
    });

    it("应该正确加载幻灯片列表 | should correctly load slides list", async () => {
      const mockSlide = {
        id: "slide-1",
        index: 0,
        load: vi.fn(),
        getImageAsBase64: vi.fn().mockReturnValue({ value: "base64-data" }),
      };

      const mockSlides = {
        items: [mockSlide],
        load: vi.fn(),
        getItemAt: vi.fn().mockReturnValue(mockSlide),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      await getAllSlidesScreenshots();

      expect(mockSlides.load).toHaveBeenCalledWith("items");
    });

    it("应该在获取失败时抛出错误 | should throw error when fetching fails", async () => {
      const mockError = new Error("获取幻灯片列表失败");
      global.PowerPoint.run = vi.fn().mockRejectedValue(mockError);

      await expect(getAllSlidesScreenshots()).rejects.toThrow("获取幻灯片列表失败");
    });
  });

  describe("边界情况测试 | Edge Cases", () => {
    it("应该处理零尺寸 | should handle zero dimensions", async () => {
      const result = await getSlideScreenshot({
        width: 0,
        height: 0,
      });

      expect(result.width).toBe(0);
      expect(result.height).toBe(0);
    });

    it("应该处理非常大的尺寸值 | should handle very large dimension values", async () => {
      const result = await getSlideScreenshot({
        width: 10000,
        height: 10000,
      });

      expect(result.width).toBe(10000);
      expect(result.height).toBe(10000);
    });

    it("应该处理负数索引（虽然不推荐）| should handle negative index (not recommended)", async () => {
      const mockContext = createMockPowerPointContext();
      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      await getSlideScreenshot({ slideIndex: -1 });

      expect(mockContext.presentation.slides.getItemAt).toHaveBeenCalledWith(-1);
    });

    it("应该处理空选项对象 | should handle empty options object", async () => {
      const result = await getSlideScreenshot({});

      expect(result).toBeDefined();
      expect(result.imageBase64).toBe("mockBase64ImageData");
    });

    it("应该返回正确的 Base64 数据格式 | should return correct Base64 data format", async () => {
      const mockBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
      
      const mockSlide = {
        id: "slide-test",
        index: 0,
        load: vi.fn(),
        getImageAsBase64: vi.fn().mockReturnValue({ value: mockBase64 }),
      };

      const mockContext = {
        presentation: {
          slides: { getItemAt: vi.fn() },
          getSelectedSlides: vi.fn().mockReturnValue({
            getItemAt: vi.fn().mockReturnValue(mockSlide),
          }),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const result = await getSlideScreenshot({});

      // 返回的应该是纯 Base64，不包含 data URL 前缀
      expect(result.imageBase64).toBe(mockBase64);
      expect(result.imageBase64).not.toContain("data:image");
    });

    it("应该处理包含特殊字符的幻灯片 ID | should handle slide ID with special characters", async () => {
      const specialId = "slide-{123-abc-456}";
      
      const mockSlide = {
        id: specialId,
        index: 0,
        load: vi.fn(),
        getImageAsBase64: vi.fn().mockReturnValue({ value: "base64data" }),
      };

      const mockContext = {
        presentation: {
          slides: { getItemAt: vi.fn() },
          getSelectedSlides: vi.fn().mockReturnValue({
            getItemAt: vi.fn().mockReturnValue(mockSlide),
          }),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const result = await getSlideScreenshot({});

      expect(result.slideId).toBe(specialId);
    });
  });

  describe("性能测试 | Performance Tests", () => {
    it("应该能够快速获取单张幻灯片截图 | should quickly get single slide screenshot", async () => {
      const startTime = Date.now();
      await getCurrentSlideScreenshot();
      const endTime = Date.now();

      // 模拟环境下应该很快完成（< 100ms）
      expect(endTime - startTime).toBeLessThan(100);
    });

    it("应该能够批量获取多张幻灯片截图 | should get multiple slides screenshots in batch", async () => {
      const slides = Array.from({ length: 10 }, (_, i) => ({
        id: `slide-${i}`,
        index: i,
        load: vi.fn(),
        getImageAsBase64: vi.fn().mockReturnValue({ value: `base64-${i}` }),
      }));

      const mockSlides = {
        items: slides,
        load: vi.fn(),
        getItemAt: vi.fn((index: number) => slides[index]),
      };

      const mockContext = {
        presentation: {
          slides: mockSlides,
          getSelectedSlides: vi.fn(),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.PowerPoint.run = vi.fn((callback) => {
        return Promise.resolve(callback(mockContext as any));
      }) as any;

      const results = await getAllSlidesScreenshots();

      expect(results).toHaveLength(10);
      results.forEach((result, index) => {
        expect(result.imageBase64).toBe(`base64-${index}`);
      });
    });
  });
});
