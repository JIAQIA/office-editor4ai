/**
 * 文件名: slideMove.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: slideMove 工具的单元测试 | slideMove tool unit tests
 */

import { describe, it, expect, beforeEach } from "vitest";
import { moveSlide, moveCurrentSlide, swapSlides, getAllSlidesInfo } from "../../../src/ppt-tools";

type MockSlide = {
  id: string;
  moveTo: (index: number) => void;
  shapes: {
    items: Array<{
      type: string;
      name?: string;
      textFrame?: {
        textRange: {
          text: string;
          load: () => void;
        };
        load: () => void;
      };
      load: () => void;
    }>;
    load: () => void;
  };
  load: () => void;
};

type MockData = {
  context: {
    presentation: {
      slides: {
        items: MockSlide[];
        load: () => void;
      };
      getSelectedSlides: () => {
        items: MockSlide[];
        load: () => void;
      };
    };
    sync: () => Promise<void>;
  };
  run: (callback: (context: MockData["context"]) => Promise<void>) => Promise<void>;
  _getSlide: (id: string) => MockSlide | undefined;
  _moveSlide: (fromIndex: number, toIndex: number) => void;
};

// 创建 mock 幻灯片对象
const createMockSlide = (id: string, title?: string): MockSlide => {
  const slide: MockSlide = {
    id,
    moveTo: function (_index: number) {
      // Mock implementation - will be handled by mockData._moveSlide
    },
    shapes: {
      items: title
        ? [
            {
              type: "TextBox",
              name: "Title",
              textFrame: {
                textRange: {
                  text: title,
                  load: function () {},
                },
                load: function () {},
              },
              load: function () {},
            },
          ]
        : [],
      load: function () {},
    },
    load: function () {},
  };
  return slide;
};

// 创建 mock 数据
const createMockData = (): MockData => {
  const slides: MockSlide[] = [
    createMockSlide("slide-1", "Slide 1"),
    createMockSlide("slide-2", "Slide 2"),
    createMockSlide("slide-3", "Slide 3"),
    createMockSlide("slide-4", "Slide 4"),
    createMockSlide("slide-5", "Slide 5"),
  ];

  let selectedSlideId = "slide-1";

  const mockData: MockData = {
    context: {
      presentation: {
        slides: {
          items: slides,
          load: function () {},
        },
        getSelectedSlides: function () {
          const selectedSlide = slides.find((s) => s.id === selectedSlideId);
          return {
            items: selectedSlide ? [selectedSlide] : [],
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
    _getSlide: function (id: string) {
      return slides.find((s) => s.id === id);
    },
    _moveSlide: function (fromIndex: number, toIndex: number) {
      if (fromIndex < 0 || fromIndex >= slides.length || toIndex < 0 || toIndex >= slides.length) {
        throw new Error("Index out of range");
      }
      const [movedSlide] = slides.splice(fromIndex, 1);
      slides.splice(toIndex, 0, movedSlide);
    },
  };

  // Override moveTo to use _moveSlide
  slides.forEach((slide, index) => {
    slide.moveTo = function (targetIndex: number) {
      mockData._moveSlide(index, targetIndex);
    };
  });

  return mockData;
};

describe("slideMove", () => {
  let mockData: MockData;

  beforeEach(() => {
    mockData = createMockData();
    // @ts-expect-error - Mock PowerPoint global
    global.PowerPoint = mockData;
  });

  describe("moveSlide", () => {
    it("应该成功移动幻灯片", async () => {
      const result = await moveSlide({ fromIndex: 1, toIndex: 3 });

      expect(result.success).toBe(true);
      expect(result.message).toContain("成功将幻灯片从位置 1 移动到位置 3");
      expect(result.fromIndex).toBe(1);
      expect(result.toIndex).toBe(3);
      expect(result.totalSlides).toBe(5);
    });

    it("应该拒绝无效的源位置索引", async () => {
      const result = await moveSlide({ fromIndex: 0, toIndex: 3 });

      expect(result.success).toBe(false);
      expect(result.message).toContain("源位置索引必须是大于0的整数");
    });

    it("应该拒绝无效的目标位置索引", async () => {
      const result = await moveSlide({ fromIndex: 1, toIndex: 0 });

      expect(result.success).toBe(false);
      expect(result.message).toContain("目标位置索引必须是大于0的整数");
    });

    it("应该拒绝源位置和目标位置相同", async () => {
      const result = await moveSlide({ fromIndex: 2, toIndex: 2 });

      expect(result.success).toBe(false);
      expect(result.message).toContain("源位置和目标位置相同，无需移动");
    });

    it("应该拒绝超出范围的索引", async () => {
      const result = await moveSlide({ fromIndex: 10, toIndex: 3 });

      expect(result.success).toBe(false);
      expect(result.message).toContain("超出范围");
    });

    it("应该拒绝非整数索引", async () => {
      const result = await moveSlide({ fromIndex: 1.5, toIndex: 3 });

      expect(result.success).toBe(false);
      expect(result.message).toContain("源位置索引必须是大于0的整数");
    });
  });

  describe("moveCurrentSlide", () => {
    it("应该成功移动当前选中的幻灯片", async () => {
      const result = await moveCurrentSlide(4);

      expect(result.success).toBe(true);
      expect(result.message).toContain("成功将当前幻灯片");
      expect(result.toIndex).toBe(4);
    });

    it("应该拒绝无效的目标位置索引", async () => {
      const result = await moveCurrentSlide(-1);

      expect(result.success).toBe(false);
      expect(result.message).toContain("目标位置索引必须是大于0的整数");
    });

    it("应该拒绝非整数目标位置", async () => {
      const result = await moveCurrentSlide(2.5);

      expect(result.success).toBe(false);
      expect(result.message).toContain("目标位置索引必须是大于0的整数");
    });
  });

  describe("swapSlides", () => {
    it("应该成功交换两张幻灯片", async () => {
      const result = await swapSlides(1, 3);

      expect(result.success).toBe(true);
      expect(result.message).toContain("成功交换位置 1 和位置 3 的幻灯片");
      expect(result.fromIndex).toBe(1);
      expect(result.toIndex).toBe(3);
    });

    it("应该拒绝无效的第一张幻灯片索引", async () => {
      const result = await swapSlides(0, 3);

      expect(result.success).toBe(false);
      expect(result.message).toContain("第一张幻灯片索引必须是大于0的整数");
    });

    it("应该拒绝无效的第二张幻灯片索引", async () => {
      const result = await swapSlides(1, 0);

      expect(result.success).toBe(false);
      expect(result.message).toContain("第二张幻灯片索引必须是大于0的整数");
    });

    it("应该拒绝两张幻灯片索引相同", async () => {
      const result = await swapSlides(2, 2);

      expect(result.success).toBe(false);
      expect(result.message).toContain("两张幻灯片索引相同，无需交换");
    });

    it("应该处理不同顺序的索引", async () => {
      const result = await swapSlides(4, 2);

      expect(result.success).toBe(true);
      expect(result.message).toContain("成功交换");
    });
  });

  describe("getAllSlidesInfo", () => {
    it("应该返回所有幻灯片的信息", async () => {
      const slidesInfo = await getAllSlidesInfo();

      expect(slidesInfo).toHaveLength(5);
      expect(slidesInfo[0]).toMatchObject({
        index: 1,
        id: "slide-1",
        title: "Slide 1",
      });
      expect(slidesInfo[4]).toMatchObject({
        index: 5,
        id: "slide-5",
        title: "Slide 5",
      });
    });

    it("应该处理没有标题的幻灯片", async () => {
      // 创建一个没有标题的幻灯片
      const mockDataNoTitle = createMockData();
      mockDataNoTitle.context.presentation.slides.items[0].shapes.items = [];
      // @ts-expect-error - Mock PowerPoint global
      global.PowerPoint = mockDataNoTitle;

      const slidesInfo = await getAllSlidesInfo();

      expect(slidesInfo[0].title).toBeUndefined();
    });

    it("应该处理空演示文稿", async () => {
      const emptyMockData = createMockData();
      emptyMockData.context.presentation.slides.items = [];
      // @ts-expect-error - Mock PowerPoint global
      global.PowerPoint = emptyMockData;

      const slidesInfo = await getAllSlidesInfo();

      expect(slidesInfo).toHaveLength(0);
    });
  });

  describe("边界情况", () => {
    it("应该处理只有一张幻灯片的情况", async () => {
      const singleSlideMock = createMockData();
      singleSlideMock.context.presentation.slides.items = [
        createMockSlide("slide-1", "Only Slide"),
      ];
      // @ts-expect-error - Mock PowerPoint global
      global.PowerPoint = singleSlideMock;

      const result = await moveSlide({ fromIndex: 1, toIndex: 1 });

      expect(result.success).toBe(false);
      expect(result.message).toContain("源位置和目标位置相同");
    });

    it("应该处理移动到末尾", async () => {
      const result = await moveSlide({ fromIndex: 1, toIndex: 5 });

      expect(result.success).toBe(true);
      expect(result.toIndex).toBe(5);
    });

    it("应该处理移动到开头", async () => {
      const result = await moveSlide({ fromIndex: 5, toIndex: 1 });

      expect(result.success).toBe(true);
      expect(result.toIndex).toBe(1);
    });
  });
});
