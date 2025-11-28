/**
 * 文件名: elementsList.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 描述: elementsList 工具的单元测试
 */

import {beforeEach, describe, expect, it} from 'vitest';
import {OfficeMockObject} from 'office-addin-mock';
import {getCurrentSlideElements, getSlideElements, getSlideElementsByPageNumber} from '../../../src/ppt-tools';

// 定义 PowerPoint context 类型
type MockContext = {
  presentation: {
    slides: {
      items: ReturnType<typeof createMockSlide>[];
    };
    getSelectedSlides?: () => {
      items: ReturnType<typeof createMockSlide>[];
      load: () => unknown;
    };
  };
};

// 创建 mock 形状数据
const createMockShape = (id: string, text?: string) => ({
  id,
  type: 'TextBox',
  left: 100,
  top: 200,
  width: 300,
  height: 150,
  name: 'Test Shape',
  load: function() {
    return this;
  },
  textFrame: text ? {
    textRange: {
      text,
      load: function() {
        return this;
      },
    },
    load: function() {
      return this;
    },
  } : undefined,
});

// 创建 mock 幻灯片数据
const createMockSlide = (shapeCount: number = 1) => ({
  shapes: {
    items: Array.from({ length: shapeCount }, (_, i) => 
      createMockShape(`shape-${i + 1}`, `Text ${i + 1}`)
    ),
    load: function() {
      return this;
    },
  },
});

describe('elementsList 工具测试', () => {
  beforeEach(() => {
    // 重置 globalThis.PowerPoint
    // 使用类型断言因为我们要删除的是 mock 对象，而非真实的 PowerPoint 命名空间
    delete (globalThis as { PowerPoint?: unknown }).PowerPoint;
  });

  describe('getSlideElements', () => {
    it('应该能够获取指定页码的元素', async () => {
      // 创建 mock 数据
      const mockData = {
        context: {
          presentation: {
            slides: {
              items: [
                createMockSlide(2),
                createMockSlide(3),
                createMockSlide(1),
              ],
            },
          },
        },
        run: async function (callback: (context: MockContext) => Promise<void>) {
          await callback(this.context);
        },
      };

      // 使用类型断言将 mock 对象赋值给 globalThis.PowerPoint
      (globalThis as { PowerPoint: unknown }).PowerPoint = new OfficeMockObject(mockData);

      const elements = await getSlideElements({ slideNumber: 2 });

      expect(elements).toBeDefined();
      expect(Array.isArray(elements)).toBe(true);
      expect(elements.length).toBe(3);
    });

    it('应该在页码不存在时返回空数组', async () => {
      const mockData = {
        context: {
          presentation: {
            slides: {
              items: [
                createMockSlide(1),
                createMockSlide(1),
              ],
            },
          },
        },
        run: async function (callback: (context: MockContext) => Promise<void>) {
          await callback(this.context);
        },
      };

      // 使用类型断言将 mock 对象赋值给 globalThis.PowerPoint
      (globalThis as { PowerPoint: unknown }).PowerPoint = new OfficeMockObject(mockData);

      const elements = await getSlideElements({ slideNumber: 5 });

      expect(elements).toEqual([]);
    });

    it('应该在页码为0或负数时返回空数组', async () => {
      const mockData = {
        context: {
          presentation: {
            slides: {
              items: [createMockSlide(1)],
            },
          },
        },
        run: async function (callback: (context: MockContext) => Promise<void>) {
          await callback(this.context);
        },
      };

      // 使用类型断言将 mock 对象赋值给 globalThis.PowerPoint
      (globalThis as { PowerPoint: unknown }).PowerPoint = new OfficeMockObject(mockData);

      const elements1 = await getSlideElements({ slideNumber: 0 });
      const elements2 = await getSlideElements({ slideNumber: -1 });

      expect(elements1).toEqual([]);
      expect(elements2).toEqual([]);
    });

    it('应该在不指定页码时使用当前选中的幻灯片', async () => {
      const mockData = {
        context: {
          presentation: {
            slides: {
              items: [createMockSlide(2)],
            },
            getSelectedSlides: function () {
              return {
                items: [this.slides.items[0]],
                load: function() {
                  return this;
                },
              };
            },
          },
        },
        run: async function (callback: (context: MockContext) => Promise<void>) {
          await callback(this.context);
        },
      };

      // 使用类型断言将 mock 对象赋值给 globalThis.PowerPoint
      (globalThis as { PowerPoint: unknown }).PowerPoint = new OfficeMockObject(mockData);

      const elements = await getSlideElements({});

      expect(elements).toBeDefined();
      expect(Array.isArray(elements)).toBe(true);
      expect(elements.length).toBe(2);
    });
  });

  describe('getCurrentSlideElements', () => {
    it('应该调用 getSlideElements 获取当前页元素', async () => {
      const mockData = {
        context: {
          presentation: {
            slides: {
              items: [createMockSlide(1)],
            },
            getSelectedSlides: function () {
              return {
                items: [this.slides.items[0]],
                load: function() {
                  return this;
                },
              };
            },
          },
        },
        run: async function (callback: (context: MockContext) => Promise<void>) {
          await callback(this.context);
        },
      };

      // 使用类型断言将 mock 对象赋值给 globalThis.PowerPoint
      (globalThis as { PowerPoint: unknown }).PowerPoint = new OfficeMockObject(mockData);

      const elements = await getCurrentSlideElements();

      expect(elements).toBeDefined();
      expect(Array.isArray(elements)).toBe(true);
      expect(elements.length).toBe(1);
    });
  });

  describe('getSlideElementsByPageNumber', () => {
    it('应该能够通过页码获取元素', async () => {
      const mockData = {
        context: {
          presentation: {
            slides: {
              items: [
                createMockSlide(2),
                createMockSlide(3),
              ],
            },
          },
        },
        run: async function (callback: (context: MockContext) => Promise<void>) {
          await callback(this.context);
        },
      };

      // 使用类型断言将 mock 对象赋值给 globalThis.PowerPoint
      (globalThis as { PowerPoint: unknown }).PowerPoint = new OfficeMockObject(mockData);

      const elements = await getSlideElementsByPageNumber(1);

      expect(elements).toBeDefined();
      expect(Array.isArray(elements)).toBe(true);
      expect(elements.length).toBe(2);
    });

    it('应该支持 includeText 参数', async () => {
      const mockData = {
        context: {
          presentation: {
            slides: {
              items: [createMockSlide(1)],
            },
          },
        },
        run: async function (callback: (context: MockContext) => Promise<void>) {
          await callback(this.context);
        },
      };

      // 使用类型断言将 mock 对象赋值给 globalThis.PowerPoint
      (globalThis as { PowerPoint: unknown }).PowerPoint = new OfficeMockObject(mockData);

      const elements = await getSlideElementsByPageNumber(1, false);

      expect(elements).toBeDefined();
      expect(Array.isArray(elements)).toBe(true);
    });
  });
});
