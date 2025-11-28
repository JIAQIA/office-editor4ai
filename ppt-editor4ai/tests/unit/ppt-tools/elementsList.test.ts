/**
 * 文件名: elementsList.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 描述: elementsList 工具的单元测试
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { getSlideElements, getCurrentSlideElements, getSlideElementsByPageNumber } from '../../../src/ppt-tools';

// Mock PowerPoint API
const mockContext = {
  presentation: {
    slides: {
      items: [],
      load: vi.fn(),
      getItemAt: vi.fn(),
    },
    getSelectedSlides: vi.fn(),
  },
  sync: vi.fn(),
};

const mockShape = {
  id: 'shape-1',
  type: 'TextBox',
  left: 100,
  top: 200,
  width: 300,
  height: 150,
  name: 'Test Shape',
  load: vi.fn(),
  textFrame: {
    textRange: {
      text: 'Test Text',
      load: vi.fn(),
    },
    load: vi.fn(),
  },
};

global.PowerPoint = {
  run: vi.fn((callback) => callback(mockContext)),
} as any;

describe('elementsList 工具测试', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('getSlideElements', () => {
    it('应该能够获取指定页码的元素', async () => {
      // 模拟3页幻灯片
      mockContext.presentation.slides.items = [
        { shapes: { items: [mockShape], load: vi.fn() } },
        { shapes: { items: [mockShape], load: vi.fn() } },
        { shapes: { items: [mockShape], load: vi.fn() } },
      ];

      const elements = await getSlideElements({ slideNumber: 2 });

      expect(elements).toBeDefined();
      expect(Array.isArray(elements)).toBe(true);
    });

    it('应该在页码不存在时返回空数组', async () => {
      // 模拟只有2页幻灯片
      mockContext.presentation.slides.items = [
        { shapes: { items: [mockShape], load: vi.fn() } },
        { shapes: { items: [mockShape], load: vi.fn() } },
      ];

      const elements = await getSlideElements({ slideNumber: 5 });

      expect(elements).toEqual([]);
    });

    it('应该在页码为0或负数时返回空数组', async () => {
      mockContext.presentation.slides.items = [
        { shapes: { items: [mockShape], load: vi.fn() } },
      ];

      const elements1 = await getSlideElements({ slideNumber: 0 });
      const elements2 = await getSlideElements({ slideNumber: -1 });

      expect(elements1).toEqual([]);
      expect(elements2).toEqual([]);
    });

    it('应该在不指定页码时使用当前选中的幻灯片', async () => {
      const mockSelectedSlide = {
        shapes: { items: [mockShape], load: vi.fn() },
      };

      mockContext.presentation.getSelectedSlides.mockReturnValue({
        items: [mockSelectedSlide],
        load: vi.fn(),
      });

      const elements = await getSlideElements({});

      expect(mockContext.presentation.getSelectedSlides).toHaveBeenCalled();
      expect(elements).toBeDefined();
    });
  });

  describe('getCurrentSlideElements', () => {
    it('应该调用 getSlideElements 获取当前页元素', async () => {
      const mockSelectedSlide = {
        shapes: { items: [], load: vi.fn() },
      };

      mockContext.presentation.getSelectedSlides.mockReturnValue({
        items: [mockSelectedSlide],
        load: vi.fn(),
      });

      const elements = await getCurrentSlideElements();

      expect(elements).toBeDefined();
      expect(Array.isArray(elements)).toBe(true);
    });
  });

  describe('getSlideElementsByPageNumber', () => {
    it('应该能够通过页码获取元素', async () => {
      mockContext.presentation.slides.items = [
        { shapes: { items: [mockShape], load: vi.fn() } },
        { shapes: { items: [mockShape], load: vi.fn() } },
      ];

      const elements = await getSlideElementsByPageNumber(1);

      expect(elements).toBeDefined();
      expect(Array.isArray(elements)).toBe(true);
    });

    it('应该支持 includeText 参数', async () => {
      mockContext.presentation.slides.items = [
        { shapes: { items: [mockShape], load: vi.fn() } },
      ];

      const elements = await getSlideElementsByPageNumber(1, false);

      expect(elements).toBeDefined();
    });
  });
});
