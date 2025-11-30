/**
 * 文件名: shapeInsertion.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: shapeInsertion 工具的单元测试 | Unit tests for shapeInsertion tool
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { OfficeMockObject } from 'office-addin-mock';
import { insertShapeToSlide, insertShape, COMMON_SHAPES } from '../../../src/ppt-tools';

// Mock PowerPoint.GeometricShapeType 枚举
const mockGeometricShapeType = {
  rectangle: 'rectangle',
  roundRectangle: 'roundRectangle',
  ellipse: 'ellipse',
  diamond: 'diamond',
  triangle: 'triangle',
  rightTriangle: 'rightTriangle',
  parallelogram: 'parallelogram',
  trapezoid: 'trapezoid',
  hexagon: 'hexagon',
  octagon: 'octagon',
  plus: 'plus',
  star5: 'star5',
  star6: 'star6',
  star7: 'star7',
  star8: 'star8',
  star10: 'star10',
  star12: 'star12',
  star16: 'star16',
  star24: 'star24',
  star32: 'star32',
  round1Rectangle: 'round1Rectangle',
  round2SameRectangle: 'round2SameRectangle',
  round2DiagonalRectangle: 'round2DiagonalRectangle',
  snipRoundRectangle: 'snipRoundRectangle',
  snip1Rectangle: 'snip1Rectangle',
  snip2SameRectangle: 'snip2SameRectangle',
  snip2DiagonalRectangle: 'snip2DiagonalRectangle',
  plaque: 'plaque',
  ellipseRibbon: 'ellipseRibbon',
  ellipseRibbon2: 'ellipseRibbon2',
  leftRightRibbon: 'leftRightRibbon',
  verticalScroll: 'verticalScroll',
  horizontalScroll: 'horizontalScroll',
  wave: 'wave',
  doubleWave: 'doubleWave',
  leftRightArrow: 'leftRightArrow',
  upDownArrow: 'upDownArrow',
  leftUpArrow: 'leftUpArrow',
  bentUpArrow: 'bentUpArrow',
  bentArrow: 'bentArrow',
  stripedRightArrow: 'stripedRightArrow',
  notchedRightArrow: 'notchedRightArrow',
  homePlate: 'homePlate',
  chevron: 'chevron',
  rightArrowCallout: 'rightArrowCallout',
  downArrowCallout: 'downArrowCallout',
  leftArrowCallout: 'leftArrowCallout',
  upArrowCallout: 'upArrowCallout',
  leftRightArrowCallout: 'leftRightArrowCallout',
  quadArrowCallout: 'quadArrowCallout',
  circularArrow: 'circularArrow',
  mathPlus: 'mathPlus',
  mathMinus: 'mathMinus',
  mathMultiply: 'mathMultiply',
  mathDivide: 'mathDivide',
  mathEqual: 'mathEqual',
  mathNotEqual: 'mathNotEqual',
  cornerTabs: 'cornerTabs',
  squareTabs: 'squareTabs',
  plaqueTabs: 'plaqueTabs',
  chartX: 'chartX',
  chartStar: 'chartStar',
  chartPlus: 'chartPlus',
};

type MockShape = {
  id: string;
  type: string;
  width: number;
  height: number;
  left: number;
  top: number;
  fill: {
    color: string;
    setSolidColor: (color: string) => void;
  };
  lineFormat: {
    color: string;
    weight: number;
  };
  textFrame: {
    textRange: {
      text: string;
    };
  };
  load: (properties: string) => void;
  _shapeType?: string;
  _options?: unknown;
};

type MockData = {
  context: {
    presentation: {
      getSelectedSlides: () => {
        getItemAt: () => {
          shapes: {
            addGeometricShape: (shapeType: string, options?: unknown) => MockShape;
          };
        };
      };
    };
    sync: () => Promise<void>;
  };
  run: (callback: (context: MockData['context']) => Promise<any>) => Promise<any>;
  _getShape: () => MockShape;
};

// Mock getSlideDimensions
vi.mock('../../../src/ppt-tools/slideLayoutInfo', () => ({
  getSlideDimensions: vi.fn().mockResolvedValue({
    width: 720,
    height: 540,
  }),
}));

// 创建 mock 形状对象
const createMockShape = (): MockShape => ({
  id: 'mock-shape-id-123',
  type: 'GeometricShape',
  width: 100,
  height: 100,
  left: 100,
  top: 100,
  fill: {
    color: '',
    setSolidColor: function(color: string) {
      this.color = color;
    },
  },
  lineFormat: {
    color: '',
    weight: 0,
  },
  textFrame: {
    textRange: {
      text: '',
    },
  },
  load: vi.fn(),
});

// 创建 mock PowerPoint 数据
const createMockData = (): MockData => {
  const mockShape = createMockShape();
  return {
    context: {
      presentation: {
        getSelectedSlides: () => ({
          getItemAt: () => ({
            shapes: {
              addGeometricShape: (shapeType: string, options?: unknown) => {
                mockShape._shapeType = shapeType;
                mockShape._options = options;
                if (options && typeof options === 'object') {
                  const opts = options as any;
                  mockShape.left = opts.left ?? mockShape.left;
                  mockShape.top = opts.top ?? mockShape.top;
                  mockShape.width = opts.width ?? mockShape.width;
                  mockShape.height = opts.height ?? mockShape.height;
                }
                return mockShape;
              },
            },
          }),
        }),
      },
      sync: vi.fn().mockResolvedValue(undefined),
    },
    run: async function(callback: (context: MockData['context']) => Promise<any>) {
      return await callback(this.context);
    },
    _getShape: () => mockShape,
  };
};

// 设置 PowerPoint mock 的辅助函数
const setupPowerPointMock = (mockData: MockData) => {
  const mockPowerPoint = new OfficeMockObject(mockData);
  mockPowerPoint.GeometricShapeType = mockGeometricShapeType;
  (global as any).PowerPoint = mockPowerPoint;
};

describe('shapeInsertion 工具测试 | shapeInsertion Tool Tests', () => {
  beforeEach(() => {
    // 重置 global.PowerPoint
    delete (global as any).PowerPoint;
    
    // 设置 PowerPoint.GeometricShapeType
    (global as any).PowerPoint = {
      GeometricShapeType: mockGeometricShapeType,
    };
    
    vi.clearAllMocks();
  });

  describe('insertShapeToSlide', () => {
    it('应该能够插入带有默认参数的矩形 | should insert rectangle with default parameters', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShapeToSlide({ shapeType: 'rectangle' });

      const shape = mockData._getShape();
      expect(result.shapeId).toBe('mock-shape-id-123');
      expect(shape.fill.color).toBe('#4472C4');
      expect(shape.lineFormat.color).toBe('#2E5090');
      expect(shape.lineFormat.weight).toBe(2);
      expect(result.width).toBe(100);
      expect(result.height).toBe(100);
    });

    it('应该能够插入带有指定位置的形状 | should insert shape with specified position', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShapeToSlide({
        shapeType: 'ellipse',
        left: 200,
        top: 150,
      });

      expect(result.left).toBe(200);
      expect(result.top).toBe(150);
    });

    it('应该能够插入带有自定义尺寸的形状 | should insert shape with custom size', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShapeToSlide({
        shapeType: 'diamond',
        left: 50,
        top: 50,
        width: 200,
        height: 150,
      });

      expect(result.width).toBe(200);
      expect(result.height).toBe(150);
      expect(result.left).toBe(50);
      expect(result.top).toBe(50);
    });

    it('应该能够插入带有自定义颜色的形状 | should insert shape with custom colors', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      await insertShapeToSlide({
        shapeType: 'triangle',
        fillColor: '#FF0000',
        lineColor: '#00FF00',
        lineWeight: 5,
      });

      const shape = mockData._getShape();
      expect(shape.fill.color).toBe('#FF0000');
      expect(shape.lineFormat.color).toBe('#00FF00');
      expect(shape.lineFormat.weight).toBe(5);
    });

    it('应该能够插入带有文本的形状 | should insert shape with text', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      await insertShapeToSlide({
        shapeType: 'rectangle',
        text: '测试文本',
      });

      const shape = mockData._getShape();
      expect(shape.textFrame.textRange.text).toBe('测试文本');
    });

    it('应该能够插入带有完整配置的形状 | should insert shape with full configuration', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShapeToSlide({
        shapeType: 'star5',
        left: 100,
        top: 200,
        width: 150,
        height: 150,
        fillColor: '#FFD700',
        lineColor: '#FFA500',
        lineWeight: 3,
        text: '五角星',
      });

      const shape = mockData._getShape();
      expect(result.left).toBe(100);
      expect(result.top).toBe(200);
      expect(result.width).toBe(150);
      expect(result.height).toBe(150);
      expect(shape.fill.color).toBe('#FFD700');
      expect(shape.lineFormat.color).toBe('#FFA500');
      expect(shape.lineFormat.weight).toBe(3);
      expect(shape.textFrame.textRange.text).toBe('五角星');
    });

    it('应该在未指定位置时使用居中位置 | should use centered position when position not specified', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShapeToSlide({
        shapeType: 'rectangle',
        width: 100,
        height: 100,
      });

      // 居中位置应该是 (720-100)/2 = 310, (540-100)/2 = 220
      expect(result.left).toBe(310);
      expect(result.top).toBe(220);
    });

    it('应该能够插入不同类型的形状 | should insert different types of shapes', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const shapeTypes = ['rectangle', 'ellipse', 'triangle', 'star5', 'hexagon'] as const;

      for (const shapeType of shapeTypes) {
        await insertShapeToSlide({ shapeType });
        const shape = mockData._getShape();
        expect(shape._shapeType).toBeDefined();
      }
    });

    it('应该在插入失败时抛出错误 | should throw error on insertion failure', async () => {
      const mockData = {
        context: {
          presentation: {
            getSelectedSlides: function() {
              throw new Error('插入失败');
            },
          },
        },
        run: async function(callback: (context: typeof mockData.context) => Promise<void>) {
          await callback(this.context);
        },
      };

      setupPowerPointMock(mockData);

      await expect(insertShapeToSlide({ shapeType: 'rectangle' })).rejects.toThrow();
    });
  });

  describe('insertShape', () => {
    it('应该能够插入简单形状 | should insert simple shape', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShape('rectangle');

      expect(result.shapeId).toBe('mock-shape-id-123');
      expect(result.width).toBe(100);
      expect(result.height).toBe(100);
    });

    it('应该能够插入带有位置的形状 | should insert shape with position', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShape('ellipse', 150, 250);

      expect(result.left).toBe(150);
      expect(result.top).toBe(250);
    });

    it('应该能够插入带有位置和尺寸的形状 | should insert shape with position and size', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShape('diamond', 100, 200, 300, 250);

      expect(result.left).toBe(100);
      expect(result.top).toBe(200);
      expect(result.width).toBe(300);
      expect(result.height).toBe(250);
    });

    it('应该正确调用 insertShapeToSlide | should correctly call insertShapeToSlide', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const shapeType = 'triangle';
      const left = 100;
      const top = 200;
      const width = 150;
      const height = 120;

      const result = await insertShape(shapeType, left, top, width, height);

      expect(result.left).toBe(left);
      expect(result.top).toBe(top);
      expect(result.width).toBe(width);
      expect(result.height).toBe(height);
    });
  });

  describe('边界情况测试 | Edge Cases Tests', () => {
    it('应该能够处理零坐标 | should handle zero coordinates', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShape('rectangle', 0, 0);

      expect(result.left).toBe(0);
      expect(result.top).toBe(0);
    });

    it('应该能够处理负坐标 | should handle negative coordinates', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShape('rectangle', -10, -20);

      expect(result.left).toBe(-10);
      expect(result.top).toBe(-20);
    });

    it('应该能够处理非常大的尺寸 | should handle very large sizes', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShape('rectangle', 0, 0, 1000, 800);

      expect(result.width).toBe(1000);
      expect(result.height).toBe(800);
    });

    it('应该能够处理非常小的尺寸 | should handle very small sizes', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const result = await insertShape('rectangle', 0, 0, 1, 1);

      expect(result.width).toBe(1);
      expect(result.height).toBe(1);
    });

    it('应该能够插入空文本 | should insert shape with empty text', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      await insertShapeToSlide({
        shapeType: 'rectangle',
        text: '',
      });

      const shape = mockData._getShape();
      expect(shape.textFrame.textRange.text).toBe('');
    });

    it('应该能够插入包含特殊字符的文本 | should insert shape with special characters in text', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const specialText = '特殊字符 !@#$%^&*() 测试';
      await insertShapeToSlide({
        shapeType: 'rectangle',
        text: specialText,
      });

      const shape = mockData._getShape();
      expect(shape.textFrame.textRange.text).toBe(specialText);
    });

    it('应该能够插入多行文本 | should insert shape with multiline text', async () => {
      const mockData = createMockData();
      setupPowerPointMock(mockData);

      const multilineText = '第一行\n第二行\n第三行';
      await insertShapeToSlide({
        shapeType: 'rectangle',
        text: multilineText,
      });

      const shape = mockData._getShape();
      expect(shape.textFrame.textRange.text).toBe(multilineText);
    });
  });

  describe('COMMON_SHAPES 常量测试 | COMMON_SHAPES Constant Tests', () => {
    it('应该包含所有基础形状 | should contain all basic shapes', () => {
      const basicShapes = COMMON_SHAPES.filter(s => s.category === '基础形状');
      expect(basicShapes.length).toBeGreaterThan(0);
      expect(basicShapes.some(s => s.type === 'rectangle')).toBe(true);
      expect(basicShapes.some(s => s.type === 'ellipse')).toBe(true);
      expect(basicShapes.some(s => s.type === 'triangle')).toBe(true);
    });

    it('应该包含所有星形 | should contain all star shapes', () => {
      const starShapes = COMMON_SHAPES.filter(s => s.category === '星形');
      expect(starShapes.length).toBeGreaterThan(0);
      expect(starShapes.some(s => s.type === 'star5')).toBe(true);
      expect(starShapes.some(s => s.type === 'star6')).toBe(true);
    });

    it('应该包含所有箭头 | should contain all arrow shapes', () => {
      const arrowShapes = COMMON_SHAPES.filter(s => s.category === '箭头');
      expect(arrowShapes.length).toBeGreaterThan(0);
      expect(arrowShapes.some(s => s.type === 'leftRightArrow')).toBe(true);
    });

    it('应该包含所有标注 | should contain all callout shapes', () => {
      const calloutShapes = COMMON_SHAPES.filter(s => s.category === '标注');
      expect(calloutShapes.length).toBeGreaterThan(0);
      expect(calloutShapes.some(s => s.type === 'rightArrowCallout')).toBe(true);
    });

    it('应该包含所有装饰形状 | should contain all decorative shapes', () => {
      const decorativeShapes = COMMON_SHAPES.filter(s => s.category === '装饰形状');
      expect(decorativeShapes.length).toBeGreaterThan(0);
      expect(decorativeShapes.some(s => s.type === 'plaque')).toBe(true);
    });

    it('应该包含所有数学符号 | should contain all math symbols', () => {
      const mathShapes = COMMON_SHAPES.filter(s => s.category === '数学符号');
      expect(mathShapes.length).toBeGreaterThan(0);
      expect(mathShapes.some(s => s.type === 'mathPlus')).toBe(true);
    });

    it('每个形状应该有必需的属性 | each shape should have required properties', () => {
      COMMON_SHAPES.forEach(shape => {
        expect(shape).toHaveProperty('type');
        expect(shape).toHaveProperty('label');
        expect(shape).toHaveProperty('category');
        expect(typeof shape.type).toBe('string');
        expect(typeof shape.label).toBe('string');
        expect(typeof shape.category).toBe('string');
      });
    });
  });
});
