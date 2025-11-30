/**
 * 文件名: tableInsertion.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: tableInsertion 工具的单元测试 | tableInsertion tool unit tests
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { OfficeMockObject } from 'office-addin-mock';
import { insertTableToSlide, insertTable, TABLE_TEMPLATES } from '../../../src/ppt-tools';

type MockCell = {
  borders: {
    bottom: { color: string };
    top: { color: string };
    left: { color: string };
    right: { color: string };
  };
  fill: {
    setSolidColor: (color: string) => void;
    _color?: string;
  };
  font: {
    color: string;
    bold: boolean;
  };
};

type MockTable = {
  getCellOrNullObject: (row: number, col: number) => MockCell;
  _cells: MockCell[][];
};

type MockTableShape = {
  id: string;
  width: number;
  height: number;
  left: number;
  top: number;
  load: (props: string) => void;
  getTable: () => MockTable;
  _rowCount?: number;
  _columnCount?: number;
  _options?: any;
};

type MockData = {
  context: {
    presentation: {
      getSelectedSlides: () => {
        getItemAt: (index: number) => {
          shapes: {
            addTable: (rowCount: number, columnCount: number, options?: any) => MockTableShape;
          };
        };
      };
    };
    sync: () => Promise<void>;
  };
  run: (callback: (context: MockData['context']) => Promise<any>) => Promise<any>;
  _getTableShape: () => MockTableShape;
};

// 创建 mock 单元格对象
const createMockCell = (): MockCell => ({
  borders: {
    bottom: { color: '' },
    top: { color: '' },
    left: { color: '' },
    right: { color: '' },
  },
  fill: {
    setSolidColor: function(color: string) {
      this._color = color;
    },
    _color: undefined,
  },
  font: {
    color: '',
    bold: false,
  },
});

// 创建 mock 表格对象
const createMockTable = (rowCount: number, columnCount: number): MockTable => {
  const cells: MockCell[][] = [];
  for (let i = 0; i < rowCount; i++) {
    cells[i] = [];
    for (let j = 0; j < columnCount; j++) {
      cells[i][j] = createMockCell();
    }
  }

  return {
    getCellOrNullObject: (row: number, col: number) => cells[row][col],
    _cells: cells,
  };
};

// 创建 mock 表格形状对象
const createMockTableShape = (rowCount: number, columnCount: number, options?: any): MockTableShape => {
  const mockTable = createMockTable(rowCount, columnCount);
  
  return {
    id: 'mock-table-id',
    width: options?.width ?? 400,
    height: options?.height ?? rowCount * 30,
    left: options?.left ?? 160,
    top: options?.top ?? 120,
    load: vi.fn(),
    getTable: () => mockTable,
    _rowCount: rowCount,
    _columnCount: columnCount,
    _options: options,
  };
};

// Mock getSlideDimensions
vi.mock('../../../src/ppt-tools/slideLayoutInfo', () => ({
  getSlideDimensions: vi.fn().mockResolvedValue({
    width: 720,
    height: 540,
  }),
}));

// 创建 mock PowerPoint 数据
const createMockData = (): MockData => {
  let mockTableShape: MockTableShape | null = null;

  return {
    context: {
      presentation: {
        getSelectedSlides: () => ({
          getItemAt: () => ({
            shapes: {
              addTable: (rowCount: number, columnCount: number, options?: any) => {
                mockTableShape = createMockTableShape(rowCount, columnCount, options);
                return mockTableShape;
              },
            },
          }),
        }),
      },
      sync: vi.fn().mockResolvedValue(undefined),
    },
    run: async function(callback: (context: MockData['context']) => Promise<any>) {
      const result = await callback(this.context);
      return result;
    },
    _getTableShape: () => mockTableShape!,
  };
};

describe('tableInsertion 工具测试 | tableInsertion tool tests', () => {
  beforeEach(() => {
    // 重置 global.PowerPoint
    delete (global as any).PowerPoint;
  });

  describe('insertTableToSlide', () => {
    it('应该能够插入带有默认参数的表格 | should insert table with default parameters', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTableToSlide({
        rowCount: 3,
        columnCount: 3,
      });

      expect(result.shapeId).toBe('mock-table-id');
      expect(result.rowCount).toBe(3);
      expect(result.columnCount).toBe(3);
      expect(result.width).toBe(400);
      
      const tableShape = mockData._getTableShape();
      expect(tableShape._rowCount).toBe(3);
      expect(tableShape._columnCount).toBe(3);
    });

    it('应该能够插入带有指定位置的表格 | should insert table with specified position', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTableToSlide({
        rowCount: 2,
        columnCount: 4,
        left: 100,
        top: 200,
      });

      expect(result.left).toBe(100);
      expect(result.top).toBe(200);
      
      const tableShape = mockData._getTableShape();
      expect(tableShape._options?.left).toBe(100);
      expect(tableShape._options?.top).toBe(200);
    });

    it('应该能够插入带有自定义尺寸的表格 | should insert table with custom size', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTableToSlide({
        rowCount: 4,
        columnCount: 5,
        width: 500,
        height: 200,
      });

      expect(result.width).toBe(500);
      expect(result.height).toBe(200);
      
      const tableShape = mockData._getTableShape();
      expect(tableShape._options?.width).toBe(500);
      expect(tableShape._options?.height).toBe(200);
    });

    it('应该能够插入带有数据的表格 | should insert table with data', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const values = [
        ['姓名', '年龄', '城市'],
        ['张三', '25', '北京'],
        ['李四', '30', '上海'],
      ];

      const result = await insertTableToSlide({
        rowCount: 3,
        columnCount: 3,
        values,
      });

      expect(result.rowCount).toBe(3);
      expect(result.columnCount).toBe(3);
      
      const tableShape = mockData._getTableShape();
      expect(tableShape._options?.values).toEqual(values);
    });

    it('应该能够设置表头样式 | should set header style', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTableToSlide({
        rowCount: 3,
        columnCount: 3,
        showHeader: true,
        headerColor: '#FF0000',
      });

      const tableShape = mockData._getTableShape();
      const table = tableShape.getTable();
      
      // 检查第一行（表头）的样式
      for (let j = 0; j < 3; j++) {
        const cell = table.getCellOrNullObject(0, j);
        expect(cell.fill._color).toBe('#FF0000');
        expect(cell.font.color).toBe('#FFFFFF');
        expect(cell.font.bold).toBe(true);
      }
    });

    it('应该能够设置边框颜色 | should set border color', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTableToSlide({
        rowCount: 2,
        columnCount: 2,
        borderColor: '#00FF00',
      });

      const tableShape = mockData._getTableShape();
      const table = tableShape.getTable();
      
      // 检查所有单元格的边框颜色
      for (let i = 0; i < 2; i++) {
        for (let j = 0; j < 2; j++) {
          const cell = table.getCellOrNullObject(i, j);
          expect(cell.borders.bottom.color).toBe('#00FF00');
          expect(cell.borders.top.color).toBe('#00FF00');
          expect(cell.borders.left.color).toBe('#00FF00');
          expect(cell.borders.right.color).toBe('#00FF00');
        }
      }
    });

    it('应该能够插入带有完整配置的表格 | should insert table with full configuration', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const values = [
        ['列1', '列2'],
        ['数据1', '数据2'],
      ];

      const result = await insertTableToSlide({
        rowCount: 2,
        columnCount: 2,
        left: 50,
        top: 100,
        width: 300,
        height: 80,
        values,
        showHeader: true,
        headerColor: '#4472C4',
        borderColor: '#D0D0D0',
      });

      expect(result.rowCount).toBe(2);
      expect(result.columnCount).toBe(2);
      expect(result.left).toBe(50);
      expect(result.top).toBe(100);
      expect(result.width).toBe(300);
      expect(result.height).toBe(80);
    });

    it('应该在行数为 0 时抛出错误 | should throw error when row count is 0', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await expect(
        insertTableToSlide({
          rowCount: 0,
          columnCount: 3,
        })
      ).rejects.toThrow('行数和列数必须大于 0');
    });

    it('应该在列数为负数时抛出错误 | should throw error when column count is negative', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await expect(
        insertTableToSlide({
          rowCount: 3,
          columnCount: -1,
        })
      ).rejects.toThrow('行数和列数必须大于 0');
    });

    it('应该在行数超过 100 时抛出错误 | should throw error when row count exceeds 100', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await expect(
        insertTableToSlide({
          rowCount: 101,
          columnCount: 3,
        })
      ).rejects.toThrow('表格过大：行数不能超过 100，列数不能超过 50');
    });

    it('应该在列数超过 50 时抛出错误 | should throw error when column count exceeds 50', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await expect(
        insertTableToSlide({
          rowCount: 3,
          columnCount: 51,
        })
      ).rejects.toThrow('表格过大：行数不能超过 100，列数不能超过 50');
    });

    it('应该在数据行数不匹配时抛出错误 | should throw error when data row count mismatch', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const values = [
        ['A', 'B'],
        ['C', 'D'],
      ];

      await expect(
        insertTableToSlide({
          rowCount: 3,
          columnCount: 2,
          values,
        })
      ).rejects.toThrow('数据行数 (2) 与指定行数 (3) 不匹配');
    });

    it('应该在数据列数不匹配时抛出错误 | should throw error when data column count mismatch', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const values = [
        ['A', 'B', 'C'],
        ['D', 'E'],
      ];

      await expect(
        insertTableToSlide({
          rowCount: 2,
          columnCount: 3,
          values,
        })
      ).rejects.toThrow('第 2 行数据列数 (2) 与指定列数 (3) 不匹配');
    });

    it('应该在不显示表头时不设置表头样式 | should not set header style when showHeader is false', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTableToSlide({
        rowCount: 3,
        columnCount: 3,
        showHeader: false,
        headerColor: '#FF0000',
      });

      const tableShape = mockData._getTableShape();
      const table = tableShape.getTable();
      
      // 检查第一行不应该有表头样式
      for (let j = 0; j < 3; j++) {
        const cell = table.getCellOrNullObject(0, j);
        expect(cell.fill._color).toBeUndefined();
        expect(cell.font.bold).toBe(false);
      }
    });

    it('应该在只指定 left 时居中 | should center when only left is specified', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTableToSlide({
        rowCount: 3,
        columnCount: 3,
        left: 100,
      });

      // 应该使用居中逻辑
      expect(result.left).toBe(100);
      expect(result.top).toBeDefined();
    });
  });

  describe('insertTable', () => {
    it('应该能够插入简单表格 | should insert simple table', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTable(3, 4);

      expect(result.rowCount).toBe(3);
      expect(result.columnCount).toBe(4);
      
      const tableShape = mockData._getTableShape();
      expect(tableShape._rowCount).toBe(3);
      expect(tableShape._columnCount).toBe(4);
    });

    it('应该能够插入带有位置的表格 | should insert table with position', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTable(2, 3, 150, 250);

      expect(result.rowCount).toBe(2);
      expect(result.columnCount).toBe(3);
      expect(result.left).toBe(150);
      expect(result.top).toBe(250);
    });

    it('应该能够插入带有尺寸的表格 | should insert table with size', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTable(3, 3, 100, 100, 500, 200);

      expect(result.width).toBe(500);
      expect(result.height).toBe(200);
    });

    it('应该正确调用 insertTableToSlide | should correctly call insertTableToSlide', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTable(2, 2, 50, 50, 300, 100);

      const tableShape = mockData._getTableShape();
      expect(tableShape._rowCount).toBe(2);
      expect(tableShape._columnCount).toBe(2);
      expect(tableShape._options?.left).toBe(50);
      expect(tableShape._options?.top).toBe(50);
      expect(tableShape._options?.width).toBe(300);
      expect(tableShape._options?.height).toBe(100);
    });
  });

  describe('TABLE_TEMPLATES', () => {
    it('应该包含预定义的表格模板 | should contain predefined table templates', () => {
      expect(TABLE_TEMPLATES).toBeDefined();
      expect(Array.isArray(TABLE_TEMPLATES)).toBe(true);
      expect(TABLE_TEMPLATES.length).toBeGreaterThan(0);
    });

    it('每个模板应该有必需的属性 | each template should have required properties', () => {
      TABLE_TEMPLATES.forEach((template) => {
        expect(template).toHaveProperty('id');
        expect(template).toHaveProperty('name');
        expect(template).toHaveProperty('rowCount');
        expect(template).toHaveProperty('columnCount');
        expect(template).toHaveProperty('description');
        
        expect(typeof template.id).toBe('string');
        expect(typeof template.name).toBe('string');
        expect(typeof template.rowCount).toBe('number');
        expect(typeof template.columnCount).toBe('number');
        expect(typeof template.description).toBe('string');
      });
    });

    it('模板的行列数应该在有效范围内 | template row and column counts should be within valid range', () => {
      TABLE_TEMPLATES.forEach((template) => {
        expect(template.rowCount).toBeGreaterThan(0);
        expect(template.rowCount).toBeLessThanOrEqual(100);
        expect(template.columnCount).toBeGreaterThan(0);
        expect(template.columnCount).toBeLessThanOrEqual(50);
      });
    });
  });

  describe('边界情况测试 | edge cases tests', () => {
    it('应该能够插入 1x1 的表格 | should insert 1x1 table', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTableToSlide({
        rowCount: 1,
        columnCount: 1,
      });

      expect(result.rowCount).toBe(1);
      expect(result.columnCount).toBe(1);
    });

    it('应该能够插入最大尺寸的表格 | should insert maximum size table', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTableToSlide({
        rowCount: 100,
        columnCount: 50,
      });

      expect(result.rowCount).toBe(100);
      expect(result.columnCount).toBe(50);
    });

    it('应该能够处理零坐标 | should handle zero coordinates', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await insertTableToSlide({
        rowCount: 2,
        columnCount: 2,
        left: 0,
        top: 0,
      });

      expect(result.left).toBe(0);
      expect(result.top).toBe(0);
    });

    it('应该能够处理空数据数组 | should handle empty values array', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const values = [
        ['', ''],
        ['', ''],
      ];

      const result = await insertTableToSlide({
        rowCount: 2,
        columnCount: 2,
        values,
      });

      expect(result.rowCount).toBe(2);
      expect(result.columnCount).toBe(2);
    });

    it('应该能够处理包含特殊字符的数据 | should handle data with special characters', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const values = [
        ['!@#$%', '^&*()'],
        ['测试', '数据'],
      ];

      const result = await insertTableToSlide({
        rowCount: 2,
        columnCount: 2,
        values,
      });

      expect(result.rowCount).toBe(2);
      expect(result.columnCount).toBe(2);
    });
  });
});
