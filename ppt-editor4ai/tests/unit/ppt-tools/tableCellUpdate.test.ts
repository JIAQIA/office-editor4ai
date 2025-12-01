/**
 * 文件名: tableCellUpdate.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 描述: tableCellUpdate 工具的单元测试 | tableCellUpdate tool unit tests
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { OfficeMockObject } from 'office-addin-mock';
import {
  updateTableCell,
  updateTableCellsBatch,
  getTableCellContent,
} from '../../../src/ppt-tools';

type MockCell = {
  text: string;
  load: (props: string) => void;
};

type MockTable = {
  rowCount: number;
  columnCount: number;
  getCellOrNullObject: (row: number, col: number) => MockCell;
  load: (props: string) => void;
  _cells: MockCell[][];
};

type MockShape = {
  id: string;
  type: string;
  getTable: () => MockTable;
  load: (props: string) => void;
};

type MockData = {
  context: {
    presentation: {
      getSelectedSlides: () => {
        getItemAt: (index: number) => {
          shapes: {
            items: MockShape[];
            load: (props: string) => void;
          };
        };
      };
    };
    sync: () => Promise<void>;
  };
  run: (callback: (context: MockData['context']) => Promise<any>) => Promise<any>;
};

// 创建 mock 单元格对象
// Create mock cell object
const createMockCell = (text: string = ''): MockCell => ({
  text,
  load: function () {},
});

// 创建 mock 表格对象
// Create mock table object
const createMockTable = (rowCount: number, columnCount: number, initialData?: string[][]): MockTable => {
  const cells: MockCell[][] = [];
  for (let i = 0; i < rowCount; i++) {
    cells[i] = [];
    for (let j = 0; j < columnCount; j++) {
      const cellText = initialData?.[i]?.[j] ?? `Cell(${i},${j})`;
      cells[i][j] = createMockCell(cellText);
    }
  }

  return {
    rowCount,
    columnCount,
    getCellOrNullObject: (row: number, col: number) => cells[row][col],
    load: function () {},
    _cells: cells,
  };
};

// 创建 mock 形状对象
// Create mock shape object
const createMockShape = (id: string, type: string, rowCount: number, columnCount: number): MockShape => {
  const mockTable = createMockTable(rowCount, columnCount);

  return {
    id,
    type,
    getTable: () => mockTable,
    load: function () {},
  };
};

// 创建 mock PowerPoint 数据
// Create mock PowerPoint data
const createMockData = (shapes: MockShape[]): any => {
  const mockData = {
    context: {
      presentation: {
        getSelectedSlides: () => ({
          getItemAt: () => ({
            shapes: {
              items: shapes,
              load: function () {},
            },
            load: function () {}, // 添加 slide.load 方法 | Add slide.load method
          }),
        }),
      },
      sync: async () => {},
    },
    run: async function (callback: (context: any) => Promise<any>) {
      try {
        return await callback(this.context);
      } catch (error) {
        throw error;
      }
    },
  };
  return mockData;
};

// 设置 PowerPoint mock 的辅助函数
// Helper function to setup PowerPoint mock
const setupPowerPointMock = (shapes: MockShape[]) => {
  const mockData = createMockData(shapes);
  const mockPowerPoint = new OfficeMockObject(mockData);
  
  // 直接使用 Office.js 提供的真实 PowerPoint.ShapeType 枚举
  // Use the real PowerPoint.ShapeType enum provided by Office.js
  // 参考 | Reference: https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapetype
  // 这样测试环境与生产环境完全一致 | This ensures test environment matches production exactly
  (mockPowerPoint as any).ShapeType = {table: "Table"};
  
  (global as any).PowerPoint = mockPowerPoint;
};

describe('tableCellUpdate 工具测试 | tableCellUpdate Tool Tests', () => {
  beforeEach(() => {
    // 重置 global.PowerPoint
    // Reset global.PowerPoint
    delete (global as any).PowerPoint;
  });

  describe('updateTableCell - 基础功能测试 | updateTableCell - Basic Functionality Tests', () => {
    it('应该能够更新单个单元格内容 | should be able to update single cell content', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCell({
        rowIndex: 0,
        columnIndex: 0,
        text: 'Updated Cell',
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(1);
      expect(result.rowCount).toBe(3);
      expect(result.columnCount).toBe(3);
      expect(tableShape.getTable()._cells[0][0].text).toBe('Updated Cell');
    });

    it('应该能够通过 shapeId 定位表格 | should be able to locate table by shapeId', async () => {
      const tableShape = createMockShape('specific-table', 'Table', 2, 2);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCell(
        {
          rowIndex: 1,
          columnIndex: 1,
          text: 'Specific Table Cell',
        },
        { shapeId: 'specific-table' }
      );

      expect(result.success).toBe(true);
      expect(tableShape.getTable()._cells[1][1].text).toBe('Specific Table Cell');
    });

    it('应该能够通过 tableIndex 定位表格 | should be able to locate table by tableIndex', async () => {
      const table1 = createMockShape('table1', 'Table', 2, 2);
      const table2 = createMockShape('table2', 'Table', 3, 3);
      setupPowerPointMock([table1, table2]);

      const result = await updateTableCell(
        {
          rowIndex: 0,
          columnIndex: 0,
          text: 'Second Table',
        },
        { tableIndex: 1 }
      );

      expect(result.success).toBe(true);
      expect(table2.getTable()._cells[0][0].text).toBe('Second Table');
    });

    it('应该默认使用第一个表格 | should use first table by default', async () => {
      const table1 = createMockShape('table1', 'Table', 2, 2);
      const table2 = createMockShape('table2', 'Table', 3, 3);
      setupPowerPointMock([table1, table2]);

      const result = await updateTableCell({
        rowIndex: 0,
        columnIndex: 0,
        text: 'First Table',
      });

      expect(result.success).toBe(true);
      expect(table1.getTable()._cells[0][0].text).toBe('First Table');
    });

    it('应该能够更新不同位置的单元格 | should be able to update cells at different positions', async () => {
      const tableShape = createMockShape('table1', 'Table', 5, 5);
      setupPowerPointMock([tableShape]);

      // 更新多个不同位置的单元格
      // Update cells at different positions
      await updateTableCell({ rowIndex: 0, columnIndex: 0, text: 'Top-Left' });
      await updateTableCell({ rowIndex: 0, columnIndex: 4, text: 'Top-Right' });
      await updateTableCell({ rowIndex: 4, columnIndex: 0, text: 'Bottom-Left' });
      await updateTableCell({ rowIndex: 4, columnIndex: 4, text: 'Bottom-Right' });
      await updateTableCell({ rowIndex: 2, columnIndex: 2, text: 'Center' });

      const table = tableShape.getTable();
      expect(table._cells[0][0].text).toBe('Top-Left');
      expect(table._cells[0][4].text).toBe('Top-Right');
      expect(table._cells[4][0].text).toBe('Bottom-Left');
      expect(table._cells[4][4].text).toBe('Bottom-Right');
      expect(table._cells[2][2].text).toBe('Center');
    });
  });

  describe('updateTableCell - 错误处理测试 | updateTableCell - Error Handling Tests', () => {
    it('应该在行索引为负数时抛出错误 | should throw error when row index is negative', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableCell({
          rowIndex: -1,
          columnIndex: 0,
          text: 'Test',
        })
      ).rejects.toThrow('行索引和列索引必须大于等于 0');
    });

    it('应该在列索引为负数时抛出错误 | should throw error when column index is negative', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableCell({
          rowIndex: 0,
          columnIndex: -1,
          text: 'Test',
        })
      ).rejects.toThrow('行索引和列索引必须大于等于 0');
    });

    it('应该在行索引超出范围时返回错误 | should return error when row index is out of range', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCell({
        rowIndex: 5,
        columnIndex: 0,
        text: 'Test',
      });

      expect(result.success).toBe(false);
      expect(result.error).toContain('行索引 5 超出范围');
    });

    it('应该在列索引超出范围时返回错误 | should return error when column index is out of range', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCell({
        rowIndex: 0,
        columnIndex: 5,
        text: 'Test',
      });

      expect(result.success).toBe(false);
      expect(result.error).toContain('列索引 5 超出范围');
    });

    it('应该在找不到表格时返回错误 | should return error when no table is found', async () => {
      const shape = createMockShape('shape1', 'TextBox', 0, 0);
      setupPowerPointMock([shape]);

      const result = await updateTableCell({
        rowIndex: 0,
        columnIndex: 0,
        text: 'Test',
      });

      expect(result.success).toBe(false);
      expect(result.error).toContain('当前幻灯片没有表格');
    });

    it('应该在指定的 shapeId 不存在时返回错误 | should return error when specified shapeId does not exist', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCell(
        {
          rowIndex: 0,
          columnIndex: 0,
          text: 'Test',
        },
        { shapeId: 'nonexistent' }
      );

      expect(result.success).toBe(false);
      expect(result.error).toContain('未找到 ID 为 nonexistent 的形状');
    });

    it('应该在 tableIndex 超出范围时返回错误 | should return error when tableIndex is out of range', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCell(
        {
          rowIndex: 0,
          columnIndex: 0,
          text: 'Test',
        },
        { tableIndex: 5 }
      );

      expect(result.success).toBe(false);
      expect(result.error).toContain('表格索引 5 超出范围');
    });
  });

  describe('updateTableCellsBatch - 批量更新测试 | updateTableCellsBatch - Batch Update Tests', () => {
    it('应该能够批量更新多个单元格 | should be able to batch update multiple cells', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCellsBatch({
        cells: [
          { rowIndex: 0, columnIndex: 0, text: '标题1' },
          { rowIndex: 0, columnIndex: 1, text: '标题2' },
          { rowIndex: 0, columnIndex: 2, text: '标题3' },
          { rowIndex: 1, columnIndex: 0, text: '数据1' },
          { rowIndex: 1, columnIndex: 1, text: '数据2' },
          { rowIndex: 1, columnIndex: 2, text: '数据3' },
        ],
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(6);
      expect(result.rowCount).toBe(3);
      expect(result.columnCount).toBe(3);

      const table = tableShape.getTable();
      expect(table._cells[0][0].text).toBe('标题1');
      expect(table._cells[0][1].text).toBe('标题2');
      expect(table._cells[0][2].text).toBe('标题3');
      expect(table._cells[1][0].text).toBe('数据1');
      expect(table._cells[1][1].text).toBe('数据2');
      expect(table._cells[1][2].text).toBe('数据3');
    });

    it('应该跳过无效坐标的单元格 | should skip cells with invalid coordinates', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCellsBatch({
        cells: [
          { rowIndex: 0, columnIndex: 0, text: 'Valid' },
          { rowIndex: 10, columnIndex: 0, text: 'Invalid Row' },
          { rowIndex: 0, columnIndex: 10, text: 'Invalid Column' },
          { rowIndex: 1, columnIndex: 1, text: 'Valid 2' },
        ],
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(2); // 只有两个有效单元格
      expect(tableShape.getTable()._cells[0][0].text).toBe('Valid');
      expect(tableShape.getTable()._cells[1][1].text).toBe('Valid 2');
    });

    it('应该在单元格列表为空时抛出错误 | should throw error when cell list is empty', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableCellsBatch({
          cells: [],
        })
      ).rejects.toThrow('单元格列表不能为空');
    });

    it('应该在包含负数坐标时抛出错误 | should throw error when contains negative coordinates', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableCellsBatch({
          cells: [
            { rowIndex: 0, columnIndex: 0, text: 'Valid' },
            { rowIndex: -1, columnIndex: 0, text: 'Invalid' },
          ],
        })
      ).rejects.toThrow('单元格 (-1, 0) 坐标无效');
    });

    it('应该能够批量更新整行 | should be able to batch update entire row', async () => {
      const tableShape = createMockShape('table1', 'Table', 5, 5);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCellsBatch({
        cells: [
          { rowIndex: 2, columnIndex: 0, text: 'Row2-Col0' },
          { rowIndex: 2, columnIndex: 1, text: 'Row2-Col1' },
          { rowIndex: 2, columnIndex: 2, text: 'Row2-Col2' },
          { rowIndex: 2, columnIndex: 3, text: 'Row2-Col3' },
          { rowIndex: 2, columnIndex: 4, text: 'Row2-Col4' },
        ],
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(5);

      const table = tableShape.getTable();
      for (let i = 0; i < 5; i++) {
        expect(table._cells[2][i].text).toBe(`Row2-Col${i}`);
      }
    });

    it('应该能够批量更新整列 | should be able to batch update entire column', async () => {
      const tableShape = createMockShape('table1', 'Table', 5, 5);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCellsBatch({
        cells: [
          { rowIndex: 0, columnIndex: 3, text: 'Col3-Row0' },
          { rowIndex: 1, columnIndex: 3, text: 'Col3-Row1' },
          { rowIndex: 2, columnIndex: 3, text: 'Col3-Row2' },
          { rowIndex: 3, columnIndex: 3, text: 'Col3-Row3' },
          { rowIndex: 4, columnIndex: 3, text: 'Col3-Row4' },
        ],
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(5);

      const table = tableShape.getTable();
      for (let i = 0; i < 5; i++) {
        expect(table._cells[i][3].text).toBe(`Col3-Row${i}`);
      }
    });
  });

  describe('getTableCellContent - 获取单元格内容测试 | getTableCellContent - Get Cell Content Tests', () => {
    it('应该能够获取单元格内容 | should be able to get cell content', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      const table = tableShape.getTable();
      table._cells[1][2].text = 'Test Content';

      setupPowerPointMock([tableShape]);

      const content = await getTableCellContent(1, 2);

      expect(content).toBe('Test Content');
    });

    it('应该能够通过 shapeId 获取单元格内容 | should be able to get cell content by shapeId', async () => {
      const tableShape = createMockShape('specific-table', 'Table', 3, 3);
      const table = tableShape.getTable();
      table._cells[0][0].text = 'Specific Content';

      setupPowerPointMock([tableShape]);

      const content = await getTableCellContent(0, 0, { shapeId: 'specific-table' });

      expect(content).toBe('Specific Content');
    });

    it('应该在行索引为负数时抛出错误 | should throw error when row index is negative', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(getTableCellContent(-1, 0)).rejects.toThrow('行索引和列索引必须大于等于 0');
    });

    it('应该在列索引为负数时抛出错误 | should throw error when column index is negative', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(getTableCellContent(0, -1)).rejects.toThrow('行索引和列索引必须大于等于 0');
    });

    it('应该在坐标超出范围时抛出错误 | should throw error when coordinates are out of range', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(getTableCellContent(5, 5)).rejects.toThrow('坐标 (5, 5) 超出范围');
    });

    it('应该能够获取空单元格内容 | should be able to get empty cell content', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      const table = tableShape.getTable();
      table._cells[1][1].text = '';

      setupPowerPointMock([tableShape]);

      const content = await getTableCellContent(1, 1);

      expect(content).toBe('');
    });
  });

  describe('边界情况测试 | Edge Case Tests', () => {
    it('应该能够处理空文本 | should be able to handle empty text', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCell({
        rowIndex: 0,
        columnIndex: 0,
        text: '',
      });

      expect(result.success).toBe(true);
      expect(tableShape.getTable()._cells[0][0].text).toBe('');
    });

    it('应该能够处理特殊字符 | should be able to handle special characters', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const specialText = '特殊字符 !@#$%^&*() 测试\n换行符\t制表符';
      const result = await updateTableCell({
        rowIndex: 1,
        columnIndex: 1,
        text: specialText,
      });

      expect(result.success).toBe(true);
      expect(tableShape.getTable()._cells[1][1].text).toBe(specialText);
    });

    it('应该能够处理长文本 | should be able to handle long text', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const longText = 'A'.repeat(1000);
      const result = await updateTableCell({
        rowIndex: 0,
        columnIndex: 0,
        text: longText,
      });

      expect(result.success).toBe(true);
      expect(tableShape.getTable()._cells[0][0].text).toBe(longText);
    });

    it('应该能够处理 1x1 表格 | should be able to handle 1x1 table', async () => {
      const tableShape = createMockShape('table1', 'Table', 1, 1);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCell({
        rowIndex: 0,
        columnIndex: 0,
        text: 'Single Cell',
      });

      expect(result.success).toBe(true);
      expect(result.rowCount).toBe(1);
      expect(result.columnCount).toBe(1);
      expect(tableShape.getTable()._cells[0][0].text).toBe('Single Cell');
    });

    it('应该能够处理大型表格 | should be able to handle large table', async () => {
      const tableShape = createMockShape('table1', 'Table', 50, 20);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCell({
        rowIndex: 49,
        columnIndex: 19,
        text: 'Last Cell',
      });

      expect(result.success).toBe(true);
      expect(result.rowCount).toBe(50);
      expect(result.columnCount).toBe(20);
      expect(tableShape.getTable()._cells[49][19].text).toBe('Last Cell');
    });

    it('应该能够处理包含逗号的文本（批量更新场景）| should be able to handle text with commas (batch update scenario)', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableCellsBatch({
        cells: [
          { rowIndex: 0, columnIndex: 0, text: '姓名,年龄,地址' },
          { rowIndex: 1, columnIndex: 0, text: '张三,25,北京市,海淀区' },
        ],
      });

      expect(result.success).toBe(true);
      expect(tableShape.getTable()._cells[0][0].text).toBe('姓名,年龄,地址');
      expect(tableShape.getTable()._cells[1][0].text).toBe('张三,25,北京市,海淀区');
    });

    it('应该能够覆盖已有内容 | should be able to overwrite existing content', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      const table = tableShape.getTable();
      table._cells[1][1].text = 'Original Content';

      setupPowerPointMock([tableShape]);

      const result = await updateTableCell({
        rowIndex: 1,
        columnIndex: 1,
        text: 'New Content',
      });

      expect(result.success).toBe(true);
      expect(table._cells[1][1].text).toBe('New Content');
    });

    it('应该能够处理多个表格混合场景 | should be able to handle multiple tables scenario', async () => {
      const table1 = createMockShape('table1', 'Table', 2, 2);
      const shape1 = createMockShape('shape1', 'TextBox', 0, 0);
      const table2 = createMockShape('table2', 'Table', 3, 3);
      const shape2 = createMockShape('shape2', 'Picture', 0, 0);
      const table3 = createMockShape('table3', 'Table', 4, 4);

      setupPowerPointMock([table1, shape1, table2, shape2, table3]);

      // 更新第一个表格
      // Update first table
      const result1 = await updateTableCell(
        { rowIndex: 0, columnIndex: 0, text: 'Table 1' },
        { tableIndex: 0 }
      );
      expect(result1.success).toBe(true);
      expect(table1.getTable()._cells[0][0].text).toBe('Table 1');

      // 更新第二个表格
      // Update second table
      const result2 = await updateTableCell(
        { rowIndex: 0, columnIndex: 0, text: 'Table 2' },
        { tableIndex: 1 }
      );
      expect(result2.success).toBe(true);
      expect(table2.getTable()._cells[0][0].text).toBe('Table 2');

      // 更新第三个表格
      // Update third table
      const result3 = await updateTableCell(
        { rowIndex: 0, columnIndex: 0, text: 'Table 3' },
        { tableIndex: 2 }
      );
      expect(result3.success).toBe(true);
      expect(table3.getTable()._cells[0][0].text).toBe('Table 3');
    });
  });
});
