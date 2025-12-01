/**
 * 文件名: tableRowColumnUpdate.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 描述: tableRowColumnUpdate 工具的单元测试 | tableRowColumnUpdate tool unit tests
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { OfficeMockObject } from 'office-addin-mock';
import {
  updateTableRow,
  updateTableColumn,
  updateTableRowsBatch,
  updateTableColumnsBatch,
  getTableRowContent,
  getTableColumnContent,
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
          load: (props: string) => void;
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
            load: function () {},
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
  
  (mockPowerPoint as any).ShapeType = {table: "Table"};
  (mockPowerPoint as any).TableCell = {};
  
  (global as any).PowerPoint = mockPowerPoint;
};

describe('tableRowColumnUpdate 工具测试 | tableRowColumnUpdate Tool Tests', () => {
  beforeEach(() => {
    delete (global as any).PowerPoint;
  });

  describe('updateTableRow - 更新整行测试 | updateTableRow - Update Row Tests', () => {
    it('应该能够更新整行内容 | should be able to update entire row', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableRow({
        rowIndex: 1,
        values: ['A', 'B', 'C'],
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(3);
      const table = tableShape.getTable();
      expect(table._cells[1][0].text).toBe('A');
      expect(table._cells[1][1].text).toBe('B');
      expect(table._cells[1][2].text).toBe('C');
    });

    it('应该能够跳过空值 | should be able to skip empty values', async () => {
      const tableShape = createMockShape('table1', 'Table', 2, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableRow({
        rowIndex: 0,
        values: ['A', '', 'C'],
        skipEmpty: true,
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(2);
      const table = tableShape.getTable();
      expect(table._cells[0][0].text).toBe('A');
      expect(table._cells[0][1].text).toBe('Cell(0,1)'); // 未更新 | Not updated
      expect(table._cells[0][2].text).toBe('C');
    });

    it('应该在行索引为负数时抛出错误 | should throw error when row index is negative', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableRow({
          rowIndex: -1,
          values: ['A', 'B', 'C'],
        })
      ).rejects.toThrow('行索引必须大于等于 0');
    });

    it('应该在行数据为空时抛出错误 | should throw error when row data is empty', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableRow({
          rowIndex: 0,
          values: [],
        })
      ).rejects.toThrow('行数据不能为空');
    });
  });

  describe('updateTableColumn - 更新整列测试 | updateTableColumn - Update Column Tests', () => {
    it('应该能够更新整列内容 | should be able to update entire column', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableColumn({
        columnIndex: 1,
        values: ['X', 'Y', 'Z'],
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(3);
      const table = tableShape.getTable();
      expect(table._cells[0][1].text).toBe('X');
      expect(table._cells[1][1].text).toBe('Y');
      expect(table._cells[2][1].text).toBe('Z');
    });

    it('应该能够跳过空值 | should be able to skip empty values', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 2);
      setupPowerPointMock([tableShape]);

      const result = await updateTableColumn({
        columnIndex: 0,
        values: ['X', '', 'Z'],
        skipEmpty: true,
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(2);
      const table = tableShape.getTable();
      expect(table._cells[0][0].text).toBe('X');
      expect(table._cells[1][0].text).toBe('Cell(1,0)'); // 未更新 | Not updated
      expect(table._cells[2][0].text).toBe('Z');
    });

    it('应该在列索引为负数时抛出错误 | should throw error when column index is negative', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableColumn({
          columnIndex: -1,
          values: ['X', 'Y', 'Z'],
        })
      ).rejects.toThrow('列索引必须大于等于 0');
    });

    it('应该在列数据为空时抛出错误 | should throw error when column data is empty', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableColumn({
          columnIndex: 0,
          values: [],
        })
      ).rejects.toThrow('列数据不能为空');
    });
  });

  describe('updateTableRowsBatch - 批量更新行测试 | updateTableRowsBatch - Batch Update Rows Tests', () => {
    it('应该能够批量更新多行 | should be able to batch update multiple rows', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableRowsBatch({
        rows: [
          { rowIndex: 0, values: ['A1', 'A2', 'A3'] },
          { rowIndex: 1, values: ['B1', 'B2', 'B3'] },
          { rowIndex: 2, values: ['C1', 'C2', 'C3'] },
        ],
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(9);
      const table = tableShape.getTable();
      expect(table._cells[0][0].text).toBe('A1');
      expect(table._cells[1][1].text).toBe('B2');
      expect(table._cells[2][2].text).toBe('C3');
    });

    it('应该在行列表为空时抛出错误 | should throw error when row list is empty', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableRowsBatch({
          rows: [],
        })
      ).rejects.toThrow('行列表不能为空');
    });
  });

  describe('updateTableColumnsBatch - 批量更新列测试 | updateTableColumnsBatch - Batch Update Columns Tests', () => {
    it('应该能够批量更新多列 | should be able to batch update multiple columns', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      const result = await updateTableColumnsBatch({
        columns: [
          { columnIndex: 0, values: ['A1', 'B1', 'C1'] },
          { columnIndex: 1, values: ['A2', 'B2', 'C2'] },
          { columnIndex: 2, values: ['A3', 'B3', 'C3'] },
        ],
      });

      expect(result.success).toBe(true);
      expect(result.cellsUpdated).toBe(9);
      const table = tableShape.getTable();
      expect(table._cells[0][0].text).toBe('A1');
      expect(table._cells[1][1].text).toBe('B2');
      expect(table._cells[2][2].text).toBe('C3');
    });

    it('应该在列列表为空时抛出错误 | should throw error when column list is empty', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(
        updateTableColumnsBatch({
          columns: [],
        })
      ).rejects.toThrow('列列表不能为空');
    });
  });

  describe('getTableRowContent - 获取行内容测试 | getTableRowContent - Get Row Content Tests', () => {
    it('应该能够获取整行内容 | should be able to get entire row content', async () => {
      const initialData = [
        ['A1', 'A2', 'A3'],
        ['B1', 'B2', 'B3'],
        ['C1', 'C2', 'C3'],
      ];
      
      // 创建带初始数据的表格
      // Create table with initial data
      const mockTable = createMockTable(3, 3, initialData);
      const tableShape = {
        id: 'table1',
        type: 'Table',
        getTable: () => mockTable,
        load: function () {},
      };
      setupPowerPointMock([tableShape]);

      const rowData = await getTableRowContent(1);

      expect(rowData).toEqual(['B1', 'B2', 'B3']);
    });

    it('应该在行索引为负数时抛出错误 | should throw error when row index is negative', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(getTableRowContent(-1)).rejects.toThrow('行索引必须大于等于 0');
    });
  });

  describe('getTableColumnContent - 获取列内容测试 | getTableColumnContent - Get Column Content Tests', () => {
    it('应该能够获取整列内容 | should be able to get entire column content', async () => {
      const initialData = [
        ['A1', 'A2', 'A3'],
        ['B1', 'B2', 'B3'],
        ['C1', 'C2', 'C3'],
      ];
      
      // 创建带初始数据的表格
      // Create table with initial data
      const mockTable = createMockTable(3, 3, initialData);
      const tableShape = {
        id: 'table1',
        type: 'Table',
        getTable: () => mockTable,
        load: function () {},
      };
      setupPowerPointMock([tableShape]);

      const columnData = await getTableColumnContent(1);

      expect(columnData).toEqual(['A2', 'B2', 'C2']);
    });

    it('应该在列索引为负数时抛出错误 | should throw error when column index is negative', async () => {
      const tableShape = createMockShape('table1', 'Table', 3, 3);
      setupPowerPointMock([tableShape]);

      await expect(getTableColumnContent(-1)).rejects.toThrow('列索引必须大于等于 0');
    });
  });

  describe('表格定位测试 | Table Location Tests', () => {
    it('应该能够通过 shapeId 定位表格 | should be able to locate table by shapeId', async () => {
      const tableShape = createMockShape('specific-table', 'Table', 2, 2);
      setupPowerPointMock([tableShape]);

      const result = await updateTableRow(
        {
          rowIndex: 0,
          values: ['X', 'Y'],
        },
        { shapeId: 'specific-table' }
      );

      expect(result.success).toBe(true);
    });

    it('应该能够通过 tableIndex 定位表格 | should be able to locate table by tableIndex', async () => {
      const table1 = createMockShape('table1', 'Table', 2, 2);
      const table2 = createMockShape('table2', 'Table', 3, 3);
      setupPowerPointMock([table1, table2]);

      const result = await updateTableColumn(
        {
          columnIndex: 0,
          values: ['X', 'Y', 'Z'],
        },
        { tableIndex: 1 }
      );

      expect(result.success).toBe(true);
      const table = table2.getTable();
      expect(table._cells[0][0].text).toBe('X');
    });

    it('应该默认使用第一个表格 | should use first table by default', async () => {
      const table1 = createMockShape('table1', 'Table', 2, 2);
      const table2 = createMockShape('table2', 'Table', 3, 3);
      setupPowerPointMock([table1, table2]);

      const result = await updateTableRow({
        rowIndex: 0,
        values: ['A', 'B'],
      });

      expect(result.success).toBe(true);
      const table = table1.getTable();
      expect(table._cells[0][0].text).toBe('A');
    });
  });
});
