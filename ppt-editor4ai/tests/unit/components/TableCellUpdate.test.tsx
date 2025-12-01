/**
 * 文件名: TableCellUpdate.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 描述: TableCellUpdate 组件单元测试 | TableCellUpdate component unit tests
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { renderWithProviders, userEvent } from '../../utils/test-utils';
import { TableCellUpdate } from '../../../src/taskpane/components/tools/TableCellUpdate';
import * as pptTools from '../../../src/ppt-tools';
import { OfficeMockObject } from 'office-addin-mock';

// Mock ppt-tools module
vi.mock('../../../src/ppt-tools', () => ({
  updateTableCell: vi.fn().mockResolvedValue({
    success: true,
    cellsUpdated: 1,
    rowCount: 3,
    columnCount: 3,
  }),
  updateTableCellsBatch: vi.fn().mockResolvedValue({
    success: true,
    cellsUpdated: 4,
    rowCount: 3,
    columnCount: 3,
  }),
  getTableCellContent: vi.fn().mockResolvedValue('Sample Cell Content'),
}));

describe('TableCellUpdate 组件单元测试 | TableCellUpdate Component Unit Tests', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    delete (global as any).PowerPoint;

    // 重新设置默认的 mock 实现
    // Reset default mock implementations
    vi.mocked(pptTools.updateTableCell).mockResolvedValue({
      success: true,
      cellsUpdated: 1,
      rowCount: 3,
      columnCount: 3,
    });
    vi.mocked(pptTools.updateTableCellsBatch).mockResolvedValue({
      success: true,
      cellsUpdated: 4,
      rowCount: 3,
      columnCount: 3,
    });
    vi.mocked(pptTools.getTableCellContent).mockResolvedValue('Sample Cell Content');
  });

  describe('组件渲染测试 | Component Rendering Tests', () => {
    it('应该正确渲染组件 | should render component correctly', () => {
      renderWithProviders(<TableCellUpdate />);

      // 验证标题
      // Verify title
      expect(screen.getByText('表格单元格更新')).toBeInTheDocument();
      expect(screen.getByText(/通过行列坐标修改表格单元格内容/)).toBeInTheDocument();

      // 验证按钮
      // Verify buttons
      expect(screen.getByRole('button', { name: /获取选中的表格/i })).toBeInTheDocument();
      expect(screen.getByRole('button', { name: /获取单元格内容/i })).toBeInTheDocument();
      expect(screen.getByRole('button', { name: /更新单元格/i })).toBeInTheDocument();
      expect(screen.getByRole('button', { name: /批量更新/i })).toBeInTheDocument();
    });

    it('应该显示表格定位输入框 | should display table location inputs', () => {
      renderWithProviders(<TableCellUpdate />);

      expect(screen.getByLabelText('表格形状 ID（可选）')).toBeInTheDocument();
      expect(screen.getByLabelText('表格索引（默认 0）')).toBeInTheDocument();
    });

    it('应该显示单元格坐标输入框 | should display cell coordinate inputs', () => {
      renderWithProviders(<TableCellUpdate />);

      expect(screen.getByLabelText('行索引')).toBeInTheDocument();
      expect(screen.getByLabelText('列索引')).toBeInTheDocument();
      expect(screen.getByLabelText('单元格内容')).toBeInTheDocument();
    });

    it('应该显示批量更新输入框 | should display batch update input', () => {
      renderWithProviders(<TableCellUpdate />);

      expect(screen.getByLabelText(/批量数据/)).toBeInTheDocument();
    });

    it('应该显示使用说明 | should display usage instructions', () => {
      renderWithProviders(<TableCellUpdate />);

      expect(screen.getByText('使用说明:')).toBeInTheDocument();
      expect(screen.getByText(/行列编号从 1 开始计数/)).toBeInTheDocument();
    });
  });

  describe('获取选中表格测试 | Get Selected Table Tests', () => {
    it('应该能够获取选中的表格 | should be able to get selected table', async () => {
      const user = userEvent.setup();

      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'table-123',
                  type: 'Table',
                  name: 'Table 1',
                  load: vi.fn(),
                  getTable: () => ({
                    rowCount: 5,
                    columnCount: 4,
                    load: vi.fn(),
                  }),
                },
              ],
              load: vi.fn(),
            }),
          },
          sync: vi.fn().mockResolvedValue(undefined),
        },
        run: async function (callback: any) {
          await callback(this.context);
        },
      };

      (global as any).PowerPoint = new OfficeMockObject(mockData);

      renderWithProviders(<TableCellUpdate />);

      const button = screen.getByRole('button', { name: /获取选中的表格/i });
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/已获取选中表格: 5 行 × 4 列/)).toBeInTheDocument();
      });

      // 验证 shapeId 已填充
      // Verify shapeId is filled
      const shapeIdInput = screen.getByLabelText('表格形状 ID（可选）') as HTMLInputElement;
      expect(shapeIdInput.value).toBe('table-123');
    });

    it('应该在未选中元素时显示错误 | should show error when no shape is selected', async () => {
      const user = userEvent.setup();

      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 0 }),
              items: [],
              load: vi.fn(),
            }),
          },
          sync: vi.fn().mockResolvedValue(undefined),
        },
        run: async function (callback: any) {
          await callback(this.context);
        },
      };

      (global as any).PowerPoint = new OfficeMockObject(mockData);

      renderWithProviders(<TableCellUpdate />);

      const button = screen.getByRole('button', { name: /获取选中的表格/i });
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('请先在幻灯片中选中一个表格')).toBeInTheDocument();
      });
    });

    it('应该在选中多个元素时显示错误 | should show error when multiple shapes are selected', async () => {
      const user = userEvent.setup();

      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 2 }),
              items: [],
              load: vi.fn(),
            }),
          },
          sync: vi.fn().mockResolvedValue(undefined),
        },
        run: async function (callback: any) {
          await callback(this.context);
        },
      };

      (global as any).PowerPoint = new OfficeMockObject(mockData);

      renderWithProviders(<TableCellUpdate />);

      const button = screen.getByRole('button', { name: /获取选中的表格/i });
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('请只选中一个表格')).toBeInTheDocument();
      });
    });

    it('应该在选中非表格元素时显示警告 | should show warning when non-table shape is selected', async () => {
      const user = userEvent.setup();

      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'shape-456',
                  type: 'TextBox',
                  name: 'Text Box 1',
                  load: vi.fn(),
                },
              ],
              load: vi.fn(),
            }),
          },
          sync: vi.fn().mockResolvedValue(undefined),
        },
        run: async function (callback: any) {
          await callback(this.context);
        },
      };

      (global as any).PowerPoint = new OfficeMockObject(mockData);

      renderWithProviders(<TableCellUpdate />);

      const button = screen.getByRole('button', { name: /获取选中的表格/i });
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/选中的元素类型是 "TextBox"，不是表格/)).toBeInTheDocument();
      });
    });
  });

  describe('获取单元格内容测试 | Get Cell Content Tests', () => {
    it('应该能够获取单元格内容 | should be able to get cell content', async () => {
      const user = userEvent.setup();

      vi.mocked(pptTools.getTableCellContent).mockResolvedValueOnce('测试内容');

      renderWithProviders(<TableCellUpdate />);

      // 输入行列索引
      // Input row and column indices
      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');

      await user.type(rowInput, '2');
      await user.type(colInput, '3');

      // 点击获取按钮
      // Click get button
      const getButton = screen.getByRole('button', { name: /获取单元格内容/i });
      await user.click(getButton);

      await waitFor(() => {
        expect(pptTools.getTableCellContent).toHaveBeenCalledWith(1, 2, { tableIndex: 0 });
      });

      await waitFor(() => {
        expect(screen.getByText(/已获取单元格 \(2, 3\) 的内容/)).toBeInTheDocument();
      });

      // 验证内容已填充到文本框
      // Verify content is filled in textarea
      const contentInput = screen.getByLabelText('单元格内容') as HTMLTextAreaElement;
      expect(contentInput.value).toBe('测试内容');
    });

    it('应该在行列索引无效时显示错误 | should show error when row/column index is invalid', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      await user.clear(rowInput);
      await user.type(rowInput, 'abc');

      const getButton = screen.getByRole('button', { name: /获取单元格内容/i });
      await user.click(getButton);

      await waitFor(() => {
        expect(screen.getByText('请输入有效的行列索引')).toBeInTheDocument();
      });
    });

    it('应该在获取失败时显示错误 | should show error when get fails', async () => {
      const user = userEvent.setup();

      vi.mocked(pptTools.getTableCellContent).mockRejectedValueOnce(new Error('获取失败'));

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');

      await user.type(rowInput, '1');
      await user.type(colInput, '1');

      const getButton = screen.getByRole('button', { name: /获取单元格内容/i });
      await user.click(getButton);

      await waitFor(() => {
        expect(screen.getByText(/获取单元格内容失败/)).toBeInTheDocument();
      });
    });
  });

  describe('更新单元格测试 | Update Cell Tests', () => {
    it('应该能够更新单个单元格 | should be able to update single cell', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      // 输入数据
      // Input data
      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');
      const contentInput = screen.getByLabelText('单元格内容');

      await user.type(rowInput, '2');
      await user.type(colInput, '3');
      await user.type(contentInput, '新内容');

      // 点击更新按钮
      // Click update button
      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTableCell).toHaveBeenCalledWith(
          { rowIndex: 1, columnIndex: 2, text: '新内容' },
          { tableIndex: 0 }
        );
      });

      await waitFor(() => {
        expect(screen.getByText(/成功更新单元格 \(2, 3\)/)).toBeInTheDocument();
      });
    });

    it('应该能够使用 shapeId 更新单元格 | should be able to update cell using shapeId', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const shapeIdInput = screen.getByLabelText('表格形状 ID（可选）');
      await user.type(shapeIdInput, 'table-123');

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');
      const contentInput = screen.getByLabelText('单元格内容');

      await user.type(rowInput, '1');
      await user.type(colInput, '1');
      await user.type(contentInput, 'Test');

      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTableCell).toHaveBeenCalledWith(
          { rowIndex: 0, columnIndex: 0, text: 'Test' },
          { shapeId: 'table-123' }
        );
      });
    });

    it('应该在内容为空时显示错误 | should show error when content is empty', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');

      await user.type(rowInput, '1');
      await user.type(colInput, '1');

      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(screen.getByText('请输入单元格内容')).toBeInTheDocument();
      });
    });

    it('应该在更新失败时显示错误 | should show error when update fails', async () => {
      const user = userEvent.setup();

      vi.mocked(pptTools.updateTableCell).mockResolvedValueOnce({
        success: false,
        cellsUpdated: 0,
        rowCount: 0,
        columnCount: 0,
        error: '更新失败',
      });

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');
      const contentInput = screen.getByLabelText('单元格内容');

      await user.type(rowInput, '1');
      await user.type(colInput, '1');
      await user.type(contentInput, 'Test');

      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(screen.getByText(/更新失败/)).toBeInTheDocument();
      });
    });

    it('应该在选中非表格元素时禁用更新按钮 | should disable update button when non-table shape is selected', async () => {
      const user = userEvent.setup();

      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'shape-456',
                  type: 'Picture',
                  name: 'Image 1',
                  load: vi.fn(),
                },
              ],
              load: vi.fn(),
            }),
          },
          sync: vi.fn().mockResolvedValue(undefined),
        },
        run: async function (callback: any) {
          await callback(this.context);
        },
      };

      (global as any).PowerPoint = new OfficeMockObject(mockData);

      renderWithProviders(<TableCellUpdate />);

      const getButton = screen.getByRole('button', { name: /获取选中的表格/i });
      await user.click(getButton);

      await waitFor(() => {
        const updateButton = screen.getByRole('button', { name: /更新单元格/i });
        expect(updateButton).toBeDisabled();
      });

      // 验证警告信息
      // Verify warning message
      expect(screen.getByText(/当前选中的元素不是表格，更新功能已禁用/)).toBeInTheDocument();
    });
  });

  describe('批量更新测试 | Batch Update Tests', () => {
    it('应该能够批量更新多个单元格 | should be able to batch update multiple cells', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const batchInput = screen.getByLabelText(/批量数据/);
      await user.type(batchInput, '1,1,标题1\n1,2,标题2\n2,1,数据1\n2,2,数据2');

      const batchButton = screen.getByRole('button', { name: /批量更新/i });
      await user.click(batchButton);

      await waitFor(() => {
        expect(pptTools.updateTableCellsBatch).toHaveBeenCalledWith(
          {
            cells: [
              { rowIndex: 0, columnIndex: 0, text: '标题1' },
              { rowIndex: 0, columnIndex: 1, text: '标题2' },
              { rowIndex: 1, columnIndex: 0, text: '数据1' },
              { rowIndex: 1, columnIndex: 1, text: '数据2' },
            ],
          },
          { tableIndex: 0 }
        );
      });

      await waitFor(() => {
        expect(screen.getByText(/成功批量更新 4 个单元格/)).toBeInTheDocument();
      });
    });

    it('应该能够处理包含逗号的文本 | should be able to handle text with commas', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const batchInput = screen.getByLabelText(/批量数据/);
      await user.type(batchInput, '1,1,姓名,年龄,地址');

      const batchButton = screen.getByRole('button', { name: /批量更新/i });
      await user.click(batchButton);

      await waitFor(() => {
        expect(pptTools.updateTableCellsBatch).toHaveBeenCalledWith(
          {
            cells: [{ rowIndex: 0, columnIndex: 0, text: '姓名,年龄,地址' }],
          },
          { tableIndex: 0 }
        );
      });
    });

    it('应该在批量数据为空时显示错误 | should show error when batch data is empty', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const batchButton = screen.getByRole('button', { name: /批量更新/i });
      await user.click(batchButton);

      await waitFor(() => {
        expect(screen.getByText('请输入批量更新数据')).toBeInTheDocument();
      });
    });

    it('应该在批量数据格式错误时显示错误 | should show error when batch data format is invalid', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const batchInput = screen.getByLabelText(/批量数据/);
      await user.type(batchInput, '1,1\n2,2,内容');

      const batchButton = screen.getByRole('button', { name: /批量更新/i });
      await user.click(batchButton);

      await waitFor(() => {
        expect(screen.getByText(/第 1 行格式错误/)).toBeInTheDocument();
      });
    });

    it('应该在行列索引无效时显示错误 | should show error when row/column index is invalid', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const batchInput = screen.getByLabelText(/批量数据/);
      await user.type(batchInput, 'abc,1,内容');

      const batchButton = screen.getByRole('button', { name: /批量更新/i });
      await user.click(batchButton);

      await waitFor(() => {
        expect(screen.getByText(/第 1 行的行列索引无效/)).toBeInTheDocument();
      });
    });

    it('应该跳过空行 | should skip empty lines', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const batchInput = screen.getByLabelText(/批量数据/);
      await user.type(batchInput, '1,1,内容1\n\n2,2,内容2\n\n');

      const batchButton = screen.getByRole('button', { name: /批量更新/i });
      await user.click(batchButton);

      await waitFor(() => {
        expect(pptTools.updateTableCellsBatch).toHaveBeenCalledWith(
          {
            cells: [
              { rowIndex: 0, columnIndex: 0, text: '内容1' },
              { rowIndex: 1, columnIndex: 1, text: '内容2' },
            ],
          },
          { tableIndex: 0 }
        );
      });
    });

    it('应该在批量更新失败时显示错误 | should show error when batch update fails', async () => {
      const user = userEvent.setup();

      vi.mocked(pptTools.updateTableCellsBatch).mockResolvedValueOnce({
        success: false,
        cellsUpdated: 0,
        rowCount: 0,
        columnCount: 0,
        error: '批量更新失败',
      });

      renderWithProviders(<TableCellUpdate />);

      const batchInput = screen.getByLabelText(/批量数据/);
      await user.type(batchInput, '1,1,内容');

      const batchButton = screen.getByRole('button', { name: /批量更新/i });
      await user.click(batchButton);

      await waitFor(() => {
        expect(screen.getByText(/批量更新失败/)).toBeInTheDocument();
      });
    });
  });

  describe('表格索引与 shapeId 交互测试 | Table Index and ShapeId Interaction Tests', () => {
    it('应该在输入 shapeId 时禁用表格索引 | should disable table index when shapeId is entered', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const shapeIdInput = screen.getByLabelText('表格形状 ID（可选）');
      const tableIndexInput = screen.getByLabelText('表格索引（默认 0）') as HTMLInputElement;

      await user.type(shapeIdInput, 'table-123');

      expect(tableIndexInput).toBeDisabled();
    });

    it('应该在清空 shapeId 时启用表格索引 | should enable table index when shapeId is cleared', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const shapeIdInput = screen.getByLabelText('表格形状 ID（可选）');
      const tableIndexInput = screen.getByLabelText('表格索引（默认 0）') as HTMLInputElement;

      await user.type(shapeIdInput, 'table-123');
      expect(tableIndexInput).toBeDisabled();

      await user.clear(shapeIdInput);
      expect(tableIndexInput).not.toBeDisabled();
    });
  });

  describe('边界情况测试 | Edge Case Tests', () => {
    it('应该能够处理特殊字符 | should be able to handle special characters', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');
      const contentInput = screen.getByLabelText('单元格内容');

      await user.type(rowInput, '1');
      await user.type(colInput, '1');
      await user.type(contentInput, '特殊字符 !@#$%^&*()');

      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTableCell).toHaveBeenCalledWith(
          { rowIndex: 0, columnIndex: 0, text: '特殊字符 !@#$%^&*()' },
          { tableIndex: 0 }
        );
      });
    });

    it('应该能够处理换行符 | should be able to handle newlines', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');
      const contentInput = screen.getByLabelText('单元格内容');

      await user.type(rowInput, '1');
      await user.type(colInput, '1');
      await user.type(contentInput, '第一行{Enter}第二行');

      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTableCell).toHaveBeenCalled();
      });
    });

    it('应该能够处理大索引值 | should be able to handle large index values', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');
      const contentInput = screen.getByLabelText('单元格内容');

      await user.type(rowInput, '100');
      await user.type(colInput, '50');
      await user.type(contentInput, 'Test');

      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTableCell).toHaveBeenCalledWith(
          { rowIndex: 99, columnIndex: 49, text: 'Test' },
          { tableIndex: 0 }
        );
      });
    });

    it('应该正确转换行列编号（从1开始到从0开始）| should correctly convert row/column numbers (1-based to 0-based)', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');
      const contentInput = screen.getByLabelText('单元格内容');

      // 用户输入第1行第1列（界面显示）
      // User inputs row 1 column 1 (UI display)
      await user.type(rowInput, '1');
      await user.type(colInput, '1');
      await user.type(contentInput, 'Test');

      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      // 应该转换为索引0,0（API调用）
      // Should convert to index 0,0 (API call)
      await waitFor(() => {
        expect(pptTools.updateTableCell).toHaveBeenCalledWith(
          { rowIndex: 0, columnIndex: 0, text: 'Test' },
          { tableIndex: 0 }
        );
      });
    });

    it('应该能够处理长文本内容 | should be able to handle long text content', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');
      const contentInput = screen.getByLabelText('单元格内容');

      const longText = 'A'.repeat(500);

      await user.type(rowInput, '1');
      await user.type(colInput, '1');
      await user.type(contentInput, longText);

      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTableCell).toHaveBeenCalledWith(
          { rowIndex: 0, columnIndex: 0, text: longText },
          { tableIndex: 0 }
        );
      });
    });
  });

  describe('加载状态测试 | Loading State Tests', () => {
    it('应该在操作进行时禁用按钮 | should disable buttons during operations', async () => {
      const user = userEvent.setup();

      // 模拟一个慢速操作
      // Mock a slow operation
      vi.mocked(pptTools.updateTableCell).mockImplementation(
        () =>
          new Promise((resolve) =>
            setTimeout(
              () =>
                resolve({
                  success: true,
                  cellsUpdated: 1,
                  rowCount: 3,
                  columnCount: 3,
                }),
              100
            )
          )
      );

      renderWithProviders(<TableCellUpdate />);

      const rowInput = screen.getByLabelText('行索引');
      const colInput = screen.getByLabelText('列索引');
      const contentInput = screen.getByLabelText('单元格内容');

      await user.type(rowInput, '1');
      await user.type(colInput, '1');
      await user.type(contentInput, 'Test');

      const updateButton = screen.getByRole('button', { name: /更新单元格/i });
      await user.click(updateButton);

      // 等待操作完成
      // Wait for operation to complete
      await waitFor(
        () => {
          expect(pptTools.updateTableCell).toHaveBeenCalled();
        },
        { timeout: 3000 }
      );
    });
  });
});
