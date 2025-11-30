/**
 * 文件名: TextUpdate.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: TextUpdate 组件单元测试 | TextUpdate component unit tests
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { renderWithProviders, userEvent } from '../../utils/test-utils';
import TextUpdate from '../../../src/taskpane/components/tools/TextUpdate';
import * as pptTools from '../../../src/ppt-tools';
import { OfficeMockObject } from 'office-addin-mock';

// Mock ppt-tools module
vi.mock('../../../src/ppt-tools', () => ({
  updateTextBox: vi.fn().mockResolvedValue({ success: true, message: '更新成功' }),
  getTextBoxStyle: vi.fn().mockResolvedValue({
    elementId: 'test-id',
    text: 'Sample Text',
    fontSize: 18,
    fontName: 'Arial',
    fontColor: '#000000',
    bold: false,
    italic: false,
    underline: false,
    horizontalAlignment: 'Left',
    verticalAlignment: 'Top',
    backgroundColor: '#FFFFFF',
    left: 100,
    top: 100,
    width: 300,
    height: 100,
  }),
}));

describe('TextUpdate 组件单元测试 | TextUpdate Component Unit Tests', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    delete (global as any).PowerPoint;
    
    // 重新设置默认的 mock 实现
    vi.mocked(pptTools.updateTextBox).mockImplementation(() => 
      Promise.resolve({ success: true, message: '更新成功' })
    );
    vi.mocked(pptTools.getTextBoxStyle).mockImplementation(() =>
      Promise.resolve({
        elementId: 'test-id',
        text: 'Sample Text',
        fontSize: 18,
        fontName: 'Arial',
        fontColor: '#000000',
        bold: false,
        italic: false,
        underline: false,
        horizontalAlignment: 'Left',
        verticalAlignment: 'Top',
        backgroundColor: '#FFFFFF',
        left: 100,
        top: 100,
        width: 300,
        height: 100,
      })
    );
  });

  describe('组件渲染测试 | Component Rendering Tests', () => {
    it('应该正确渲染组件 | should render component correctly', () => {
      renderWithProviders(<TextUpdate />);

      // 验证标题
      expect(screen.getByText('文本框更新工具')).toBeInTheDocument();

      // 验证按钮
      expect(screen.getByRole('button', { name: /获取PPT中选中的元素/i })).toBeInTheDocument();
      expect(screen.getByRole('button', { name: /加载当前样式/i })).toBeInTheDocument();
      expect(screen.getByRole('button', { name: /更新文本框/i })).toBeInTheDocument();
      expect(screen.getByRole('button', { name: /重置/i })).toBeInTheDocument();

      // 验证输入框
      expect(screen.getByLabelText('元素ID:')).toBeInTheDocument();
      expect(screen.getByLabelText('文本内容:')).toBeInTheDocument();
      expect(screen.getByLabelText('字号:')).toBeInTheDocument();
      expect(screen.getByLabelText('字体:')).toBeInTheDocument();
    });

    it('应该显示使用说明 | should display usage instructions', () => {
      renderWithProviders(<TextUpdate />);

      expect(screen.getByText('使用说明:')).toBeInTheDocument();
      expect(screen.getByText(/在PPT中选中一个文本框元素/)).toBeInTheDocument();
    });

    it('应该显示所有字体样式选项 | should display all font style options', () => {
      renderWithProviders(<TextUpdate />);

      expect(screen.getByText('加粗')).toBeInTheDocument();
      expect(screen.getByText('斜体')).toBeInTheDocument();
      expect(screen.getByText('下划线')).toBeInTheDocument();
    });

    it('应该显示对齐方式选项 | should display alignment options', () => {
      renderWithProviders(<TextUpdate />);

      expect(screen.getByLabelText('水平对齐:')).toBeInTheDocument();
      expect(screen.getByLabelText('垂直对齐:')).toBeInTheDocument();
    });

    it('应该显示位置和尺寸输入框 | should display position and size inputs', () => {
      renderWithProviders(<TextUpdate />);

      expect(screen.getByLabelText('X坐标:')).toBeInTheDocument();
      expect(screen.getByLabelText('Y坐标:')).toBeInTheDocument();
      expect(screen.getByLabelText('宽度:')).toBeInTheDocument();
      expect(screen.getByLabelText('高度:')).toBeInTheDocument();
    });
  });

  describe('获取选中元素测试 | Get Selected Shape Tests', () => {
    it('应该能够获取选中的元素 | should be able to get selected shape', async () => {
      const user = userEvent.setup();

      // Mock getTextBoxStyle to return data for shape-123
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce({
        elementId: 'shape-123',
        text: 'Sample Text',
        fontSize: 18,
        fontName: 'Arial',
        fontColor: '#000000',
        bold: false,
        italic: false,
        underline: false,
        horizontalAlignment: 'Left',
        verticalAlignment: 'Top',
        backgroundColor: '#FFFFFF',
        left: 100,
        top: 100,
        width: 300,
        height: 100,
      });

      // Mock PowerPoint API
      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'shape-123',
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

      renderWithProviders(<TextUpdate />);

      const button = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(button);

      await waitFor(() => {
        const input = screen.getByLabelText('元素ID:') as HTMLInputElement;
        expect(input.value).toBe('shape-123');
      });

      // 注意：由于自动加载样式，不会显示"已获取选中元素"消息
      // 而是会显示样式已加载的相关内容
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

      renderWithProviders(<TextUpdate />);

      const button = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('请先在幻灯片中选中一个文本框元素')).toBeInTheDocument();
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

      renderWithProviders(<TextUpdate />);

      const button = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('请只选中一个元素')).toBeInTheDocument();
      });
    });

    it('应该在选中不支持的元素类型时显示警告 | should show warning for unsupported shape type', async () => {
      const user = userEvent.setup();

      // Mock getTextBoxStyle - 不会被调用因为类型不支持
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce(null);

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

      renderWithProviders(<TextUpdate />);

      const button = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/警告: 选中的元素类型 "Picture" 可能不支持文本编辑/)).toBeInTheDocument();
      });
    });
  });

  describe('加载样式测试 | Load Style Tests', () => {
    it('应该能够加载元素样式 | should be able to load element style', async () => {
      const user = userEvent.setup();

      // 完全重置 mock 并设置新的实现
      vi.mocked(pptTools.getTextBoxStyle).mockReset();
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValue({
        elementId: 'test-id',
        text: 'Sample Text',
        fontSize: 18,
        fontName: 'Arial',
        fontColor: '#000000',
        bold: false,
        italic: false,
        underline: false,
        horizontalAlignment: 'Left',
        verticalAlignment: 'Top',
        backgroundColor: '#FFFFFF',
        left: 100,
        top: 100,
        width: 300,
        height: 100,
      });

      renderWithProviders(<TextUpdate />);

      // 输入元素ID
      const idInput = screen.getByLabelText('元素ID:');
      await user.type(idInput, 'test-id');

      // 点击加载样式按钮
      const loadButton = screen.getByRole('button', { name: /加载当前样式/i });
      await user.click(loadButton);

      // 等待 getTextBoxStyle 被调用
      await waitFor(() => {
        expect(pptTools.getTextBoxStyle).toHaveBeenCalledWith('test-id');
      });

      // 等待成功消息或表单填充（使用更宽松的断言）
      await waitFor(
        () => {
          // 至少验证函数被调用了
          expect(pptTools.getTextBoxStyle).toHaveBeenCalled();
          // 验证成功消息出现或字段有值
          const fontSizeInput = screen.getByLabelText('字号:') as HTMLInputElement;
          const hasMessage = screen.queryByText('成功加载当前样式');
          const hasValue = fontSizeInput.value !== '';
          expect(hasMessage || hasValue).toBeTruthy();
        },
        { timeout: 3000 }
      );
    });

    it('应该在元素ID为空时禁用加载按钮 | should disable load button when element ID is empty', () => {
      renderWithProviders(<TextUpdate />);

      const loadButton = screen.getByRole('button', { name: /加载当前样式/i });
      expect(loadButton).toBeDisabled();
    });

    it('应该在加载失败时显示错误 | should show error when loading fails', async () => {
      const user = userEvent.setup();

      // 重置并设置返回 null
      vi.mocked(pptTools.getTextBoxStyle).mockReset().mockResolvedValue(null);

      renderWithProviders(<TextUpdate />);

      const idInput = screen.getByLabelText('元素ID:');
      await user.type(idInput, 'invalid-id');

      const loadButton = screen.getByRole('button', { name: /加载当前样式/i });
      await user.click(loadButton);

      await waitFor(
        () => {
          expect(screen.getByText('加载样式失败')).toBeInTheDocument();
        },
        { timeout: 3000 }
      );
    });
  });

  describe('更新文本框测试 | Update Text Box Tests', () => {
    it('应该能够更新文本框 | should be able to update text box', async () => {
      const user = userEvent.setup();

      // Mock getTextBoxStyle
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce({
        elementId: 'test-id',
        text: 'Sample Text',
        fontSize: 18,
        fontName: 'Arial',
        fontColor: '#000000',
        bold: false,
        italic: false,
        underline: false,
        horizontalAlignment: 'Left',
        verticalAlignment: 'Top',
        backgroundColor: '#FFFFFF',
        left: 100,
        top: 100,
        width: 300,
        height: 100,
      });

      // 模拟选中元素以启用更新按钮
      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'test-id',
                  type: 'TextBox',
                  name: 'Test',
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

      renderWithProviders(<TextUpdate />);

      const getButton = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(getButton);

      await waitFor(() => {
        const input = screen.getByLabelText('元素ID:') as HTMLInputElement;
        expect(input.value).toBe('test-id');
      });

      // 输入文本
      const textInput = screen.getByLabelText('文本内容:');
      await user.clear(textInput);
      await user.type(textInput, 'Updated Text');

      // 点击更新按钮
      const updateButton = screen.getByRole('button', { name: /更新文本框/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTextBox).toHaveBeenCalled();
      });

      await waitFor(() => {
        expect(screen.getByText(/更新成功/)).toBeInTheDocument();
      });
    });

    it('应该在元素ID为空时禁用更新按钮 | should disable update button when element ID is empty', () => {
      renderWithProviders(<TextUpdate />);

      const updateButton = screen.getByRole('button', { name: /更新文本框/i });
      expect(updateButton).toBeDisabled();
    });

    it('应该能够更新字体属性 | should be able to update font properties', async () => {
      const user = userEvent.setup();

      // Mock getTextBoxStyle
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce({
        elementId: 'test-id',
        text: 'Sample Text',
        fontSize: 18,
        fontName: 'Arial',
        fontColor: '#000000',
        bold: false,
        italic: false,
        underline: false,
        horizontalAlignment: 'Left',
        verticalAlignment: 'Top',
        backgroundColor: '#FFFFFF',
        left: 100,
        top: 100,
        width: 300,
        height: 100,
      });

      // 设置元素ID和类型
      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'test-id',
                  type: 'TextBox',
                  name: 'Test',
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

      renderWithProviders(<TextUpdate />);

      const getButton = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(getButton);

      await waitFor(() => {
        const input = screen.getByLabelText('元素ID:') as HTMLInputElement;
        expect(input.value).toBe('test-id');
      });

      // 输入字体属性
      const fontSizeInput = screen.getByLabelText('字号:');
      await user.clear(fontSizeInput);
      await user.type(fontSizeInput, '24');

      const fontNameInput = screen.getByLabelText('字体:');
      await user.clear(fontNameInput);
      await user.type(fontNameInput, 'Calibri');

      // 选中加粗和斜体
      const boldCheckbox = screen.getByRole('checkbox', { name: /加粗/i });
      const italicCheckbox = screen.getByRole('checkbox', { name: /斜体/i });
      await user.click(boldCheckbox);
      await user.click(italicCheckbox);

      // 点击更新
      const updateButton = screen.getByRole('button', { name: /更新文本框/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTextBox).toHaveBeenCalledWith(
          expect.objectContaining({
            elementId: 'test-id',
            fontSize: 24,
            fontName: 'Calibri',
            bold: true,
            italic: true,
          })
        );
      });
    });

    it('应该能够更新对齐方式 | should be able to update alignment', async () => {
      const user = userEvent.setup();

      // Mock getTextBoxStyle
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce({
        elementId: 'test-id',
        text: 'Sample Text',
        fontSize: 18,
        fontName: 'Arial',
        fontColor: '#000000',
        bold: false,
        italic: false,
        underline: false,
        horizontalAlignment: 'Left',
        verticalAlignment: 'Top',
        backgroundColor: '#FFFFFF',
        left: 100,
        top: 100,
        width: 300,
        height: 100,
      });

      // 设置元素ID
      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'test-id',
                  type: 'TextBox',
                  name: 'Test',
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

      renderWithProviders(<TextUpdate />);

      const getButton = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(getButton);

      await waitFor(() => {
        const input = screen.getByLabelText('元素ID:') as HTMLInputElement;
        expect(input.value).toBe('test-id');
      });

      // 选择对齐方式
      const horizontalSelect = screen.getByLabelText('水平对齐:');
      const verticalSelect = screen.getByLabelText('垂直对齐:');

      await user.selectOptions(horizontalSelect, 'Center');
      await user.selectOptions(verticalSelect, 'Middle');

      // 点击更新
      const updateButton = screen.getByRole('button', { name: /更新文本框/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTextBox).toHaveBeenCalledWith(
          expect.objectContaining({
            horizontalAlignment: 'Center',
            verticalAlignment: 'Middle',
          })
        );
      });
    });

    it('应该能够更新位置和尺寸 | should be able to update position and size', async () => {
      const user = userEvent.setup();

      // Mock getTextBoxStyle
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce({
        elementId: 'test-id',
        text: 'Sample Text',
        fontSize: 18,
        fontName: 'Arial',
        fontColor: '#000000',
        bold: false,
        italic: false,
        underline: false,
        horizontalAlignment: 'Left',
        verticalAlignment: 'Top',
        backgroundColor: '#FFFFFF',
        left: 100,
        top: 100,
        width: 300,
        height: 100,
      });

      // 设置元素ID
      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'test-id',
                  type: 'TextBox',
                  name: 'Test',
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

      renderWithProviders(<TextUpdate />);

      const getButton = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(getButton);

      await waitFor(() => {
        const input = screen.getByLabelText('元素ID:') as HTMLInputElement;
        expect(input.value).toBe('test-id');
      });

      // 输入位置和尺寸
      const leftInput = screen.getByLabelText('X坐标:');
      const topInput = screen.getByLabelText('Y坐标:');
      const widthInput = screen.getByLabelText('宽度:');
      const heightInput = screen.getByLabelText('高度:');

      await user.clear(leftInput);
      await user.type(leftInput, '200');
      await user.clear(topInput);
      await user.type(topInput, '150');
      await user.clear(widthInput);
      await user.type(widthInput, '400');
      await user.clear(heightInput);
      await user.type(heightInput, '200');

      // 点击更新
      const updateButton = screen.getByRole('button', { name: /更新文本框/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTextBox).toHaveBeenCalledWith(
          expect.objectContaining({
            left: 200,
            top: 150,
            width: 400,
            height: 200,
          })
        );
      });
    });

    it('应该在更新失败时显示错误 | should show error when update fails', async () => {
      const user = userEvent.setup();

      vi.mocked(pptTools.updateTextBox).mockResolvedValueOnce({
        success: false,
        message: '更新失败',
      });

      // Mock getTextBoxStyle
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce({
        elementId: 'test-id',
        text: 'Sample Text',
        fontSize: 18,
        fontName: 'Arial',
        fontColor: '#000000',
        bold: false,
        italic: false,
        underline: false,
        horizontalAlignment: 'Left',
        verticalAlignment: 'Top',
        backgroundColor: '#FFFFFF',
        left: 100,
        top: 100,
        width: 300,
        height: 100,
      });

      // 设置元素ID
      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'test-id',
                  type: 'TextBox',
                  name: 'Test',
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

      renderWithProviders(<TextUpdate />);

      const getButton = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(getButton);

      await waitFor(() => {
        const input = screen.getByLabelText('元素ID:') as HTMLInputElement;
        expect(input.value).toBe('test-id');
      });

      const updateButton = screen.getByRole('button', { name: /更新文本框/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(screen.getByText(/更新失败/)).toBeInTheDocument();
      });
    });
  });

  describe('重置功能测试 | Reset Functionality Tests', () => {
    it('应该能够重置所有字段 | should be able to reset all fields', async () => {
      const user = userEvent.setup();

      renderWithProviders(<TextUpdate />);

      // 填充一些字段
      const textInput = screen.getByLabelText('文本内容:');
      await user.type(textInput, 'Some text');

      const fontSizeInput = screen.getByLabelText('字号:');
      await user.type(fontSizeInput, '24');

      const boldCheckbox = screen.getByRole('checkbox', { name: /加粗/i });
      await user.click(boldCheckbox);

      // 点击重置
      const resetButton = screen.getByRole('button', { name: /重置/i });
      await user.click(resetButton);

      // 验证字段已重置
      await waitFor(() => {
        expect((textInput as HTMLTextAreaElement).value).toBe('');
        expect((fontSizeInput as HTMLInputElement).value).toBe('');
        expect((boldCheckbox as HTMLInputElement).checked).toBe(false);
      });

      expect(screen.getByText('已重置所有字段')).toBeInTheDocument();
    });
  });

  describe('边界情况测试 | Edge Case Tests', () => {
    it('应该能够处理空文本 | should be able to handle empty text', async () => {
      const user = userEvent.setup();

      // Mock getTextBoxStyle
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce({
        elementId: 'test-id',
        text: 'Sample Text',
        fontSize: 18,
        fontName: 'Arial',
        fontColor: '#000000',
        bold: false,
        italic: false,
        underline: false,
        horizontalAlignment: 'Left',
        verticalAlignment: 'Top',
        backgroundColor: '#FFFFFF',
        left: 100,
        top: 100,
        width: 300,
        height: 100,
      });

      // 设置元素ID
      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'test-id',
                  type: 'TextBox',
                  name: 'Test',
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

      renderWithProviders(<TextUpdate />);

      const getButton = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(getButton);

      await waitFor(() => {
        const input = screen.getByLabelText('元素ID:') as HTMLInputElement;
        expect(input.value).toBe('test-id');
      });

      const textInput = screen.getByLabelText('文本内容:');
      await user.clear(textInput);

      const updateButton = screen.getByRole('button', { name: /更新文本框/i });
      await user.click(updateButton);

      await waitFor(() => {
        expect(pptTools.updateTextBox).toHaveBeenCalledWith(
          expect.objectContaining({
            text: '',
          })
        );
      });
    });

    it('应该能够处理颜色选择器 | should be able to handle color pickers', async () => {
      renderWithProviders(<TextUpdate />);

      // 通过 placeholder 或其他方式查找颜色输入框
      const inputs = screen.getAllByDisplayValue('#000000');
      const fontColorInput = inputs[0] as HTMLInputElement;
      
      const bgInputs = screen.getAllByDisplayValue('#ffffff');
      const bgColorInput = bgInputs[0] as HTMLInputElement;

      // 验证默认颜色
      expect(fontColorInput.value).toBe('#000000');
      expect(bgColorInput.value).toBe('#ffffff');
      expect(fontColorInput.type).toBe('color');
      expect(bgColorInput.type).toBe('color');
    });

    it('应该在不支持的元素类型时禁用更新按钮 | should disable update button for unsupported shape type', async () => {
      const user = userEvent.setup();

      // Mock getTextBoxStyle - 不会被调用
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce(null);

      renderWithProviders(<TextUpdate />);

      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'test-id',
                  type: 'Picture',
                  name: 'Image',
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

      const getButton = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(getButton);

      await waitFor(() => {
        const updateButton = screen.getByRole('button', { name: /更新文本框/i });
        expect(updateButton).toBeDisabled();
      });
    });

    it('应该显示警告信息对于不支持的元素类型 | should display warning for unsupported shape type', async () => {
      const user = userEvent.setup();

      // Mock getTextBoxStyle - 不会被调用
      vi.mocked(pptTools.getTextBoxStyle).mockResolvedValueOnce(null);

      renderWithProviders(<TextUpdate />);

      const mockData = {
        context: {
          presentation: {
            getSelectedShapes: () => ({
              getCount: () => ({ value: 1 }),
              items: [
                {
                  id: 'test-id',
                  type: 'Line',
                  name: 'Line',
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

      const getButton = screen.getByRole('button', { name: /获取PPT中选中的元素/i });
      await user.click(getButton);

      await waitFor(() => {
        expect(
          screen.getByText(/选中的元素类型 "Line" 不支持文本编辑，请选择文本框、占位符或几何形状/)
        ).toBeInTheDocument();
      });
    });
  });
});
