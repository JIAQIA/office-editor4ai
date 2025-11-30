/**
 * 文件名: TableInsertion.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: TableInsertion 组件单元测试 | TableInsertion component unit tests
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { renderWithProviders, userEvent } from '../../utils/test-utils';
import TableInsertion from '../../../src/taskpane/components/tools/TableInsertion';
import * as pptTools from '../../../src/ppt-tools';

// Mock ppt-tools module
vi.mock('../../../src/ppt-tools', async () => {
  const actual = await vi.importActual('../../../src/ppt-tools');
  return {
    ...actual,
    insertTableToSlide: vi.fn().mockResolvedValue({
      shapeId: 'mock-table-id',
      rowCount: 3,
      columnCount: 3,
      width: 400,
      height: 90,
      left: 160,
      top: 225,
    }),
  };
});

describe('TableInsertion 组件单元测试 | TableInsertion Component Unit Tests', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('应该正确渲染组件 | should render component correctly', () => {
    const { container } = renderWithProviders(<TableInsertion />);

    // 验证模板选择器存在 | Verify template selector exists
    expect(screen.getByText('选择表格模板')).toBeInTheDocument();
    
    // 验证表格尺寸输入框存在 | Verify table size inputs exist
    expect(screen.getByLabelText('行数')).toBeInTheDocument();
    expect(screen.getByLabelText('列数')).toBeInTheDocument();
    
    // 验证位置和尺寸输入框存在 | Verify position and size inputs exist
    expect(screen.getByLabelText('X 坐标')).toBeInTheDocument();
    expect(screen.getByLabelText('Y 坐标')).toBeInTheDocument();
    expect(screen.getByLabelText('宽度（磅）')).toBeInTheDocument();
    expect(screen.getByLabelText('高度（磅）')).toBeInTheDocument();
    
    // 验证样式设置存在 | Verify style settings exist
    const colorInputs = container.querySelectorAll('input[type="color"]');
    expect(colorInputs.length).toBe(2);
    expect(screen.getByText('显示表头样式')).toBeInTheDocument();
    
    // 验证插入按钮存在 | Verify insert button exists
    expect(screen.getByRole('button', { name: '确认插入' })).toBeInTheDocument();
  });

  it('应该显示默认值 | should display default values', () => {
    renderWithProviders(<TableInsertion />);

    const rowInput = screen.getByLabelText('行数') as HTMLInputElement;
    const columnInput = screen.getByLabelText('列数') as HTMLInputElement;
    const widthInput = screen.getByLabelText('宽度（磅）') as HTMLInputElement;

    expect(rowInput.value).toBe('3');
    expect(columnInput.value).toBe('3');
    expect(widthInput.value).toBe('400');
  });

  it('应该能够修改行列数 | should be able to modify row and column count', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const rowInput = screen.getByLabelText('行数') as HTMLInputElement;
    const columnInput = screen.getByLabelText('列数') as HTMLInputElement;

    // 清空并输入新值 | Clear and input new values
    await user.clear(rowInput);
    await user.type(rowInput, '5');
    await user.clear(columnInput);
    await user.type(columnInput, '4');

    expect(rowInput.value).toBe('5');
    expect(columnInput.value).toBe('4');
  });

  it('应该能够修改位置坐标 | should be able to modify position coordinates', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const leftInput = screen.getByLabelText('X 坐标') as HTMLInputElement;
    const topInput = screen.getByLabelText('Y 坐标') as HTMLInputElement;

    // 输入坐标 | Input coordinates
    await user.type(leftInput, '100');
    await user.type(topInput, '200');

    expect(leftInput.value).toBe('100');
    expect(topInput.value).toBe('200');
  });

  it('应该能够修改尺寸 | should be able to modify size', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const widthInput = screen.getByLabelText('宽度（磅）') as HTMLInputElement;
    const heightInput = screen.getByLabelText('高度（磅）') as HTMLInputElement;

    // 清空并输入新值 | Clear and input new values
    await user.clear(widthInput);
    await user.type(widthInput, '500');
    await user.type(heightInput, '200');

    expect(widthInput.value).toBe('500');
    expect(heightInput.value).toBe('200');
  });

  it('点击插入按钮应该调用 insertTableToSlide 函数（默认参数）| clicking insert button should call insertTableToSlide (default params)', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertTableToSlide).toHaveBeenCalledTimes(1);
      expect(pptTools.insertTableToSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          rowCount: 3,
          columnCount: 3,
          width: 400,
          showHeader: true,
          headerColor: '#4472C4',
          borderColor: '#D0D0D0',
        })
      );
    });
  });

  it('点击插入按钮应该调用 insertTableToSlide 函数（带位置）| clicking insert button should call insertTableToSlide (with position)', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const leftInput = screen.getByLabelText('X 坐标');
    const topInput = screen.getByLabelText('Y 坐标');
    const button = screen.getByRole('button', { name: '确认插入' });

    // 输入坐标 | Input coordinates
    await user.type(leftInput, '150');
    await user.type(topInput, '250');
    
    // 点击插入 | Click insert
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertTableToSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          left: 150,
          top: 250,
        })
      );
    });
  });

  it('点击插入按钮应该调用 insertTableToSlide 函数（完整配置）| clicking insert button should call insertTableToSlide (full config)', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const rowInput = screen.getByLabelText('行数');
    const columnInput = screen.getByLabelText('列数');
    const leftInput = screen.getByLabelText('X 坐标');
    const topInput = screen.getByLabelText('Y 坐标');
    const widthInput = screen.getByLabelText('宽度（磅）');
    const heightInput = screen.getByLabelText('高度（磅）');
    const button = screen.getByRole('button', { name: '确认插入' });

    // 输入所有参数 | Input all parameters
    await user.clear(rowInput);
    await user.type(rowInput, '5');
    await user.clear(columnInput);
    await user.type(columnInput, '4');
    await user.type(leftInput, '100');
    await user.type(topInput, '200');
    await user.clear(widthInput);
    await user.type(widthInput, '500');
    await user.type(heightInput, '250');
    
    // 点击插入 | Click insert
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertTableToSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          rowCount: 5,
          columnCount: 4,
          left: 100,
          top: 200,
          width: 500,
          height: 250,
        })
      );
    });
  });

  it('应该显示成功消息 | should display success message', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('插入成功', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText(/表格已插入！/)).toBeInTheDocument();
    });
  });

  it('应该在插入失败时显示错误消息 | should display error message on insertion failure', async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.insertTableToSlide).mockRejectedValueOnce(new Error('插入失败'));
    
    renderWithProviders(<TableInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('插入失败', { selector: 'span' })).toBeInTheDocument();
    });
  });

  it('应该在行数为 0 时显示警告消息 | should display warning message when row count is 0', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const rowInput = screen.getByLabelText('行数');
    const button = screen.getByRole('button', { name: '确认插入' });

    await user.clear(rowInput);
    await user.type(rowInput, '0');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('参数错误', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText('行数必须是大于 0 的整数')).toBeInTheDocument();
    });

    expect(pptTools.insertTableToSlide).not.toHaveBeenCalled();
  });

  it('应该在列数为负数时显示警告消息 | should display warning message when column count is negative', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const columnInput = screen.getByLabelText('列数');
    const button = screen.getByRole('button', { name: '确认插入' });

    await user.clear(columnInput);
    await user.type(columnInput, '-1');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('参数错误', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText('列数必须是大于 0 的整数')).toBeInTheDocument();
    });

    expect(pptTools.insertTableToSlide).not.toHaveBeenCalled();
  });

  it('应该在行数超过 100 时显示警告消息 | should display warning message when row count exceeds 100', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const rowInput = screen.getByLabelText('行数');
    const button = screen.getByRole('button', { name: '确认插入' });

    await user.clear(rowInput);
    await user.type(rowInput, '101');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('参数错误', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText('行数不能超过 100')).toBeInTheDocument();
    });

    expect(pptTools.insertTableToSlide).not.toHaveBeenCalled();
  });

  it('应该在列数超过 50 时显示警告消息 | should display warning message when column count exceeds 50', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const columnInput = screen.getByLabelText('列数');
    const button = screen.getByRole('button', { name: '确认插入' });

    await user.clear(columnInput);
    await user.type(columnInput, '51');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('参数错误', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText('列数不能超过 50')).toBeInTheDocument();
    });

    expect(pptTools.insertTableToSlide).not.toHaveBeenCalled();
  });

  it('应该在宽度为 0 时显示警告消息 | should display warning message when width is 0', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const widthInput = screen.getByLabelText('宽度（磅）');
    const button = screen.getByRole('button', { name: '确认插入' });

    await user.clear(widthInput);
    await user.type(widthInput, '0');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('参数错误', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText('宽度必须大于 0')).toBeInTheDocument();
    });

    expect(pptTools.insertTableToSlide).not.toHaveBeenCalled();
  });

  it('应该显示位置提示信息 | should display position hint information', () => {
    renderWithProviders(<TableInsertion />);

    expect(screen.getByText(/位置范围提示/)).toBeInTheDocument();
    expect(screen.getByText(/标准 16:9 幻灯片尺寸约为 720×540 磅/)).toBeInTheDocument();
  });

  it('应该在插入时禁用按钮 | should disable button during insertion', async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.insertTableToSlide).mockImplementation(
      () => new Promise(resolve => setTimeout(() => resolve({
        shapeId: 'mock-id',
        rowCount: 3,
        columnCount: 3,
        width: 400,
        height: 90,
        left: 160,
        top: 225,
      }), 100))
    );

    renderWithProviders(<TableInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    // 按钮应该显示"插入中..." | Button should show "插入中..."
    await waitFor(() => {
      expect(screen.getByRole('button', { name: '插入中...' })).toBeDisabled();
    });

    await waitFor(() => {
      expect(screen.getByRole('button', { name: '确认插入' })).not.toBeDisabled();
    }, { timeout: 3000 });
  });

  it('应该能够切换表头样式开关 | should be able to toggle header style switch', async () => {
    const user = userEvent.setup();
    const { container } = renderWithProviders(<TableInsertion />);

    // 查找表头样式开关
    const switchElement = container.querySelector('input[type="checkbox"]') as HTMLInputElement;
    expect(switchElement).toBeInTheDocument();
    expect(switchElement.checked).toBe(true);

    // 切换开关 | Toggle switch
    await user.click(switchElement);
    expect(switchElement.checked).toBe(false);

    await user.click(switchElement);
    expect(switchElement.checked).toBe(true);
  });

  it('应该能够修改颜色 | should be able to modify colors', async () => {
    const user = userEvent.setup();
    const { container } = renderWithProviders(<TableInsertion />);

    // 使用容器查询获取原生color input
    const colorInputs = container.querySelectorAll('input[type="color"]');
    const headerColorInput = colorInputs[0] as HTMLInputElement;
    const borderColorInput = colorInputs[1] as HTMLInputElement;

    // 验证默认颜色 | Verify default colors
    expect(headerColorInput.value).toBe('#4472c4');
    expect(borderColorInput.value).toBe('#d0d0d0');

    // 修改颜色 | Modify colors
    await user.click(headerColorInput);
    headerColorInput.value = '#ff0000';
    headerColorInput.dispatchEvent(new Event('change', { bubbles: true }));
    
    await user.click(borderColorInput);
    borderColorInput.value = '#00ff00';
    borderColorInput.dispatchEvent(new Event('change', { bubbles: true }));

    expect(headerColorInput.value).toBe('#ff0000');
    expect(borderColorInput.value).toBe('#00ff00');
  });

  it('应该能够选择表格模板 | should be able to select table template', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    // 查找并点击下拉框
    const dropdown = screen.getByRole('combobox');
    await user.click(dropdown);

    // 等待选项出现并选择
    await waitFor(() => {
      const option = screen.getByRole('option', { name: /简单表格/ });
      expect(option).toBeInTheDocument();
    });

    const simpleTableOption = screen.getByRole('option', { name: /简单表格/ });
    await user.click(simpleTableOption);

    // 验证行列数已更新
    await waitFor(() => {
      const rowInput = screen.getByLabelText('行数') as HTMLInputElement;
      const columnInput = screen.getByLabelText('列数') as HTMLInputElement;
      expect(rowInput.value).toBe('2');
      expect(columnInput.value).toBe('3');
    });
  });

  it('应该能够切换数据输入开关 | should be able to toggle data input switch', async () => {
    const user = userEvent.setup();
    const { container } = renderWithProviders(<TableInsertion />);

    // 查找数据输入开关（第二个checkbox）
    const switches = container.querySelectorAll('input[type="checkbox"]');
    const dataSwitch = switches[1] as HTMLInputElement;
    
    expect(dataSwitch.checked).toBe(false);

    // 切换开关 | Toggle switch
    await user.click(dataSwitch);
    
    await waitFor(() => {
      expect(dataSwitch.checked).toBe(true);
      // 验证文本框出现
      expect(screen.getByLabelText(/表格数据/)).toBeInTheDocument();
    });
  });

  it('应该能够输入表格数据 | should be able to input table data', async () => {
    const user = userEvent.setup();
    const { container } = renderWithProviders(<TableInsertion />);

    // 打开数据输入
    const switches = container.querySelectorAll('input[type="checkbox"]');
    const dataSwitch = switches[1] as HTMLInputElement;
    await user.click(dataSwitch);

    // 输入数据
    await waitFor(() => {
      const textarea = screen.getByLabelText(/表格数据/) as HTMLTextAreaElement;
      expect(textarea).toBeInTheDocument();
    });

    const textarea = screen.getByLabelText(/表格数据/) as HTMLTextAreaElement;
    await user.type(textarea, '姓名,年龄\n张三,25');

    expect(textarea.value).toBe('姓名,年龄\n张三,25');
  });

  it('应该在数据维度不匹配时显示警告 | should display warning when data dimensions mismatch', async () => {
    const user = userEvent.setup();
    const { container } = renderWithProviders(<TableInsertion />);

    // 设置行列数
    const rowInput = screen.getByLabelText('行数');
    const columnInput = screen.getByLabelText('列数');
    await user.clear(rowInput);
    await user.type(rowInput, '2');
    await user.clear(columnInput);
    await user.type(columnInput, '2');

    // 打开数据输入
    const switches = container.querySelectorAll('input[type="checkbox"]');
    const dataSwitch = switches[1] as HTMLInputElement;
    await user.click(dataSwitch);

    // 输入不匹配的数据（3行而不是2行）
    await waitFor(() => {
      const textarea = screen.getByLabelText(/表格数据/) as HTMLTextAreaElement;
      expect(textarea).toBeInTheDocument();
    });

    const textarea = screen.getByLabelText(/表格数据/) as HTMLTextAreaElement;
    await user.type(textarea, 'A,B\nC,D\nE,F');

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('数据维度不匹配', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText(/数据有 3 行，但指定了 2 行/)).toBeInTheDocument();
    });

    expect(pptTools.insertTableToSlide).not.toHaveBeenCalled();
  });

  it('应该正确解析浮点数坐标 | should correctly parse floating point coordinates', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const leftInput = screen.getByLabelText('X 坐标');
    const topInput = screen.getByLabelText('Y 坐标');
    const button = screen.getByRole('button', { name: '确认插入' });

    // 输入浮点数坐标 | Input floating point coordinates
    await user.type(leftInput, '123.45');
    await user.type(topInput, '678.90');
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertTableToSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          left: 123.45,
          top: 678.90,
        })
      );
    });
  });

  it('应该在位置为空时不传递位置参数 | should not pass position parameters when empty', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      const callArgs = vi.mocked(pptTools.insertTableToSlide).mock.calls[0][0];
      expect(callArgs.left).toBeUndefined();
      expect(callArgs.top).toBeUndefined();
    });
  });

  it('应该在高度为空时不传递高度参数 | should not pass height parameter when empty', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      const callArgs = vi.mocked(pptTools.insertTableToSlide).mock.calls[0][0];
      expect(callArgs.height).toBeUndefined();
    });
  });

  it('应该显示模板预览信息 | should display template preview information', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    // 选择一个模板
    const dropdown = screen.getByRole('combobox');
    await user.click(dropdown);

    await waitFor(() => {
      const option = screen.getByRole('option', { name: /方形表格/ });
      expect(option).toBeInTheDocument();
    });

    const squareTableOption = screen.getByRole('option', { name: /方形表格/ });
    await user.click(squareTableOption);

    // 验证预览信息显示
    await waitFor(() => {
      expect(screen.getByText('方形表格 (3行3列)', { selector: 'strong' })).toBeInTheDocument();
      expect(screen.getByText('适合对比数据')).toBeInTheDocument();
    });
  });

  it('应该在选择自定义模板时不显示预览 | should not display preview when custom template is selected', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TableInsertion />);

    // 先选择一个模板
    const dropdown = screen.getByRole('combobox');
    await user.click(dropdown);

    await waitFor(() => {
      const option = screen.getByRole('option', { name: /简单表格/ });
      expect(option).toBeInTheDocument();
    });

    const simpleTableOption = screen.getByRole('option', { name: /简单表格/ });
    await user.click(simpleTableOption);

    // 再切换回自定义
    await user.click(dropdown);
    
    await waitFor(() => {
      const customOption = screen.getByRole('option', { name: '自定义' });
      expect(customOption).toBeInTheDocument();
    });

    const customOption = screen.getByRole('option', { name: '自定义' });
    await user.click(customOption);

    // 验证预览信息不显示
    await waitFor(() => {
      expect(screen.queryByText(/适合/)).not.toBeInTheDocument();
    });
  });
});
