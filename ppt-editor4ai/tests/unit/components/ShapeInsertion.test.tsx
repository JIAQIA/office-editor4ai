/**
 * 文件名: ShapeInsertion.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: ShapeInsertion 组件单元测试 | ShapeInsertion component unit tests
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { renderWithProviders, userEvent } from '../../utils/test-utils';
import ShapeInsertion from '../../../src/taskpane/components/tools/ShapeInsertion';
import * as pptTools from '../../../src/ppt-tools';

// Mock ppt-tools module
vi.mock('../../../src/ppt-tools', async () => {
  const actual = await vi.importActual('../../../src/ppt-tools');
  return {
    ...actual,
    insertShapeToSlide: vi.fn().mockResolvedValue({
      shapeId: 'mock-shape-id',
      shapeType: 'GeometricShape',
      width: 100,
      height: 100,
      left: 100,
      top: 100,
    }),
  };
});

describe('ShapeInsertion 组件单元测试 | ShapeInsertion Component Unit Tests', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('应该正确渲染组件 | should render component correctly', () => {
    const { container } = renderWithProviders(<ShapeInsertion />);

    // 验证形状类型选择器存在 | Verify shape type selector exists
    expect(screen.getByText('选择形状类型')).toBeInTheDocument();
    
    // 验证位置和尺寸输入框存在 | Verify position and size inputs exist
    expect(screen.getByLabelText('X 坐标')).toBeInTheDocument();
    expect(screen.getByLabelText('Y 坐标')).toBeInTheDocument();
    expect(screen.getByLabelText('宽度')).toBeInTheDocument();
    expect(screen.getByLabelText('高度')).toBeInTheDocument();
    
    // 验证样式设置存在 | Verify style settings exist - 使用容器查询原生color input
    const colorInputs = container.querySelectorAll('input[type="color"]');
    expect(colorInputs.length).toBe(2);
    expect(screen.getByLabelText('边框粗细（磅）')).toBeInTheDocument();
    
    // 验证文本输入框存在 | Verify text input exists
    expect(screen.getByLabelText('形状内文本（可选）')).toBeInTheDocument();
    
    // 验证插入按钮存在 | Verify insert button exists
    expect(screen.getByRole('button', { name: '确认插入' })).toBeInTheDocument();
  });

  it('应该显示默认值 | should display default values', () => {
    renderWithProviders(<ShapeInsertion />);

    const widthInput = screen.getByLabelText('宽度') as HTMLInputElement;
    const heightInput = screen.getByLabelText('高度') as HTMLInputElement;
    const lineWeightInput = screen.getByLabelText('边框粗细（磅）') as HTMLInputElement;

    expect(widthInput.value).toBe('100');
    expect(heightInput.value).toBe('100');
    expect(lineWeightInput.value).toBe('2');
  });

  it('应该能够修改位置坐标 | should be able to modify position coordinates', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const leftInput = screen.getByLabelText('X 坐标') as HTMLInputElement;
    const topInput = screen.getByLabelText('Y 坐标') as HTMLInputElement;

    // 输入坐标 | Input coordinates
    await user.type(leftInput, '200');
    await user.type(topInput, '300');

    expect(leftInput.value).toBe('200');
    expect(topInput.value).toBe('300');
  });

  it('应该能够修改尺寸 | should be able to modify size', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const widthInput = screen.getByLabelText('宽度') as HTMLInputElement;
    const heightInput = screen.getByLabelText('高度') as HTMLInputElement;

    // 清空并输入新值 | Clear and input new values
    await user.clear(widthInput);
    await user.type(widthInput, '250');
    await user.clear(heightInput);
    await user.type(heightInput, '180');

    expect(widthInput.value).toBe('250');
    expect(heightInput.value).toBe('180');
  });

  it('应该能够修改边框粗细 | should be able to modify line weight', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const lineWeightInput = screen.getByLabelText('边框粗细（磅）') as HTMLInputElement;

    await user.clear(lineWeightInput);
    await user.type(lineWeightInput, '5');

    expect(lineWeightInput.value).toBe('5');
  });

  it('应该能够输入文本 | should be able to input text', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const textInput = screen.getByLabelText('形状内文本（可选）') as HTMLInputElement;

    await user.type(textInput, '测试文本');

    expect(textInput.value).toBe('测试文本');
  });

  it('点击插入按钮应该调用 insertShapeToSlide 函数（默认参数）| clicking insert button should call insertShapeToSlide (default params)', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertShapeToSlide).toHaveBeenCalledTimes(1);
      expect(pptTools.insertShapeToSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          shapeType: 'rectangle',
          width: 100,
          height: 100,
          fillColor: '#4472C4',
          lineColor: '#2E5090',
          lineWeight: 2,
        })
      );
    });
  });

  it('点击插入按钮应该调用 insertShapeToSlide 函数（带位置）| clicking insert button should call insertShapeToSlide (with position)', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const leftInput = screen.getByLabelText('X 坐标');
    const topInput = screen.getByLabelText('Y 坐标');
    const button = screen.getByRole('button', { name: '确认插入' });

    // 输入坐标 | Input coordinates
    await user.type(leftInput, '150');
    await user.type(topInput, '250');
    
    // 点击插入 | Click insert
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertShapeToSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          left: 150,
          top: 250,
        })
      );
    });
  });

  it('点击插入按钮应该调用 insertShapeToSlide 函数（带文本）| clicking insert button should call insertShapeToSlide (with text)', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const textInput = screen.getByLabelText('形状内文本（可选）');
    const button = screen.getByRole('button', { name: '确认插入' });

    // 输入文本 | Input text
    await user.type(textInput, '测试文本');
    
    // 点击插入 | Click insert
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertShapeToSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          text: '测试文本',
        })
      );
    });
  });

  it('点击插入按钮应该调用 insertShapeToSlide 函数（完整配置）| clicking insert button should call insertShapeToSlide (full config)', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const leftInput = screen.getByLabelText('X 坐标');
    const topInput = screen.getByLabelText('Y 坐标');
    const widthInput = screen.getByLabelText('宽度');
    const heightInput = screen.getByLabelText('高度');
    const lineWeightInput = screen.getByLabelText('边框粗细（磅）');
    const textInput = screen.getByLabelText('形状内文本（可选）');
    const button = screen.getByRole('button', { name: '确认插入' });

    // 输入所有参数 | Input all parameters
    await user.type(leftInput, '100');
    await user.type(topInput, '200');
    await user.clear(widthInput);
    await user.type(widthInput, '300');
    await user.clear(heightInput);
    await user.type(heightInput, '250');
    await user.clear(lineWeightInput);
    await user.type(lineWeightInput, '5');
    await user.type(textInput, '完整配置');
    
    // 点击插入 | Click insert
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertShapeToSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          shapeType: 'rectangle',
          left: 100,
          top: 200,
          width: 300,
          height: 250,
          lineWeight: 5,
          text: '完整配置',
        })
      );
    });
  });

  it('应该显示成功消息 | should display success message', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('插入成功', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText(/形状已插入！位置:/)).toBeInTheDocument();
    });
  });

  it('应该在插入失败时显示错误消息 | should display error message on insertion failure', async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.insertShapeToSlide).mockRejectedValueOnce(new Error('插入失败'));
    
    renderWithProviders(<ShapeInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('插入失败', { selector: 'span' })).toBeInTheDocument();
    });
  });

  it('应该在宽度为 0 时显示警告消息 | should display warning message when width is 0', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const widthInput = screen.getByLabelText('宽度');
    const button = screen.getByRole('button', { name: '确认插入' });

    await user.clear(widthInput);
    await user.type(widthInput, '0');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('参数错误', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText('宽度和高度必须大于 0')).toBeInTheDocument();
    });

    expect(pptTools.insertShapeToSlide).not.toHaveBeenCalled();
  });

  it('应该在高度为负数时显示警告消息 | should display warning message when height is negative', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const heightInput = screen.getByLabelText('高度');
    const button = screen.getByRole('button', { name: '确认插入' });

    await user.clear(heightInput);
    await user.type(heightInput, '-10');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('参数错误', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText('宽度和高度必须大于 0')).toBeInTheDocument();
    });

    expect(pptTools.insertShapeToSlide).not.toHaveBeenCalled();
  });

  it('应该在边框粗细为负数时显示警告消息 | should display warning message when line weight is negative', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const lineWeightInput = screen.getByLabelText('边框粗细（磅）');
    const button = screen.getByRole('button', { name: '确认插入' });

    await user.clear(lineWeightInput);
    await user.type(lineWeightInput, '-5');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('参数错误', { selector: 'span' })).toBeInTheDocument();
      expect(screen.getByText('边框粗细不能为负数')).toBeInTheDocument();
    });

    expect(pptTools.insertShapeToSlide).not.toHaveBeenCalled();
  });

  it('应该显示位置提示信息 | should display position hint information', () => {
    renderWithProviders(<ShapeInsertion />);

    expect(screen.getByText(/位置范围提示/)).toBeInTheDocument();
    expect(screen.getByText(/标准 16:9 幻灯片尺寸约为 720×540 磅/)).toBeInTheDocument();
  });

  it('应该在插入时禁用按钮 | should disable button during insertion', async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.insertShapeToSlide).mockImplementation(
      () => new Promise(resolve => setTimeout(() => resolve({
        shapeId: 'mock-id',
        shapeType: 'GeometricShape',
        width: 100,
        height: 100,
        left: 100,
        top: 100,
      }), 100))
    );

    renderWithProviders(<ShapeInsertion />);

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

  it('应该正确解析浮点数坐标 | should correctly parse floating point coordinates', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const leftInput = screen.getByLabelText('X 坐标');
    const topInput = screen.getByLabelText('Y 坐标');
    const button = screen.getByRole('button', { name: '确认插入' });

    // 输入浮点数坐标 | Input floating point coordinates
    await user.type(leftInput, '123.45');
    await user.type(topInput, '678.90');
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertShapeToSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          left: 123.45,
          top: 678.90,
        })
      );
    });
  });

  it('应该在位置为空时不传递位置参数 | should not pass position parameters when empty', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      const callArgs = vi.mocked(pptTools.insertShapeToSlide).mock.calls[0][0];
      expect(callArgs.left).toBeUndefined();
      expect(callArgs.top).toBeUndefined();
    });
  });

  it('应该在文本为空时不传递文本参数 | should not pass text parameter when empty', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ShapeInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      const callArgs = vi.mocked(pptTools.insertShapeToSlide).mock.calls[0][0];
      expect(callArgs.text).toBeUndefined();
    });
  });

  it('应该显示形状预览信息 | should display shape preview information', () => {
    renderWithProviders(<ShapeInsertion />);

    // 默认应该显示矩形 | Should display rectangle by default
    // 使用更精确的查询避免匹配到下拉框中的文本
    const shapePreview = screen.getByText('矩形', { selector: 'strong' });
    expect(shapePreview).toBeInTheDocument();
    expect(screen.getByText('基础形状')).toBeInTheDocument();
  });

  it('应该能够修改颜色 | should be able to modify colors', async () => {
    const user = userEvent.setup();
    const { container } = renderWithProviders(<ShapeInsertion />);

    // 使用容器查询获取原生color input
    const colorInputs = container.querySelectorAll('input[type="color"]');
    const fillColorInput = colorInputs[0] as HTMLInputElement;
    const lineColorInput = colorInputs[1] as HTMLInputElement;

    // 验证默认颜色 | Verify default colors
    expect(fillColorInput.value).toBe('#4472c4');
    expect(lineColorInput.value).toBe('#2e5090');

    // 修改颜色 | Modify colors - 使用fireEvent因为color input不支持userEvent
    await user.click(fillColorInput);
    fillColorInput.value = '#ff0000';
    fillColorInput.dispatchEvent(new Event('change', { bubbles: true }));
    
    await user.click(lineColorInput);
    lineColorInput.value = '#00ff00';
    lineColorInput.dispatchEvent(new Event('change', { bubbles: true }));

    expect(fillColorInput.value).toBe('#ff0000');
    expect(lineColorInput.value).toBe('#00ff00');
  });
});
