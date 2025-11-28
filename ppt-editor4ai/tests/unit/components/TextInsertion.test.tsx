/**
 * 文件名: TextInsertion.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: vitest, @testing-library/react
 * 描述: TextInsertion 组件单元测试 | TextInsertion component unit tests
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { renderWithProviders, userEvent } from '../../utils/test-utils';
import TextInsertion from '../../../src/taskpane/components/tools/TextInsertion';
import * as pptTools from '../../../src/ppt-tools';

// Mock ppt-tools module
vi.mock('../../../src/ppt-tools', () => ({
  insertText: vi.fn().mockResolvedValue(undefined),
}));

describe('TextInsertion 组件单元测试 | TextInsertion Component Unit Tests', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('应该正确渲染组件 | should render component correctly', () => {
    renderWithProviders(<TextInsertion />);

    // 验证文本输入框存在 | Verify text input exists
    expect(screen.getByLabelText('输入待插入文本')).toBeInTheDocument();
    
    // 验证坐标输入框存在 | Verify coordinate inputs exist
    expect(screen.getByLabelText('X 坐标 (可选)')).toBeInTheDocument();
    expect(screen.getByLabelText('Y 坐标 (可选)')).toBeInTheDocument();
    
    // 验证插入按钮存在 | Verify insert button exists
    expect(screen.getByRole('button', { name: '确认插入' })).toBeInTheDocument();
  });

  it('应该显示默认文本 | should display default text', () => {
    renderWithProviders(<TextInsertion />);

    const textarea = screen.getByLabelText('输入待插入文本') as HTMLTextAreaElement;
    expect(textarea.value).toBe('Some text.');
  });

  it('应该能够修改文本内容 | should be able to modify text content', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TextInsertion />);

    const textarea = screen.getByLabelText('输入待插入文本') as HTMLTextAreaElement;
    
    // 清空并输入新文本 | Clear and input new text
    await user.clear(textarea);
    await user.type(textarea, '测试文本');

    expect(textarea.value).toBe('测试文本');
  });

  it('应该能够输入坐标值 | should be able to input coordinate values', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TextInsertion />);

    const leftInput = screen.getByLabelText('X 坐标 (可选)') as HTMLInputElement;
    const topInput = screen.getByLabelText('Y 坐标 (可选)') as HTMLInputElement;

    // 输入坐标 | Input coordinates
    await user.type(leftInput, '100');
    await user.type(topInput, '200');

    expect(leftInput.value).toBe('100');
    expect(topInput.value).toBe('200');
  });

  it('点击插入按钮应该调用 insertText 函数（不带坐标）| clicking insert button should call insertText function (without coordinates)', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TextInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertText).toHaveBeenCalledTimes(1);
      expect(pptTools.insertText).toHaveBeenCalledWith('Some text.', undefined, undefined);
    });
  });

  it('点击插入按钮应该调用 insertText 函数（带坐标）| clicking insert button should call insertText function (with coordinates)', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TextInsertion />);

    const leftInput = screen.getByLabelText('X 坐标 (可选)');
    const topInput = screen.getByLabelText('Y 坐标 (可选)');
    const button = screen.getByRole('button', { name: '确认插入' });

    // 输入坐标 | Input coordinates
    await user.type(leftInput, '150');
    await user.type(topInput, '250');
    
    // 点击插入 | Click insert
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertText).toHaveBeenCalledTimes(1);
      expect(pptTools.insertText).toHaveBeenCalledWith('Some text.', 150, 250);
    });
  });

  it('应该显示位置提示信息 | should display position hint information', () => {
    renderWithProviders(<TextInsertion />);

    expect(screen.getByText(/位置范围提示/)).toBeInTheDocument();
    expect(screen.getByText(/标准 16:9 幻灯片尺寸约为 720×540 磅/)).toBeInTheDocument();
  });

  it('插入按钮应该始终启用 | insert button should always be enabled', () => {
    renderWithProviders(<TextInsertion />);

    const button = screen.getByRole('button', { name: '确认插入' });
    expect(button).not.toBeDisabled();
  });

  it('应该正确解析浮点数坐标 | should correctly parse floating point coordinates', async () => {
    const user = userEvent.setup();
    renderWithProviders(<TextInsertion />);

    const leftInput = screen.getByLabelText('X 坐标 (可选)');
    const topInput = screen.getByLabelText('Y 坐标 (可选)');
    const button = screen.getByRole('button', { name: '确认插入' });

    // 输入浮点数坐标 | Input floating point coordinates
    await user.type(leftInput, '123.45');
    await user.type(topInput, '678.90');
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.insertText).toHaveBeenCalledWith('Some text.', 123.45, 678.90);
    });
  });
});
