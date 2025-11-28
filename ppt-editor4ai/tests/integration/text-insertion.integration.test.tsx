/**
 * 文件名: text-insertion.integration.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: vitest, @testing-library/react
 * 描述: 文本插入功能集成测试 | Text insertion feature integration tests
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import { renderWithProviders, userEvent } from '../utils/test-utils';
import ToolsDebugPage from '../../src/taskpane/components/ToolsDebugPage';

// 模拟 ppt-tools 模块 | Mock ppt-tools module
vi.mock('../../src/ppt-tools', () => ({
  insertText: vi.fn().mockResolvedValue(undefined),
  getCurrentSlideElements: vi.fn().mockResolvedValue([]),
}));

describe('文本插入功能集成测试 | Text Insertion Feature Integration Tests', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('应该完整渲染文本插入工具页面 | should render complete text insertion tool page', () => {
    renderWithProviders(<ToolsDebugPage selectedTool="text-insertion" />);

    // 验证页面标题 | Verify page title
    expect(screen.getByText('文本插入工具')).toBeInTheDocument();
    expect(screen.getByText('在幻灯片中插入文本框，支持自定义位置')).toBeInTheDocument();

    // 验证文本输入组件存在 | Verify text input component exists
    expect(screen.getByLabelText('输入待插入文本')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: '确认插入' })).toBeInTheDocument();
  });

  it('应该能够完成完整的文本插入流程（不带坐标）| should complete full text insertion flow (without coordinates)', async () => {
    const user = userEvent.setup();
    const pptTools = await import('../../src/ppt-tools');
    
    renderWithProviders(<ToolsDebugPage selectedTool="text-insertion" />);

    // 修改文本 | Modify text
    const textarea = screen.getByLabelText('输入待插入文本');
    await user.clear(textarea);
    await user.type(textarea, '集成测试文本');

    // 点击插入按钮 | Click insert button
    const insertButton = screen.getByRole('button', { name: '确认插入' });
    await user.click(insertButton);

    // 验证 insertText 被调用 | Verify insertText was called
    await waitFor(() => {
      expect(pptTools.insertText).toHaveBeenCalledWith('集成测试文本', undefined, undefined);
    });
  });

  it('应该能够完成完整的文本插入流程（带坐标）| should complete full text insertion flow (with coordinates)', async () => {
    const user = userEvent.setup();
    const pptTools = await import('../../src/ppt-tools');
    
    renderWithProviders(<ToolsDebugPage selectedTool="text-insertion" />);

    // 修改文本 | Modify text
    const textarea = screen.getByLabelText('输入待插入文本');
    await user.clear(textarea);
    await user.type(textarea, '带坐标的文本');

    // 输入坐标 | Input coordinates
    const leftInput = screen.getByLabelText('X 坐标 (可选)');
    const topInput = screen.getByLabelText('Y 坐标 (可选)');
    await user.type(leftInput, '100');
    await user.type(topInput, '200');

    // 点击插入按钮 | Click insert button
    const insertButton = screen.getByRole('button', { name: '确认插入' });
    await user.click(insertButton);

    // 验证 insertText 被正确调用 | Verify insertText was called correctly
    await waitFor(() => {
      expect(pptTools.insertText).toHaveBeenCalledWith('带坐标的文本', 100, 200);
    });
  });

  it('应该在未选择工具时显示提示信息 | should show hint when no tool is selected', () => {
    renderWithProviders(<ToolsDebugPage selectedTool="" />);

    expect(screen.getByText('请选择工具')).toBeInTheDocument();
    expect(screen.getByText('从左侧菜单选择要调试的工具')).toBeInTheDocument();
  });

  it('应该能够连续插入多个文本 | should be able to insert multiple texts consecutively', async () => {
    const user = userEvent.setup();
    const pptTools = await import('../../src/ppt-tools');
    
    renderWithProviders(<ToolsDebugPage selectedTool="text-insertion" />);

    const textarea = screen.getByLabelText('输入待插入文本');
    const insertButton = screen.getByRole('button', { name: '确认插入' });

    // 第一次插入 | First insertion
    await user.clear(textarea);
    await user.type(textarea, '第一段文本');
    await user.click(insertButton);

    await waitFor(() => {
      expect(pptTools.insertText).toHaveBeenCalledWith('第一段文本', undefined, undefined);
    });

    // 第二次插入 | Second insertion
    await user.clear(textarea);
    await user.type(textarea, '第二段文本');
    await user.click(insertButton);

    await waitFor(() => {
      expect(pptTools.insertText).toHaveBeenCalledWith('第二段文本', undefined, undefined);
      expect(pptTools.insertText).toHaveBeenCalledTimes(2);
    });
  });

  it('应该正确处理边界坐标值 | should correctly handle boundary coordinate values', async () => {
    const user = userEvent.setup();
    const pptTools = await import('../../src/ppt-tools');
    
    renderWithProviders(<ToolsDebugPage selectedTool="text-insertion" />);

    const leftInput = screen.getByLabelText('X 坐标 (可选)');
    const topInput = screen.getByLabelText('Y 坐标 (可选)');
    const insertButton = screen.getByRole('button', { name: '确认插入' });

    // 测试最大边界值 | Test maximum boundary values
    await user.type(leftInput, '720');
    await user.type(topInput, '540');
    await user.click(insertButton);

    await waitFor(() => {
      expect(pptTools.insertText).toHaveBeenCalledWith('Some text.', 720, 540);
    });
  });
});
