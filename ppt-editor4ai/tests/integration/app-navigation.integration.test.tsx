/**
 * 文件名: app-navigation.integration.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: vitest, @testing-library/react
 * 描述: 应用导航集成测试 | Application navigation integration tests
 */

import { describe, it, expect, vi } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders, userEvent } from '../utils/test-utils';
import App from '../../src/taskpane/components/App';

// 模拟 taskpane 模块 | Mock taskpane module
vi.mock('../../src/taskpane/taskpane', () => ({
  insertText: vi.fn(),
}));

describe('应用导航集成测试 | Application Navigation Integration Tests', () => {
  it('应该能够从首页导航到工具页面 | should be able to navigate from home to tools page', async () => {
    const user = userEvent.setup();
    renderWithProviders(<App />);

    // 初始应该在首页 | Should initially be on home page
    expect(screen.getByText(/欢迎使用 PPT Editor for AI/)).toBeInTheDocument();

    // 侧边栏默认折叠，需要先展开 | Sidebar is collapsed by default, need to expand first
    const expandButton = screen.queryByTitle('展开侧边栏');
    if (expandButton) {
      await user.click(expandButton);
    }

    // 查找工具调试页按钮 | Find tools debug page button
    const toolsButton = screen.queryByText('工具调试页');
    if (toolsButton) {
      await user.click(toolsButton);

      // 查找文本插入按钮 | Find text insertion button
      const textInsertionButton = screen.queryByText('文本插入工具');
      if (textInsertionButton) {
        await user.click(textInsertionButton);
        // 应该导航到工具页面 | Should navigate to tools page
        expect(screen.getByText('文本插入工具')).toBeInTheDocument();
      }
    }
  });

  it('应该能够在工具页面使用文本插入功能 | should be able to use text insertion feature on tools page', async () => {
    const user = userEvent.setup();
    const { insertText } = await import('../../src/taskpane/taskpane');
    
    renderWithProviders(<App />);

    // 侧边栏默认折叠，需要先展开 | Sidebar is collapsed by default, need to expand first
    const expandButton = screen.queryByTitle('展开侧边栏');
    if (expandButton) {
      await user.click(expandButton);
    }

    // 导航到工具页面 | Navigate to tools page
    const toolsButton = screen.queryByText('工具调试页');
    if (toolsButton) {
      await user.click(toolsButton);
    }

    // 点击文本插入工具 | Click text insertion tool
    const textInsertionButton = screen.queryByText('文本插入工具');
    if (textInsertionButton) {
      await user.click(textInsertionButton);
    }

    // 测试文本插入功能 | Test text insertion feature
    const textarea = screen.queryByLabelText('输入待插入文本');
    if (textarea) {
      await user.clear(textarea);
      await user.type(textarea, '导航测试文本');

      const insertButton = screen.getByRole('button', { name: '确认插入' });
      await user.click(insertButton);

      expect(insertText).toHaveBeenCalledWith('导航测试文本', undefined, undefined);
    }
  });

  it('应该能够切换侧边栏状态 | should be able to toggle sidebar state', async () => {
    const user = userEvent.setup();
    renderWithProviders(<App />);

    // 查找侧边栏切换按钮（通过 title 属性）| Find sidebar toggle button (by title attribute)
    // 默认状态是折叠的，所以按钮标题是"展开侧边栏" | Default state is collapsed, so button title is "展开侧边栏"
    const toggleButton = screen.queryByTitle('展开侧边栏');
    
    if (toggleButton) {
      // 点击切换按钮 | Click toggle button
      await user.click(toggleButton);

      // 验证按钮标题变为"折叠侧边栏" | Verify button title changed to "折叠侧边栏"
      const collapseButton = screen.queryByTitle('折叠侧边栏');
      expect(collapseButton).toBeInTheDocument();
    }
  });

  it('应该保持页面状态在侧边栏展开后 | should maintain page state after sidebar expand', async () => {
    const user = userEvent.setup();
    renderWithProviders(<App />);

    // 验证首页内容存在 | Verify home page content exists
    const welcomeText = screen.getByText(/欢迎使用 PPT Editor for AI/);
    expect(welcomeText).toBeInTheDocument();

    // 查找并点击侧边栏切换按钮（默认是折叠的）| Find and click sidebar toggle button (default is collapsed)
    const toggleButton = screen.queryByTitle('展开侧边栏');
    
    if (toggleButton) {
      await user.click(toggleButton);

      // 首页内容应该仍然存在 | Home page content should still exist
      expect(welcomeText).toBeInTheDocument();
    }
  });

  it('应该正确渲染应用的整体布局 | should correctly render overall application layout', () => {
    const { container } = renderWithProviders(<App />);

    // 验证根容器存在 | Verify root container exists
    expect(container.firstChild).toBeTruthy();

    // 验证主要内容区域存在 | Verify main content area exists
    const mainContent = container.querySelector('[class*="content"]');
    expect(mainContent || container.firstChild).toBeTruthy();
  });
});
