/**
 * 文件名: App.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: vitest, @testing-library/react
 * 描述: App 组件单元测试 | App component unit tests
 */

import { describe, it, expect, vi } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../../utils/test-utils';
import App from '../../../src/taskpane/components/App';

// 模拟子组件 | Mock child components
vi.mock('../../../src/taskpane/components/Sidebar', () => ({
  default: ({ currentPage, onNavigate, isCollapsed, onToggleCollapse }: any) => (
    <div data-testid="sidebar">
      <button onClick={() => onNavigate('home')}>Home</button>
      <button onClick={() => onNavigate('create', 'text-insertion')}>Create</button>
      <button onClick={() => onNavigate('query', 'elements-list')}>Query</button>
      <button onClick={onToggleCollapse}>Toggle</button>
      <span>Current: {currentPage}</span>
      <span>Collapsed: {isCollapsed.toString()}</span>
    </div>
  ),
}));

vi.mock('../../../src/taskpane/components/HomePage', () => ({
  default: () => <div data-testid="home-page">Home Page</div>,
}));

vi.mock('../../../src/taskpane/components/ToolsDebugPage', () => ({
  default: ({ selectedTool }: any) => (
    <div data-testid="tools-page">Tools Page - {selectedTool || 'none'}</div>
  ),
}));

describe('App 组件单元测试 | App Component Unit Tests', () => {
  it('应该正确渲染组件 | should render component correctly', () => {
    renderWithProviders(<App />);

    // 验证侧边栏和主页都渲染了 | Verify sidebar and home page are rendered
    expect(screen.getByTestId('sidebar')).toBeInTheDocument();
    expect(screen.getByTestId('home-page')).toBeInTheDocument();
  });

  it('应该默认显示首页 | should display home page by default', () => {
    renderWithProviders(<App />);

    expect(screen.getByTestId('home-page')).toBeInTheDocument();
    expect(screen.getByText('Current: home')).toBeInTheDocument();
  });

  it('应该能够导航到创建元素类页面 | should be able to navigate to create page', async () => {
    const { userEvent } = await import('../../utils/test-utils');
    const user = userEvent.setup();
    
    renderWithProviders(<App />);

    const createButton = screen.getByText('Create');
    await user.click(createButton);

    expect(screen.getByTestId('tools-page')).toBeInTheDocument();
    expect(screen.getByText('Tools Page - text-insertion')).toBeInTheDocument();
    expect(screen.getByText('Current: create')).toBeInTheDocument();
  });

  it('应该能够切换侧边栏折叠状态 | should be able to toggle sidebar collapse state', async () => {
    const { userEvent } = await import('../../utils/test-utils');
    const user = userEvent.setup();
    
    renderWithProviders(<App />);

    // 初始状态应该是折叠的 | Initial state should be collapsed
    expect(screen.getByText('Collapsed: true')).toBeInTheDocument();

    // 点击切换按钮 | Click toggle button
    const toggleButton = screen.getByText('Toggle');
    await user.click(toggleButton);

    // 状态应该变为展开 | State should change to not collapsed
    expect(screen.getByText('Collapsed: false')).toBeInTheDocument();
  });

  it('应该能够在页面间导航 | should be able to navigate between pages', async () => {
    const { userEvent } = await import('../../utils/test-utils');
    const user = userEvent.setup();
    
    renderWithProviders(<App />);

    // 导航到创建元素类页面 | Navigate to create page
    await user.click(screen.getByText('Create'));
    expect(screen.getByTestId('tools-page')).toBeInTheDocument();
    expect(screen.getByText('Current: create')).toBeInTheDocument();

    // 导航回首页 | Navigate back to home page
    await user.click(screen.getByText('Home'));
    expect(screen.getByTestId('home-page')).toBeInTheDocument();
    expect(screen.getByText('Current: home')).toBeInTheDocument();
  });

  it('应该正确传递 currentTool 属性 | should correctly pass currentTool prop', async () => {
    const { userEvent } = await import('../../utils/test-utils');
    const user = userEvent.setup();
    
    renderWithProviders(<App />);

    // 测试创建元素类页面的工具
    await user.click(screen.getByText('Create'));
    expect(screen.getByText('Tools Page - text-insertion')).toBeInTheDocument();

    // 测试查询元素类页面的工具
    await user.click(screen.getByText('Query'));
    expect(screen.getByText('Tools Page - elements-list')).toBeInTheDocument();
  });
});
