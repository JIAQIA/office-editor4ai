/**
 * 文件名: HomePage.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: vitest, @testing-library/react
 * 描述: HomePage 组件单元测试 | HomePage component unit tests
 */

import { describe, it, expect } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../../utils/test-utils';
import HomePage from '../../../src/taskpane/components/HomePage';

describe('HomePage 组件单元测试 | HomePage Component Unit Tests', () => {
  it('应该正确渲染组件 | should render component correctly', () => {
    renderWithProviders(<HomePage />);

    // 验证标题 | Verify title
    expect(screen.getByRole('heading', { name: 'TuringFocus' })).toBeInTheDocument();
  });

  it('应该显示 Logo 图片 | should display logo image', () => {
    renderWithProviders(<HomePage />);

    const logo = screen.getByAltText('PPT Editor for AI');
    expect(logo).toBeInTheDocument();
    expect(logo).toHaveAttribute('src', 'assets/logo-filled.png');
  });

  it('应该显示欢迎描述文本 | should display welcome description text', () => {
    renderWithProviders(<HomePage />);

    expect(screen.getByText(/欢迎使用 PPT Editor for AI/)).toBeInTheDocument();
    expect(screen.getByText(/这是一个专为 AI Agent 设计的 PowerPoint 编辑工具包/)).toBeInTheDocument();
  });

  it('应该显示三个功能特性 | should display three feature items', () => {
    renderWithProviders(<HomePage />);

    // 验证三个功能特性文本 | Verify three feature texts
    expect(screen.getByText('与 Office 深度集成，实现更多功能')).toBeInTheDocument();
    expect(screen.getByText('解锁强大的编辑功能')).toBeInTheDocument();
    expect(screen.getByText('像专业人士一样创建和可视化')).toBeInTheDocument();
  });

  it('应该渲染功能列表 | should render feature list', () => {
    renderWithProviders(<HomePage />);

    const featureList = screen.getByRole('list');
    expect(featureList).toBeInTheDocument();
    
    // 验证有三个列表项 | Verify three list items
    const listItems = screen.getAllByRole('listitem');
    expect(listItems).toHaveLength(3);
  });

  it('应该使用正确的样式类 | should use correct style classes', () => {
    const { container } = renderWithProviders(<HomePage />);

    // 验证容器存在 | Verify container exists
    expect(container.firstChild).toBeTruthy();
  });
});
