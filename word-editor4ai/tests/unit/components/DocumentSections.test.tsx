/**
 * 文件名: DocumentSections.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: DocumentSections 组件的测试文件 | Test file for DocumentSections component
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import DocumentSections from '../../../src/taskpane/components/tools/DocumentSections';
import * as wordTools from '../../../src/word-tools';
import type { SectionInfo } from '../../../src/word-tools';

// Mock word-tools 模块
vi.mock('../../../src/word-tools', () => ({
  getDocumentSections: vi.fn(),
  HeaderFooterType: {
    FirstPage: 'firstPage',
    OddPages: 'oddPages',
    EvenPages: 'evenPages',
  },
}));

describe('DocumentSections 组件测试 | DocumentSections Component Tests', () => {
  const mockSections: SectionInfo[] = [
    {
      index: 0,
      headers: [
        { type: 'firstPage' as any, exists: true, text: 'First Page Header' },
        { type: 'oddPages' as any, exists: true, text: 'Odd Pages Header' },
        { type: 'evenPages' as any, exists: false },
      ],
      footers: [
        { type: 'firstPage' as any, exists: true, text: 'First Page Footer' },
        { type: 'oddPages' as any, exists: true, text: 'Odd Pages Footer' },
        { type: 'evenPages' as any, exists: false },
      ],
      pageSetup: {
        pageWidth: 612,
        pageHeight: 792,
        topMargin: 72,
        bottomMargin: 72,
        leftMargin: 72,
        rightMargin: 72,
        orientation: 'portrait' as const,
      },
      sectionType: 'nextPage' as const,
      differentFirstPage: true,
      differentOddAndEven: false,
      columnCount: 1,
    },
    {
      index: 1,
      headers: [
        { type: 'firstPage' as any, exists: false },
        { type: 'oddPages' as any, exists: true, text: 'Section 2 Header' },
        { type: 'evenPages' as any, exists: false },
      ],
      footers: [
        { type: 'firstPage' as any, exists: false },
        { type: 'oddPages' as any, exists: true, text: 'Section 2 Footer' },
        { type: 'evenPages' as any, exists: false },
      ],
      pageSetup: {
        pageWidth: 792,
        pageHeight: 612,
        topMargin: 72,
        bottomMargin: 72,
        leftMargin: 72,
        rightMargin: 72,
        orientation: 'landscape' as const,
      },
      sectionType: 'continuous' as const,
      differentFirstPage: false,
      differentOddAndEven: true,
      columnCount: 2,
      columnSpacing: 36,
    },
  ];

  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('应该正确渲染初始状态 | Should render initial state correctly', () => {
    render(<DocumentSections />);

    expect(screen.getByText('获取节信息')).toBeInTheDocument();
    expect(screen.getByText('导出 JSON')).toBeInTheDocument();
    expect(screen.getByText('包含页眉页脚内容')).toBeInTheDocument();
    expect(screen.getByText('包含页面设置详情')).toBeInTheDocument();
  });

  it('应该显示空状态提示 | Should show empty state message', () => {
    render(<DocumentSections />);

    expect(screen.getByText(/点击"获取节信息"按钮查看文档的分节信息/)).toBeInTheDocument();
  });

  it('应该在点击按钮后获取节信息 | Should fetch sections on button click', async () => {
    const user = userEvent.setup();
    vi.mocked(wordTools.getDocumentSections).mockResolvedValue(mockSections);

    render(<DocumentSections />);

    const button = screen.getByText('获取节信息');
    await user.click(button);

    await waitFor(() => {
      expect(wordTools.getDocumentSections).toHaveBeenCalledTimes(1);
    });
  });

  it('应该显示成功消息 | Should show success message', async () => {
    const user = userEvent.setup();
    vi.mocked(wordTools.getDocumentSections).mockResolvedValue(mockSections);

    render(<DocumentSections />);

    const button = screen.getByText('获取节信息');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText(/成功获取 2 个文档节信息/)).toBeInTheDocument();
    });
  });

  it('应该显示节信息列表 | Should display sections list', async () => {
    const user = userEvent.setup();
    vi.mocked(wordTools.getDocumentSections).mockResolvedValue(mockSections);

    render(<DocumentSections />);

    const button = screen.getByText('获取节信息');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText('节 1')).toBeInTheDocument();
      expect(screen.getByText('节 2')).toBeInTheDocument();
    });
  });

  it('应该显示统计信息 | Should display statistics', async () => {
    const user = userEvent.setup();
    vi.mocked(wordTools.getDocumentSections).mockResolvedValue(mockSections);

    render(<DocumentSections />);

    const button = screen.getByText('获取节信息');
    await user.click(button);

    await waitFor(() => {
      // 查找包含统计信息的元素
      const statsElement = screen.getByText('文档共有', { exact: false });
      expect(statsElement).toBeInTheDocument();
      expect(statsElement.textContent).toContain('2');
      expect(statsElement.textContent).toContain('个节');
    });
  });

  it('应该正确处理错误 | Should handle errors correctly', async () => {
    const user = userEvent.setup();
    const errorMessage = '获取节信息失败';
    vi.mocked(wordTools.getDocumentSections).mockRejectedValue(new Error(errorMessage));

    render(<DocumentSections />);

    const button = screen.getByText('获取节信息');
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText(new RegExp(errorMessage))).toBeInTheDocument();
    });
  });

  it('应该支持切换选项 | Should support toggling options', async () => {
    const user = userEvent.setup();
    render(<DocumentSections />);

    const switches = screen.getAllByRole('switch');
    expect(switches).toHaveLength(2);

    // 默认状态：includeContent = false, includePageSetup = true
    expect(switches[0]).not.toBeChecked(); // includeContent
    expect(switches[1]).toBeChecked(); // includePageSetup

    // 切换 includeContent
    await user.click(switches[0]);
    expect(switches[0]).toBeChecked();

    // 切换 includePageSetup
    await user.click(switches[1]);
    expect(switches[1]).not.toBeChecked();
  });

  it('应该在加载时禁用按钮 | Should disable buttons during loading', async () => {
    const user = userEvent.setup();
    vi.mocked(wordTools.getDocumentSections).mockImplementation(
      () => new Promise((resolve) => setTimeout(() => resolve(mockSections), 100))
    );

    render(<DocumentSections />);

    const button = screen.getByText('获取节信息');
    await user.click(button);

    // 按钮应该被禁用
    expect(button).toBeDisabled();

    await waitFor(() => {
      expect(button).not.toBeDisabled();
    });
  });

  it('应该在没有数据时禁用导出按钮 | Should disable export button when no data', () => {
    render(<DocumentSections />);

    const exportButton = screen.getByText('导出 JSON');
    expect(exportButton).toBeDisabled();
  });

  it('应该在有数据时启用导出按钮 | Should enable export button when data exists', async () => {
    const user = userEvent.setup();
    vi.mocked(wordTools.getDocumentSections).mockResolvedValue(mockSections);

    render(<DocumentSections />);

    const fetchButton = screen.getByText('获取节信息');
    await user.click(fetchButton);

    await waitFor(() => {
      const exportButton = screen.getByText('导出 JSON');
      expect(exportButton).not.toBeDisabled();
    });
  });

  it('应该正确传递选项参数 | Should pass options correctly', async () => {
    const user = userEvent.setup();
    vi.mocked(wordTools.getDocumentSections).mockResolvedValue(mockSections);

    render(<DocumentSections />);

    // 切换选项
    const switches = screen.getAllByRole('switch');
    await user.click(switches[0]); // includeContent = true
    await user.click(switches[1]); // includePageSetup = false

    const button = screen.getByText('获取节信息');
    await user.click(button);

    await waitFor(() => {
      expect(wordTools.getDocumentSections).toHaveBeenCalledWith({
        includeContent: true,
        includePageSetup: false,
      });
    });
  });
});
