/**
 * 文件名: DocumentStructure.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 文档结构组件的测试文件 | Test file for DocumentStructure component
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as React from 'react';
import { renderWithProviders, screen, waitFor, userEvent } from '../../utils/test-utils';
import { DocumentStructure } from '../../../src/taskpane/components/tools/DocumentStructure';
import * as documentStructureTools from '../../../src/word-tools/documentStructure';
import type { DocumentOutline } from '../../../src/word-tools/documentStructure';

// Mock word-tools 模块
vi.mock('../../../src/word-tools/documentStructure', async () => {
  const actual = await vi.importActual('../../../src/word-tools/documentStructure');
  return {
    ...actual,
    getDocumentOutline: vi.fn(),
    getDocumentOutlineFlat: vi.fn(),
    navigateToOutlineNode: vi.fn(),
    exportOutlineAsMarkdown: vi.fn(),
    exportOutlineAsJSON: vi.fn(),
  };
});

describe('DocumentStructure - 文档结构组件', () => {
  const mockOutline: DocumentOutline = {
    nodes: [
      {
        id: 'heading-0',
        text: '第一章',
        level: 1,
        style: 'Heading 1',
        children: [
          {
            id: 'heading-1',
            text: '1.1 节',
            level: 2,
            style: 'Heading 2',
            children: [],
            index: 1,
          },
        ],
        index: 0,
      },
      {
        id: 'heading-2',
        text: '第二章',
        level: 1,
        style: 'Heading 1',
        children: [],
        index: 2,
      },
    ],
    totalHeadings: 3,
    maxDepth: 2,
    levelCounts: { 1: 2, 2: 1 },
  };

  beforeEach(() => {
    vi.clearAllMocks();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('组件渲染 | Component rendering', () => {
    it('应该正确渲染组件 | Should render component correctly', () => {
      renderWithProviders(<DocumentStructure />);

      expect(screen.getByText('文档结构获取')).toBeInTheDocument();
      expect(screen.getByText('获取文档的大纲结构（标题层级）')).toBeInTheDocument();
      expect(screen.getByText('获取大纲')).toBeInTheDocument();
    });

    it('应该显示所有选项开关 | Should display all option switches', () => {
      renderWithProviders(<DocumentStructure />);

      expect(screen.getByText('包含格式信息')).toBeInTheDocument();
      expect(screen.getByText('树形结构')).toBeInTheDocument();
      expect(screen.getByText(/最大层级深度/)).toBeInTheDocument();
    });

    it('应该显示最大层级深度输入框 | Should display max depth input', () => {
      renderWithProviders(<DocumentStructure />);

      const input = screen.getByRole('spinbutton') as HTMLInputElement;
      expect(input).toBeInTheDocument();
      expect(input.value).toBe('0');
      expect(input.min).toBe('0');
      expect(input.max).toBe('9');
    });
  });

  describe('获取大纲功能 | Get outline functionality', () => {
    it('应该成功获取文档大纲 | Should successfully get document outline', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(mockOutline);

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(documentStructureTools.getDocumentOutline).toHaveBeenCalled();
        expect(screen.getByText('统计信息')).toBeInTheDocument();
        expect(screen.getByText('总标题数')).toBeInTheDocument();
        expect(screen.getByText('3')).toBeInTheDocument();
      });
    });

    it('应该在加载时显示加载状态 | Should show loading state during fetch', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockImplementation(
        () => new Promise((resolve) => setTimeout(() => resolve(mockOutline), 100))
      );

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      // 检查按钮是否被禁用
      expect(button).toBeDisabled();

      await waitFor(() => {
        expect(button).not.toBeDisabled();
      });
    });

    it('应该处理获取大纲失败的情况 | Should handle outline fetch failure', async () => {
      const user = userEvent.setup();
      const errorMessage = '获取文档大纲失败';
      vi.mocked(documentStructureTools.getDocumentOutline).mockRejectedValue(
        new Error(errorMessage)
      );

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(errorMessage)).toBeInTheDocument();
      });
    });

    it('应该根据选项调用正确的函数 | Should call correct function based on options', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutlineFlat).mockResolvedValue(
        mockOutline.nodes
      );

      renderWithProviders(<DocumentStructure />);

      // 切换到扁平列表模式
      const treeSwitch = screen.getAllByRole('switch')[1]; // 第二个开关是树形结构
      await user.click(treeSwitch);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(documentStructureTools.getDocumentOutlineFlat).toHaveBeenCalled();
      });
    });

    it('应该传递正确的选项参数 | Should pass correct option parameters', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(mockOutline);

      renderWithProviders(<DocumentStructure />);

      // 启用格式信息
      const formatSwitch = screen.getAllByRole('switch')[0];
      await user.click(formatSwitch);

      // 设置最大层级
      const depthInput = screen.getByRole('spinbutton') as HTMLInputElement;
      await user.clear(depthInput);
      await user.type(depthInput, '3');

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(documentStructureTools.getDocumentOutline).toHaveBeenCalledWith({
          includeFormat: true,
          maxDepth: 3,
        });
      });
    });
  });

  describe('大纲显示 | Outline display', () => {
    it('应该显示统计信息 | Should display statistics', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(mockOutline);

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(
        () => {
          expect(screen.getByText('统计信息')).toBeInTheDocument();
          expect(screen.getByText('总标题数')).toBeInTheDocument();
          expect(screen.getByText('最大层级')).toBeInTheDocument();
          // 验证统计数值存在
          expect(screen.getAllByText('3').length).toBeGreaterThan(0); // totalHeadings
          expect(screen.getAllByText('2').length).toBeGreaterThan(0); // maxDepth
        },
        { timeout: 5000 }
      );
    });

    it('应该显示各层级的标题数量 | Should display heading count by level', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(mockOutline);

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('H1 数量')).toBeInTheDocument();
        expect(screen.getByText('H2 数量')).toBeInTheDocument();
      });
    });

    it('应该显示大纲树节点 | Should display outline tree nodes', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(mockOutline);

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(
        () => {
          // 验证根节点显示
          expect(screen.getByText('第一章')).toBeInTheDocument();
          expect(screen.getByText('第二章')).toBeInTheDocument();
        },
        { timeout: 5000 }
      );

      // 验证显示了标题级别徽章
      const badges = screen.getAllByText(/H\d/);
      expect(badges.length).toBeGreaterThan(0);
    });

    it('应该显示空状态提示 | Should display empty state message', async () => {
      const user = userEvent.setup();
      const emptyOutline: DocumentOutline = {
        nodes: [],
        totalHeadings: 0,
        maxDepth: 0,
        levelCounts: {},
      };
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(emptyOutline);

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('文档中没有找到标题')).toBeInTheDocument();
      });
    });
  });

  describe('导航功能 | Navigation functionality', () => {
    it('应该能够跳转到标题节点 | Should be able to navigate to heading node', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(mockOutline);
      vi.mocked(documentStructureTools.navigateToOutlineNode).mockResolvedValue();

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('第一章')).toBeInTheDocument();
      });

      // 点击标题进行跳转
      const heading = screen.getByText('第一章');
      await user.click(heading);

      await waitFor(() => {
        expect(documentStructureTools.navigateToOutlineNode).toHaveBeenCalledWith('heading-0');
      });
    });

    it('应该处理跳转失败的情况 | Should handle navigation failure', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(mockOutline);
      vi.mocked(documentStructureTools.navigateToOutlineNode).mockRejectedValue(
        new Error('跳转失败')
      );

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('第一章')).toBeInTheDocument();
      });

      const heading = screen.getByText('第一章');
      await user.click(heading);

      await waitFor(() => {
        expect(screen.getByText('跳转失败')).toBeInTheDocument();
      });
    });
  });

  describe('导出功能 | Export functionality', () => {
    it('应该显示导出按钮 | Should display export button', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(mockOutline);
      
      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('导出')).toBeInTheDocument();
      });
    });

    it('应该能够打开导出菜单 | Should be able to open export menu', async () => {
      const user = userEvent.setup();
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(mockOutline);

      renderWithProviders(<DocumentStructure />);

      const getButton = screen.getByText('获取大纲');
      await user.click(getButton);

      await waitFor(() => {
        expect(screen.getByText('导出')).toBeInTheDocument();
      });

      const exportButton = screen.getByText('导出');
      await user.click(exportButton);

      await waitFor(() => {
        expect(screen.getByText('导出为 Markdown')).toBeInTheDocument();
        expect(screen.getByText('导出为 JSON')).toBeInTheDocument();
        expect(screen.getByText('复制到剪贴板')).toBeInTheDocument();
      });
    });

    it('应该能够复制到剪贴板 | Should be able to copy to clipboard', async () => {
      const user = userEvent.setup();
      const mockWriteText = vi.fn().mockResolvedValue(undefined);
      Object.defineProperty(navigator, 'clipboard', {
        value: {
          writeText: mockWriteText,
        },
        writable: true,
        configurable: true,
      });

      // Mock window.alert
      vi.spyOn(window, 'alert').mockImplementation(() => {});

      renderWithProviders(<DocumentStructure />);

      const getButton = screen.getByText('获取大纲');
      await user.click(getButton);

      await waitFor(() => {
        expect(screen.getByText('导出')).toBeInTheDocument();
      });

      const exportButton = screen.getByText('导出');
      await user.click(exportButton);

      await waitFor(() => {
        expect(screen.getByText('复制到剪贴板')).toBeInTheDocument();
      });

      const copyOption = screen.getByText('复制到剪贴板');
      await user.click(copyOption);

      await waitFor(() => {
        expect(documentStructureTools.exportOutlineAsJSON).toHaveBeenCalledWith(mockOutline);
        expect(mockWriteText).toHaveBeenCalled();
        expect(window.alert).toHaveBeenCalledWith('已复制到剪贴板');
      });
    });

    it('应该处理复制失败的情况 | Should handle copy failure', async () => {
      const user = userEvent.setup();
      const mockWriteText = vi.fn().mockRejectedValue(new Error('Copy failed'));
      Object.defineProperty(navigator, 'clipboard', {
        value: {
          writeText: mockWriteText,
        },
        writable: true,
        configurable: true,
      });

      renderWithProviders(<DocumentStructure />);

      const getButton = screen.getByText('获取大纲');
      await user.click(getButton);

      await waitFor(() => {
        expect(screen.getByText('导出')).toBeInTheDocument();
      });

      const exportButton = screen.getByText('导出');
      await user.click(exportButton);

      await waitFor(() => {
        expect(screen.getByText('复制到剪贴板')).toBeInTheDocument();
      });

      const copyOption = screen.getByText('复制到剪贴板');
      await user.click(copyOption);

      await waitFor(() => {
        expect(screen.getByText('复制到剪贴板失败')).toBeInTheDocument();
      });
    });
  });

  describe('选项交互 | Options interaction', () => {
    it('应该能够切换格式信息选项 | Should be able to toggle format option', async () => {
      const user = userEvent.setup();
      renderWithProviders(<DocumentStructure />);

      const formatSwitch = screen.getAllByRole('switch')[0];
      expect(formatSwitch).not.toBeChecked();

      await user.click(formatSwitch);
      expect(formatSwitch).toBeChecked();

      await user.click(formatSwitch);
      expect(formatSwitch).not.toBeChecked();
    });

    it('应该能够切换树形结构选项 | Should be able to toggle tree structure option', async () => {
      const user = userEvent.setup();
      renderWithProviders(<DocumentStructure />);

      const treeSwitch = screen.getAllByRole('switch')[1];
      expect(treeSwitch).toBeChecked(); // 默认开启

      await user.click(treeSwitch);
      expect(treeSwitch).not.toBeChecked();
    });

    it('应该能够修改最大层级深度 | Should be able to change max depth', async () => {
      const user = userEvent.setup();
      renderWithProviders(<DocumentStructure />);

      const depthInput = screen.getByRole('spinbutton') as HTMLInputElement;
      expect(depthInput.value).toBe('0');

      await user.clear(depthInput);
      await user.type(depthInput, '5');

      expect(depthInput.value).toBe('5');
    });

    it('应该限制最大层级深度的范围 | Should limit max depth range', () => {
      renderWithProviders(<DocumentStructure />);

      const depthInput = screen.getByRole('spinbutton') as HTMLInputElement;
      expect(depthInput.min).toBe('0');
      expect(depthInput.max).toBe('9');
    });
  });

  describe('边界情况 | Edge cases', () => {
    it('应该处理包含格式信息的大纲 | Should handle outline with format info', async () => {
      const user = userEvent.setup();
      const outlineWithFormat: DocumentOutline = {
        ...mockOutline,
        nodes: [
          {
            ...mockOutline.nodes[0],
            format: {
              font: 'Arial',
              fontSize: 16,
              bold: true,
              italic: false,
              color: '#000000',
              alignment: 'Left',
            },
          },
        ],
      };
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(outlineWithFormat);

      renderWithProviders(<DocumentStructure />);

      // 启用格式信息
      const formatSwitch = screen.getAllByRole('switch')[0];
      await user.click(formatSwitch);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('第一章')).toBeInTheDocument();
        // 格式信息应该显示
        expect(screen.getByText(/Arial/)).toBeInTheDocument();
      });
    });

    it('应该处理空文本的标题 | Should handle headings with empty text', async () => {
      const user = userEvent.setup();
      const outlineWithEmptyText: DocumentOutline = {
        nodes: [
          {
            id: 'heading-0',
            text: '',
            level: 1,
            style: 'Heading 1',
            children: [],
            index: 0,
          },
        ],
        totalHeadings: 1,
        maxDepth: 1,
        levelCounts: { 1: 1 },
      };
      vi.mocked(documentStructureTools.getDocumentOutline).mockResolvedValue(
        outlineWithEmptyText
      );

      renderWithProviders(<DocumentStructure />);

      const button = screen.getByText('获取大纲');
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText('(空标题)')).toBeInTheDocument();
      });
    });
  });
});
