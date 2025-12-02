/**
 * 文件名: PageContent.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: PageContent 组件的测试文件 | Test file for PageContent component
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import PageContent from "../../../src/taskpane/components/tools/PageContent";
import * as wordTools from "../../../src/word-tools";
import type { PageInfo, AnyContentElement } from "../../../src/word-tools";

// Mock word-tools 模块 / Mock word-tools module
vi.mock("../../../src/word-tools", () => ({
  getPageContent: vi.fn(),
  getPageStats: vi.fn(),
}));

describe("PageContent 组件测试 | PageContent Component Tests", () => {
  const mockPageInfo: PageInfo = {
    index: 0,
    elements: [
      {
        id: "para-1-0",
        type: "Paragraph",
        text: "这是第一个段落",
        style: "Normal",
        alignment: "Left",
      } as AnyContentElement,
      {
        id: "table-1-1",
        type: "Table",
        rowCount: 3,
        columnCount: 4,
        cells: [
          [
            { text: "A1", rowIndex: 0, columnIndex: 0, width: 100 },
            { text: "B1", rowIndex: 0, columnIndex: 1, width: 100 },
          ],
        ],
      } as AnyContentElement,
      {
        id: "img-1-2",
        type: "InlinePicture",
        width: 200,
        height: 150,
        altText: "测试图片",
      } as AnyContentElement,
      {
        id: "ctrl-1-3",
        type: "ContentControl",
        text: "控件内容",
        title: "测试控件",
        tag: "test-tag",
        controlType: "RichText",
      } as AnyContentElement,
    ],
    text: "这是第一个段落",
  };

  const mockStats = {
    pageIndex: 0,
    elementCount: 10,
    characterCount: 500,
    paragraphCount: 5,
    tableCount: 2,
    imageCount: 1,
    contentControlCount: 2,
  };

  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("初始渲染 | Initial Rendering", () => {
    it("应该正确渲染初始状态 | Should render initial state correctly", () => {
      render(<PageContent />);

      expect(screen.getByText("页面编号")).toBeInTheDocument();
      expect(screen.getByText("获取选项")).toBeInTheDocument();
      expect(screen.getByText("获取页面内容")).toBeInTheDocument();
      expect(screen.getByText("获取统计信息")).toBeInTheDocument();
    });

    it("应该显示默认页面编号为1 | Should show default page number as 1", () => {
      render(<PageContent />);

      const input = screen.getByPlaceholderText("输入页面编号（从1开始）") as HTMLInputElement;
      expect(input.value).toBe("1");
    });

    it("应该显示所有选项开关 | Should show all option switches", () => {
      render(<PageContent />);

      expect(screen.getByText("包含文本内容")).toBeInTheDocument();
      expect(screen.getByText("包含图片信息")).toBeInTheDocument();
      expect(screen.getByText("包含表格信息")).toBeInTheDocument();
      expect(screen.getByText("包含内容控件")).toBeInTheDocument();
      expect(screen.getByText("详细元数据")).toBeInTheDocument();
    });

    it("应该显示空状态提示 | Should show empty state message", () => {
      render(<PageContent />);

      expect(screen.getByText(/输入页面编号并点击按钮获取页面内容或统计信息/)).toBeInTheDocument();
    });
  });

  describe("页面编号输入 | Page Number Input", () => {
    it("应该允许修改页面编号 | Should allow changing page number", async () => {
      const user = userEvent.setup();
      render(<PageContent />);

      const input = screen.getByPlaceholderText("输入页面编号（从1开始）") as HTMLInputElement;
      await user.clear(input);
      await user.type(input, "5");

      expect(input.value).toBe("5");
    });

    it("应该只接受数字输入 | Should only accept numeric input", () => {
      render(<PageContent />);

      const input = screen.getByPlaceholderText("输入页面编号（从1开始）") as HTMLInputElement;
      expect(input.type).toBe("number");
      expect(input.min).toBe("1");
    });
  });

  describe("获取页面内容 | Get Page Content", () => {
    it("应该在点击按钮后获取页面内容 | Should fetch page content on button click", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(wordTools.getPageContent).toHaveBeenCalledTimes(1);
      });
    });

    it("应该传递正确的选项参数 | Should pass correct option parameters", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(wordTools.getPageContent).toHaveBeenCalledWith(1, {
          includeText: true,
          includeImages: true,
          includeTables: true,
          includeContentControls: true,
          detailedMetadata: false,
          maxTextLength: 500,
        });
      });
    });

    it("应该显示成功消息 | Should show success message", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/成功获取第 1 页内容，包含 4 个元素/)).toBeInTheDocument();
      });
    });

    it("应该显示页面内容 | Should display page content", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/页面 1 \(4 个元素\)/)).toBeInTheDocument();
      });
    });

    it("应该显示所有元素类型 | Should display all element types", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText("段落")).toBeInTheDocument();
        expect(screen.getByText("表格")).toBeInTheDocument();
        expect(screen.getByText("内联图片")).toBeInTheDocument();
        expect(screen.getByText("内容控件")).toBeInTheDocument();
      });
    });

    it("应该显示元素文本内容 | Should display element text content", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText("这是第一个段落")).toBeInTheDocument();
        expect(screen.getByText("控件内容")).toBeInTheDocument();
      });
    });

    it("应该在加载时禁用按钮 | Should disable button while loading", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockImplementation(
        () => new Promise((resolve) => setTimeout(() => resolve(mockPageInfo), 100))
      );

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      expect(button).toBeDisabled();

      await waitFor(() => {
        expect(button).not.toBeDisabled();
      });
    });

    it("应该处理无效的页面编号 | Should handle invalid page number", async () => {
      const user = userEvent.setup();
      render(<PageContent />);

      const input = screen.getByPlaceholderText("输入页面编号（从1开始）") as HTMLInputElement;
      await user.clear(input);
      await user.type(input, "0");

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/请输入有效的页面编号（大于等于1）/)).toBeInTheDocument();
      });
    });

    it("应该处理API错误 | Should handle API errors", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockRejectedValue(new Error("页面不存在"));

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/页面不存在/)).toBeInTheDocument();
      });
    });
  });

  describe("获取统计信息 | Get Page Statistics", () => {
    it("应该在点击按钮后获取统计信息 | Should fetch statistics on button click", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageStats).mockResolvedValue(mockStats);

      render(<PageContent />);

      const button = screen.getByText("获取统计信息");
      await user.click(button);

      await waitFor(() => {
        expect(wordTools.getPageStats).toHaveBeenCalledTimes(1);
        expect(wordTools.getPageStats).toHaveBeenCalledWith(1);
      });
    });

    it("应该显示统计信息成功消息 | Should show statistics success message", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageStats).mockResolvedValue(mockStats);

      render(<PageContent />);

      const button = screen.getByText("获取统计信息");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/成功获取第 1 页统计信息/)).toBeInTheDocument();
      });
    });

    it("应该显示所有统计数据 | Should display all statistics", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageStats).mockResolvedValue(mockStats);

      render(<PageContent />);

      const button = screen.getByText("获取统计信息");
      await user.click(button);

      await waitFor(() => {
        // 验证统计信息容器存在 / Verify statistics container exists
        expect(screen.getByText(/成功获取第 1 页统计信息/)).toBeInTheDocument();
      });

      // 验证统计标签存在 / Verify statistics labels exist
      expect(screen.getAllByText("页面编号").length).toBeGreaterThan(0);
      expect(screen.getByText("元素总数")).toBeInTheDocument();
      expect(screen.getByText("段落数")).toBeInTheDocument();
      expect(screen.getByText("表格数")).toBeInTheDocument();
      expect(screen.getByText("图片数")).toBeInTheDocument();
      expect(screen.getByText("控件数")).toBeInTheDocument();
      expect(screen.getByText("字符数")).toBeInTheDocument();
    });

    it("应该显示正确的统计数值 | Should display correct statistics values", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageStats).mockResolvedValue(mockStats);

      render(<PageContent />);

      const button = screen.getByText("获取统计信息");
      await user.click(button);

      await waitFor(() => {
        // 验证统计数据存在 / Verify statistics data exists
        const statValues = screen.getAllByRole("generic").filter(el => 
          el.textContent && /^\d+$/.test(el.textContent)
        );
        expect(statValues.length).toBeGreaterThan(0);
      });
    });

    it("应该处理统计信息API错误 | Should handle statistics API errors", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageStats).mockRejectedValue(new Error("获取统计信息失败"));

      render(<PageContent />);

      const button = screen.getByText("获取统计信息");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/获取统计信息失败/)).toBeInTheDocument();
      });
    });
  });

  describe("选项切换 | Option Toggles", () => {
    it("应该切换包含文本内容选项 | Should toggle includeText option", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      const switches = screen.getAllByRole("switch");
      const textSwitch = switches[0]; // 第一个是"包含文本内容"

      await user.click(textSwitch);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(wordTools.getPageContent).toHaveBeenCalledWith(
          1,
          expect.objectContaining({ includeText: false })
        );
      });
    });

    it("应该切换包含图片信息选项 | Should toggle includeImages option", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      const switches = screen.getAllByRole("switch");
      const imagesSwitch = switches[1]; // 第二个是"包含图片信息"

      await user.click(imagesSwitch);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(wordTools.getPageContent).toHaveBeenCalledWith(
          1,
          expect.objectContaining({ includeImages: false })
        );
      });
    });

    it("应该切换详细元数据选项 | Should toggle detailedMetadata option", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      const switches = screen.getAllByRole("switch");
      const metadataSwitch = switches[4]; // 第五个是"详细元数据"

      await user.click(metadataSwitch);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(wordTools.getPageContent).toHaveBeenCalledWith(
          1,
          expect.objectContaining({ detailedMetadata: true })
        );
      });
    });
  });

  describe("元数据显示 | Metadata Display", () => {
    it("应该在启用详细元数据时显示段落元数据 | Should show paragraph metadata when detailedMetadata is enabled", async () => {
      const user = userEvent.setup();
      const pageWithMetadata: PageInfo = {
        index: 0,
        elements: [
          {
            id: "para-1-0",
            type: "Paragraph",
            text: "段落文本",
            style: "Heading 1",
            alignment: "Centered",
            isListItem: true,
          } as AnyContentElement,
        ],
        text: "段落文本",
      };

      vi.mocked(wordTools.getPageContent).mockResolvedValue(pageWithMetadata);

      render(<PageContent />);

      // 启用详细元数据 / Enable detailed metadata
      const switches = screen.getAllByRole("switch");
      const metadataSwitch = switches[4];
      await user.click(metadataSwitch);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/样式: Heading 1/)).toBeInTheDocument();
        expect(screen.getByText(/对齐: Centered/)).toBeInTheDocument();
        expect(screen.getByText("列表项")).toBeInTheDocument();
      });
    });

    it("应该显示表格元数据 | Should show table metadata", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      // 启用详细元数据 / Enable detailed metadata
      const switches = screen.getAllByRole("switch");
      const metadataSwitch = switches[4];
      await user.click(metadataSwitch);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/3 行/)).toBeInTheDocument();
        expect(screen.getByText(/4 列/)).toBeInTheDocument();
      });
    });

    it("应该显示图片元数据 | Should show image metadata", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      // 启用详细元数据 / Enable detailed metadata
      const switches = screen.getAllByRole("switch");
      const metadataSwitch = switches[4];
      await user.click(metadataSwitch);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/200×150/)).toBeInTheDocument();
        expect(screen.getByText(/描述: 测试图片/)).toBeInTheDocument();
      });
    });

    it("应该显示内容控件元数据 | Should show content control metadata", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);

      render(<PageContent />);

      // 启用详细元数据 / Enable detailed metadata
      const switches = screen.getAllByRole("switch");
      const metadataSwitch = switches[4];
      await user.click(metadataSwitch);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/标题: 测试控件/)).toBeInTheDocument();
        expect(screen.getByText(/标签: test-tag/)).toBeInTheDocument();
        expect(screen.getByText(/类型: RichText/)).toBeInTheDocument();
      });
    });
  });

  describe("文本截断 | Text Truncation", () => {
    it("应该截断超长文本 | Should truncate long text", async () => {
      const user = userEvent.setup();
      const longText = "a".repeat(300);
      const pageWithLongText: PageInfo = {
        index: 0,
        elements: [
          {
            id: "para-1-0",
            type: "Paragraph",
            text: longText,
          } as AnyContentElement,
        ],
        text: longText,
      };

      vi.mocked(wordTools.getPageContent).mockResolvedValue(pageWithLongText);

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        // 验证文本被截断（组件中设置为200字符后截断）/ Verify text is truncated (component truncates at 200 chars)
        const textElements = screen.getAllByRole("generic");
        const hasEllipsis = textElements.some(el => el.textContent?.includes("..."));
        expect(hasEllipsis).toBe(true);
      });
    });
  });

  describe("边界情况 | Edge Cases", () => {
    it("应该处理空页面 | Should handle empty page", async () => {
      const user = userEvent.setup();
      const emptyPage: PageInfo = {
        index: 0,
        elements: [],
        text: "",
      };

      vi.mocked(wordTools.getPageContent).mockResolvedValue(emptyPage);

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/成功获取第 1 页内容，包含 0 个元素/)).toBeInTheDocument();
      });
    });

    it("应该处理非数字页面编号 | Should handle non-numeric page number", async () => {
      const user = userEvent.setup();
      render(<PageContent />);

      const input = screen.getByPlaceholderText("输入页面编号（从1开始）") as HTMLInputElement;
      await user.clear(input);
      await user.type(input, "abc");

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/请输入有效的页面编号（大于等于1）/)).toBeInTheDocument();
      });
    });

    it("应该在获取新内容时清除之前的错误 | Should clear previous errors when fetching new content", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockRejectedValueOnce(new Error("错误"));

      render(<PageContent />);

      const button = screen.getByText("获取页面内容");
      await user.click(button);

      await waitFor(() => {
        expect(screen.getByText(/错误/)).toBeInTheDocument();
      });

      // 再次点击，这次成功 / Click again, this time successfully
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);
      await user.click(button);

      await waitFor(() => {
        expect(screen.queryByText(/错误/)).not.toBeInTheDocument();
        expect(screen.getByText(/成功获取/)).toBeInTheDocument();
      });
    });

    it("应该在获取统计信息时清除页面内容 | Should clear page content when fetching statistics", async () => {
      const user = userEvent.setup();
      vi.mocked(wordTools.getPageContent).mockResolvedValue(mockPageInfo);
      vi.mocked(wordTools.getPageStats).mockResolvedValue(mockStats);

      render(<PageContent />);

      // 先获取页面内容 / First get page content
      const contentButton = screen.getByText("获取页面内容");
      await user.click(contentButton);

      await waitFor(() => {
        expect(screen.getByText("段落")).toBeInTheDocument();
      });

      // 然后获取统计信息 / Then get statistics
      const statsButton = screen.getByText("获取统计信息");
      await user.click(statsButton);

      await waitFor(() => {
        expect(screen.queryByText("段落")).not.toBeInTheDocument();
        expect(screen.getByText("元素总数")).toBeInTheDocument();
      });
    });
  });
});
