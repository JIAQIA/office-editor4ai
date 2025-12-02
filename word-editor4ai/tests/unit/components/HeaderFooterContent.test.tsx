/**
 * 文件名: HeaderFooterContent.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 页眉页脚内容获取组件的单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { render, screen, fireEvent, waitFor } from "@testing-library/react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import HeaderFooterContent from "../../../src/taskpane/components/tools/HeaderFooterContent";
import * as wordTools from "../../../src/word-tools";
import type { DocumentHeaderFooterInfo } from "../../../src/word-tools";

// Mock word-tools 模块
vi.mock("../../../src/word-tools", async () => {
  const actual = await vi.importActual("../../../src/word-tools");
  return {
    ...actual,
    getHeaderFooterContent: vi.fn(),
  };
});

const mockGetHeaderFooterContent = vi.mocked(wordTools.getHeaderFooterContent);

// 测试数据 / Test data
const mockHeaderFooterData: DocumentHeaderFooterInfo = {
  totalSections: 2,
  sections: [
    {
      sectionIndex: 0,
      headers: [
        {
          type: "firstPage" as any,
          exists: true,
          text: "首页页眉内容",
        },
        {
          type: "oddPages" as any,
          exists: true,
          text: "奇数页页眉内容",
        },
        {
          type: "evenPages" as any,
          exists: false,
        },
      ],
      footers: [
        {
          type: "firstPage" as any,
          exists: true,
          text: "首页页脚内容",
        },
        {
          type: "oddPages" as any,
          exists: true,
          text: "奇数页页脚内容",
        },
        {
          type: "evenPages" as any,
          exists: false,
        },
      ],
      differentFirstPage: true,
      differentOddAndEven: false,
    },
    {
      sectionIndex: 1,
      headers: [
        {
          type: "firstPage" as any,
          exists: false,
        },
        {
          type: "oddPages" as any,
          exists: true,
          text: "节2奇数页页眉",
        },
        {
          type: "evenPages" as any,
          exists: true,
          text: "节2偶数页页眉",
        },
      ],
      footers: [
        {
          type: "firstPage" as any,
          exists: false,
        },
        {
          type: "oddPages" as any,
          exists: true,
          text: "节2奇数页页脚",
        },
        {
          type: "evenPages" as any,
          exists: true,
          text: "节2偶数页页脚",
        },
      ],
      differentFirstPage: false,
      differentOddAndEven: true,
    },
  ],
  metadata: {
    hasAnyHeader: true,
    hasAnyFooter: true,
    totalHeaders: 4,
    totalFooters: 4,
  },
};

// 辅助函数：渲染组件
const renderComponent = () => {
  return render(
    <FluentProvider theme={webLightTheme}>
      <HeaderFooterContent />
    </FluentProvider>
  );
};

describe("HeaderFooterContent 组件", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("应该正确渲染初始状态", () => {
    renderComponent();

    expect(screen.getByText("获取页眉页脚")).toBeInTheDocument();
    expect(screen.getByText("导出 JSON")).toBeInTheDocument();
    expect(screen.getByText('点击"获取页眉页脚"按钮查看文档的页眉页脚内容')).toBeInTheDocument();
  });

  it("应该显示选项控制", () => {
    renderComponent();

    expect(screen.getByText("节索引（可选，留空获取所有节）")).toBeInTheDocument();
    expect(screen.getByText("包含详细内容元素")).toBeInTheDocument();
    expect(screen.getByText("包含元数据统计")).toBeInTheDocument();
  });

  it("应该在点击获取按钮时调用 API", async () => {
    mockGetHeaderFooterContent.mockResolvedValueOnce(mockHeaderFooterData);

    renderComponent();

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    await waitFor(() => {
      expect(mockGetHeaderFooterContent).toHaveBeenCalledWith({
        includeElements: false,
        includeMetadata: true,
      });
    });
  });

  it("应该在成功获取后显示结果", async () => {
    mockGetHeaderFooterContent.mockResolvedValueOnce(mockHeaderFooterData);

    renderComponent();

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    await waitFor(() => {
      expect(screen.getByText(/成功获取/)).toBeInTheDocument();
    });

    // 检查统计信息 / Check statistics
    expect(screen.getByText("总节数")).toBeInTheDocument();
    const statValues = screen.getAllByText("2");
    expect(statValues.length).toBeGreaterThan(0);
    expect(screen.getByText("页眉总数")).toBeInTheDocument();
    const headerCounts = screen.getAllByText("4");
    expect(headerCounts.length).toBeGreaterThan(0);
  });

  it("应该显示节信息", async () => {
    mockGetHeaderFooterContent.mockResolvedValueOnce(mockHeaderFooterData);

    renderComponent();

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    await waitFor(() => {
      expect(screen.getByText("节 1")).toBeInTheDocument();
      expect(screen.getByText("节 2")).toBeInTheDocument();
    });
  });

  it("应该在发生错误时显示错误信息", async () => {
    const errorMessage = "获取失败";
    mockGetHeaderFooterContent.mockRejectedValueOnce(new Error(errorMessage));

    renderComponent();

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    await waitFor(() => {
      expect(screen.getByText(/获取失败/)).toBeInTheDocument();
    });
  });

  it("应该支持节索引输入", async () => {
    mockGetHeaderFooterContent.mockResolvedValueOnce({
      ...mockHeaderFooterData,
      sections: [mockHeaderFooterData.sections[0]],
      totalSections: 2,
    });

    renderComponent();

    const input = screen.getByPlaceholderText("例如: 0, 1, 2...");
    fireEvent.change(input, { target: { value: "0" } });

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    await waitFor(() => {
      expect(mockGetHeaderFooterContent).toHaveBeenCalledWith({
        sectionIndex: 0,
        includeElements: false,
        includeMetadata: true,
      });
    });
  });

  it("应该在节索引无效时显示错误", async () => {
    renderComponent();

    const input = screen.getByPlaceholderText("例如: 0, 1, 2...");
    fireEvent.change(input, { target: { value: "-1" } });

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    await waitFor(() => {
      expect(screen.getByText(/节索引必须是非负整数/)).toBeInTheDocument();
    });
  });

  it("应该支持切换选项", async () => {
    mockGetHeaderFooterContent.mockResolvedValueOnce(mockHeaderFooterData);

    renderComponent();

    // 切换"包含详细内容元素"选项
    const switches = screen.getAllByRole("switch");
    const includeElementsSwitch = switches[0];
    fireEvent.click(includeElementsSwitch);

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    await waitFor(() => {
      expect(mockGetHeaderFooterContent).toHaveBeenCalledWith({
        includeElements: true,
        includeMetadata: true,
      });
    });
  });

  it("应该支持导出 JSON", async () => {
    mockGetHeaderFooterContent.mockResolvedValueOnce(mockHeaderFooterData);

    // 保存原始函数 / Save original functions
    const originalCreateObjectURL = global.URL.createObjectURL;
    const originalRevokeObjectURL = global.URL.revokeObjectURL;
    const originalCreateElement = document.createElement.bind(document);

    // Mock URL.createObjectURL 和 URL.revokeObjectURL
    global.URL.createObjectURL = vi.fn(() => "blob:mock-url");
    global.URL.revokeObjectURL = vi.fn();

    // Mock document.createElement 和 click
    const mockLink = {
      href: "",
      download: "",
      click: vi.fn(),
    };
    const createElementSpy = vi.spyOn(document, "createElement").mockImplementation((tagName: string) => {
      if (tagName === "a") {
        return mockLink as any;
      }
      return originalCreateElement(tagName);
    });

    renderComponent();

    // 先获取数据
    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    await waitFor(() => {
      expect(screen.getByText(/成功获取/)).toBeInTheDocument();
    });

    // 然后导出
    const exportButton = screen.getByText("导出 JSON");
    fireEvent.click(exportButton);

    await waitFor(() => {
      expect(mockLink.click).toHaveBeenCalled();
      expect(global.URL.createObjectURL).toHaveBeenCalled();
      expect(global.URL.revokeObjectURL).toHaveBeenCalled();
    });

    // 恢复原始函数 / Restore original functions
    createElementSpy.mockRestore();
    global.URL.createObjectURL = originalCreateObjectURL;
    global.URL.revokeObjectURL = originalRevokeObjectURL;
  });

  it("应该在没有数据时禁用导出按钮", () => {
    renderComponent();

    const exportButton = screen.getByText("导出 JSON");
    expect(exportButton).toBeDisabled();
  });

  it("应该在加载时禁用按钮", async () => {
    mockGetHeaderFooterContent.mockImplementation(
      () => new Promise((resolve) => setTimeout(() => resolve(mockHeaderFooterData), 100))
    );

    renderComponent();

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    // 检查按钮是否被禁用
    expect(getButton).toBeDisabled();

    await waitFor(
      () => {
        expect(getButton).not.toBeDisabled();
      },
      { timeout: 200 }
    );
  });

  it("应该显示页眉页脚内容文本", async () => {
    mockGetHeaderFooterContent.mockResolvedValueOnce(mockHeaderFooterData);

    renderComponent();

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    // 等待数据加载完成 / Wait for data to load
    await waitFor(() => {
      expect(screen.getByText("节 1")).toBeInTheDocument();
    });

    // 展开第一个 Accordion 项 / Expand first accordion item
    const accordionButton = screen.getByText("节 1").closest("button");
    if (accordionButton) {
      fireEvent.click(accordionButton);
    }

    // 等待内容显示 / Wait for content to appear
    await waitFor(() => {
      expect(screen.getByText("首页页眉内容")).toBeInTheDocument();
      expect(screen.getByText("奇数页页眉内容")).toBeInTheDocument();
      expect(screen.getByText("首页页脚内容")).toBeInTheDocument();
    });
  });

  it("应该显示节的特殊设置标记", async () => {
    mockGetHeaderFooterContent.mockResolvedValueOnce(mockHeaderFooterData);

    renderComponent();

    const getButton = screen.getByText("获取页眉页脚");
    fireEvent.click(getButton);

    await waitFor(() => {
      // 第一个节有"首页不同"标记
      const badges = screen.getAllByText("首页不同");
      expect(badges.length).toBeGreaterThan(0);

      // 第二个节有"奇偶页不同"标记
      const oddEvenBadges = screen.getAllByText("奇偶页不同");
      expect(oddEvenBadges.length).toBeGreaterThan(0);
    });
  });
});
