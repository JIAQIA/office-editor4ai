/**
 * 文件名: SlideDeletion.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片删除组件单元测试
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { screen, waitFor } from "@testing-library/react";
import { renderWithProviders, userEvent } from "../../utils/test-utils";
import { SlideDeletion } from "../../../src/taskpane/components/tools/SlideDeletion";
import * as pptTools from "../../../src/ppt-tools";

// Mock ppt-tools
vi.mock("../../../src/ppt-tools", () => ({
  deleteCurrentSlide: vi.fn(),
  deleteSlidesByNumbers: vi.fn(),
}));

// Mock PowerPoint API
const mockPowerPoint = {
  run: vi.fn(),
};

global.PowerPoint = mockPowerPoint as any;

describe("SlideDeletion Component", () => {
  beforeEach(() => {
    vi.clearAllMocks();

    // Mock PowerPoint.run for fetching slide info
    mockPowerPoint.run.mockImplementation(async (callback: any) => {
      const mockSlides = [
        { id: "slide1", load: vi.fn() },
        { id: "slide2", load: vi.fn() },
        { id: "slide3", load: vi.fn() },
      ];

      const mockContext = {
        presentation: {
          slides: {
            items: mockSlides,
            load: vi.fn(),
          },
          getSelectedSlides: vi.fn(() => ({
            items: [mockSlides[1]], // 选中第2页
            load: vi.fn(),
          })),
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };
      await callback(mockContext);
    });
  });

  it("应该正确渲染组件", () => {
    renderWithProviders(<SlideDeletion />);

    expect(screen.getByText("幻灯片删除调试工具")).toBeInTheDocument();
    expect(screen.getByText(/总页数:/)).toBeInTheDocument();
    expect(screen.getByText(/当前页:/)).toBeInTheDocument();
  });

  it("应该能够删除当前幻灯片", async () => {
    const user = userEvent.setup();
    const mockDeleteResult = {
      success: true,
      deletedCount: 1,
      failedCount: 0,
      message: "成功删除第 2 页",
      details: {
        deleted: [2],
        notFound: [],
        errors: [],
      },
    };

    vi.mocked(pptTools.deleteCurrentSlide).mockResolvedValue(mockDeleteResult);

    renderWithProviders(<SlideDeletion />);

    // 等待组件加载完成
    await waitFor(() => {
      expect(
        screen.getByText((_content, element) => {
          return element?.textContent === "总页数: 3 页";
        })
      ).toBeInTheDocument();
    });

    // 点击删除当前页按钮
    const deleteButton = screen.getByRole("button", { name: /删除当前页/ });
    await user.click(deleteButton);

    await waitFor(() => {
      expect(pptTools.deleteCurrentSlide).toHaveBeenCalled();
      expect(screen.getByText("成功删除第 2 页")).toBeInTheDocument();
    });
  });

  it("应该能够按页码删除幻灯片", async () => {
    const user = userEvent.setup();
    const mockDeleteResult = {
      success: true,
      deletedCount: 2,
      failedCount: 0,
      message: "删除操作完成: 成功 2 页",
      details: {
        deleted: [1, 3],
        notFound: [],
        errors: [],
      },
    };

    vi.mocked(pptTools.deleteSlidesByNumbers).mockResolvedValue(mockDeleteResult);

    renderWithProviders(<SlideDeletion />);

    // 输入页码
    const textarea = screen.getByPlaceholderText(/输入页码/);
    await user.clear(textarea);
    await user.type(textarea, "1, 3");

    // 点击删除按钮
    const deleteButton = screen.getByText("删除指定页码");
    await user.click(deleteButton);

    await waitFor(() => {
      expect(pptTools.deleteSlidesByNumbers).toHaveBeenCalledWith([1, 3]);
      expect(screen.getByText("删除操作完成: 成功 2 页")).toBeInTheDocument();
    });
  });

  it("应该处理页码不存在的情况", async () => {
    const user = userEvent.setup();
    const mockDeleteResult = {
      success: true,
      deletedCount: 1,
      failedCount: 1,
      message: "删除操作完成: 成功 1 页, 页码不存在 1 页 (10)",
      details: {
        deleted: [1],
        notFound: [10],
        errors: [],
      },
    };

    vi.mocked(pptTools.deleteSlidesByNumbers).mockResolvedValue(mockDeleteResult);

    renderWithProviders(<SlideDeletion />);

    // 输入包含不存在页码的列表
    const textarea = screen.getByPlaceholderText(/输入页码/);
    await user.clear(textarea);
    await user.type(textarea, "1, 10");

    // 点击删除按钮
    const deleteButton = screen.getByText("删除指定页码");
    await user.click(deleteButton);

    await waitFor(() => {
      expect(pptTools.deleteSlidesByNumbers).toHaveBeenCalledWith([1, 10]);
      expect(screen.getByText(/页码不存在 1 页 \(10\)/)).toBeInTheDocument();
    });
  });

  it("应该显示删除详情", async () => {
    const user = userEvent.setup();
    const mockDeleteResult = {
      success: true,
      deletedCount: 2,
      failedCount: 1,
      message: "删除操作完成: 成功 2 页, 页码不存在 1 页 (5)",
      details: {
        deleted: [1, 3],
        notFound: [5],
        errors: [],
      },
    };

    vi.mocked(pptTools.deleteSlidesByNumbers).mockResolvedValue(mockDeleteResult);

    renderWithProviders(<SlideDeletion />);

    const textarea = screen.getByPlaceholderText(/输入页码/);
    await user.clear(textarea);
    await user.type(textarea, "1, 3, 5");

    const deleteButton = screen.getByText("删除指定页码");
    await user.click(deleteButton);

    await waitFor(() => {
      expect(screen.getByText("删除详情:")).toBeInTheDocument();
      expect(screen.getByText(/成功删除: 1, 3/)).toBeInTheDocument();
      expect(screen.getByText(/页码不存在: 5/)).toBeInTheDocument();
    });
  });

  it("应该验证输入的页码", async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideDeletion />);

    // 输入无效的页码（只有逗号和空格）
    const textarea = screen.getByPlaceholderText(/输入页码/);
    await user.type(textarea, ", , ,");

    // 点击删除按钮
    const deleteButton = screen.getByText("删除指定页码");
    await user.click(deleteButton);

    await waitFor(() => {
      expect(screen.getByText("请输入有效的页码")).toBeInTheDocument();
    });
  });

  it("应该支持快速选择页码", async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideDeletion />);

    await waitFor(() => {
      expect(
        screen.getByText((_content, element) => {
          return element?.textContent === "总页数: 3 页";
        })
      ).toBeInTheDocument();
    });

    // 点击快速选择按钮
    const quickSelectButton = screen.getByRole("button", { name: "1" });
    await user.click(quickSelectButton);

    // 验证页码已添加到输入框
    const textarea = screen.getByPlaceholderText(/输入页码/) as HTMLTextAreaElement;
    expect(textarea.value).toBe("1");
  });
});
