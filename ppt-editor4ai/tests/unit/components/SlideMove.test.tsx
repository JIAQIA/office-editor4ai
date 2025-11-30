/**
 * 文件名: SlideMove.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: SlideMove 组件的单元测试 | SlideMove component unit tests
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { screen, waitFor } from "@testing-library/react";
import { renderWithProviders, userEvent } from "../../utils/test-utils";
import { SlideMove } from "../../../src/taskpane/components/tools/SlideMove";
import * as pptTools from "../../../src/ppt-tools";

// Mock ppt-tools 模块
vi.mock("../../../src/ppt-tools", () => ({
  moveSlide: vi.fn(),
  moveCurrentSlide: vi.fn(),
  swapSlides: vi.fn(),
  getAllSlidesInfo: vi.fn(),
}));

// Mock PowerPoint global
const mockPowerPoint = {
  run: vi.fn(),
};

describe("SlideMove 组件", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    // @ts-expect-error - Mock PowerPoint global
    global.PowerPoint = mockPowerPoint;

    // 默认 mock 返回值
    vi.mocked(pptTools.getAllSlidesInfo).mockResolvedValue([
      { index: 1, id: "slide-1", title: "Slide 1" },
      { index: 2, id: "slide-2", title: "Slide 2" },
      { index: 3, id: "slide-3", title: "Slide 3" },
      { index: 4, id: "slide-4", title: "Slide 4" },
      { index: 5, id: "slide-5", title: "Slide 5" },
    ]);

    mockPowerPoint.run.mockImplementation(async (callback: any) => {
      const mockSlides = [
        { id: "slide-1", load: vi.fn().mockReturnThis() },
        { id: "slide-2", load: vi.fn().mockReturnThis() },
        { id: "slide-3", load: vi.fn().mockReturnThis() },
        { id: "slide-4", load: vi.fn().mockReturnThis() },
        { id: "slide-5", load: vi.fn().mockReturnThis() },
      ];

      const mockContext = {
        presentation: {
          getSelectedSlides: () => ({
            items: [mockSlides[0]],
            load: vi.fn().mockReturnThis(),
          }),
          slides: {
            items: mockSlides,
            load: vi.fn().mockReturnThis(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      await callback(mockContext);
    });
  });

  describe("渲染", () => {
    it("应该正确渲染组件", async () => {
      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByText("幻灯片移动工具")).toBeInTheDocument();
      });

      expect(screen.getByText("方法1: 移动指定幻灯片")).toBeInTheDocument();
      expect(screen.getByText("方法2: 移动当前选中的幻灯片")).toBeInTheDocument();
      expect(screen.getByText("方法3: 交换两张幻灯片位置")).toBeInTheDocument();
    });

    it("应该显示幻灯片信息概览", async () => {
      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByText(/总幻灯片数:/)).toBeInTheDocument();
      });

      // 验证总数显示为5
      await waitFor(() => {
        const overview = screen.getByText((_content, element) => {
          return element?.textContent === "总幻灯片数: 5" || false;
        });
        expect(overview).toBeInTheDocument();
      });
    });

    it("应该显示幻灯片列表", async () => {
      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByText("Slide 1")).toBeInTheDocument();
      });

      expect(screen.getByText("Slide 2")).toBeInTheDocument();
      expect(screen.getByText("Slide 3")).toBeInTheDocument();
      expect(screen.getByText("Slide 4")).toBeInTheDocument();
      expect(screen.getByText("Slide 5")).toBeInTheDocument();
    });
  });

  describe("方法1: 移动指定幻灯片", () => {
    it("应该成功移动幻灯片", async () => {
      const user = userEvent.setup();
      vi.mocked(pptTools.moveSlide).mockResolvedValue({
        success: true,
        message: "成功将幻灯片从位置 1 移动到位置 3",
        fromIndex: 1,
        toIndex: 3,
        totalSlides: 5,
      });

      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText("源位置:")).toBeInTheDocument();
      });

      const fromInput = screen.getByLabelText("源位置:") as HTMLInputElement;
      const toInput = screen.getByLabelText("目标位置:") as HTMLInputElement;
      const moveButton = screen.getByRole("button", { name: "移动幻灯片" });

      await user.clear(fromInput);
      await user.type(fromInput, "1");
      await user.clear(toInput);
      await user.type(toInput, "3");
      await user.click(moveButton);

      await waitFor(() => {
        expect(pptTools.moveSlide).toHaveBeenCalledWith({ fromIndex: 1, toIndex: 3 });
      });

      // 操作成功后会刷新列表，消息会变成"已加载 X 张幻灯片信息"
      await waitFor(() => {
        expect(screen.getByText(/已加载.*张幻灯片信息/)).toBeInTheDocument();
      });
    });

    it("应该显示错误信息当输入无效", async () => {
      const user = userEvent.setup();
      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText("源位置:")).toBeInTheDocument();
      });

      // 输入无效值
      const fromInput = screen.getByLabelText("源位置:") as HTMLInputElement;
      const toInput = screen.getByLabelText("目标位置:") as HTMLInputElement;
      await user.type(fromInput, "0");
      await user.type(toInput, "1");

      const moveButton = screen.getByRole("button", { name: "移动幻灯片" });
      await user.click(moveButton);

      await waitFor(() => {
        expect(screen.getByText(/请输入有效的源位置/)).toBeInTheDocument();
      });
    });

    it("应该禁用按钮当输入为空", async () => {
      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByRole("button", { name: "移动幻灯片" })).toBeDisabled();
      });
    });
  });

  describe("方法2: 移动当前幻灯片", () => {
    it("应该成功移动当前幻灯片", async () => {
      const user = userEvent.setup();
      vi.mocked(pptTools.moveCurrentSlide).mockResolvedValue({
        success: true,
        message: "成功将当前幻灯片从位置 1 移动到位置 4",
        fromIndex: 1,
        toIndex: 4,
        totalSlides: 5,
      });

      renderWithProviders(<SlideMove />);

      // 等待组件加载完成并设置 currentSlideIndex
      await waitFor(() => {
        expect(screen.getByLabelText("移动到位置:")).toBeInTheDocument();
      });

      // 等待 currentSlideIndex 被设置
      await waitFor(() => {
        expect(screen.getByText(/当前选中:/)).toBeInTheDocument();
      });

      const toInput = screen.getByLabelText("移动到位置:") as HTMLInputElement;
      const moveButton = screen.getByRole("button", { name: "移动当前幻灯片" });

      await user.clear(toInput);
      await user.type(toInput, "4");
      await user.click(moveButton);

      await waitFor(() => {
        expect(pptTools.moveCurrentSlide).toHaveBeenCalledWith(4);
      });

      // 操作成功后会刷新列表，消息会变成"已加载 X 张幻灯片信息"
      await waitFor(() => {
        expect(screen.getByText(/已加载.*张幻灯片信息/)).toBeInTheDocument();
      });
    });

    it("应该显示快速移动按钮", async () => {
      renderWithProviders(<SlideMove />);

      // 等待 currentSlideIndex 被设置（大于0时才显示快速移动按钮）
      await waitFor(() => {
        expect(screen.getByText(/当前选中:/)).toBeInTheDocument();
      });

      await waitFor(() => {
        expect(screen.getByText("快速移动到:")).toBeInTheDocument();
      });

      expect(screen.getByRole("button", { name: "开头" })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: "末尾" })).toBeInTheDocument();
    });

    it("应该通过快速按钮设置目标位置", async () => {
      const user = userEvent.setup();
      renderWithProviders(<SlideMove />);

      // 等待 currentSlideIndex 被设置
      await waitFor(() => {
        expect(screen.getByText(/当前选中:/)).toBeInTheDocument();
      });

      await waitFor(() => {
        expect(screen.getByRole("button", { name: "末尾" })).toBeInTheDocument();
      });

      const endButton = screen.getByRole("button", { name: "末尾" });
      await user.click(endButton);

      const toInput = screen.getByLabelText("移动到位置:") as HTMLInputElement;
      await waitFor(() => {
        expect(toInput.value).toBe("5");
      });
    });
  });

  describe("方法3: 交换幻灯片", () => {
    it("应该成功交换两张幻灯片", async () => {
      const user = userEvent.setup();
      vi.mocked(pptTools.swapSlides).mockResolvedValue({
        success: true,
        message: "成功交换位置 2 和位置 4 的幻灯片",
        fromIndex: 2,
        toIndex: 4,
        totalSlides: 5,
      });

      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText("第一张位置:")).toBeInTheDocument();
      });

      const index1Input = screen.getByLabelText("第一张位置:") as HTMLInputElement;
      const index2Input = screen.getByLabelText("第二张位置:") as HTMLInputElement;
      const swapButton = screen.getByRole("button", { name: "交换幻灯片" });

      await user.clear(index1Input);
      await user.type(index1Input, "2");
      await user.clear(index2Input);
      await user.type(index2Input, "4");
      await user.click(swapButton);

      await waitFor(() => {
        expect(pptTools.swapSlides).toHaveBeenCalledWith(2, 4);
      });

      // 操作成功后会刷新列表，消息会变成"已加载 X 张幻灯片信息"
      await waitFor(() => {
        expect(screen.getByText(/已加载.*张幻灯片信息/)).toBeInTheDocument();
      });
    });

    it("应该显示错误信息当交换失败", async () => {
      const user = userEvent.setup();
      vi.mocked(pptTools.swapSlides).mockResolvedValue({
        success: false,
        message: "两张幻灯片索引相同，无需交换",
        fromIndex: 2,
        toIndex: 2,
      });

      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText("第一张位置:")).toBeInTheDocument();
      });

      const index1Input = screen.getByLabelText("第一张位置:") as HTMLInputElement;
      const index2Input = screen.getByLabelText("第二张位置:") as HTMLInputElement;
      const swapButton = screen.getByRole("button", { name: "交换幻灯片" });

      await user.clear(index1Input);
      await user.type(index1Input, "2");
      await user.clear(index2Input);
      await user.type(index2Input, "2");
      await user.click(swapButton);

      await waitFor(() => {
        expect(screen.getByText(/交换失败/)).toBeInTheDocument();
      });
    });
  });

  describe("刷新功能", () => {
    it("应该刷新幻灯片列表", async () => {
      const user = userEvent.setup();
      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByRole("button", { name: /刷新列表/ })).toBeInTheDocument();
      });

      const refreshButton = screen.getByRole("button", { name: /刷新列表/ });
      await user.click(refreshButton);

      await waitFor(() => {
        expect(pptTools.getAllSlidesInfo).toHaveBeenCalled();
      });
    });

    it("应该在操作成功后自动刷新", async () => {
      const user = userEvent.setup();
      vi.mocked(pptTools.moveSlide).mockResolvedValue({
        success: true,
        message: "成功移动",
        fromIndex: 1,
        toIndex: 3,
        totalSlides: 5,
      });

      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText("源位置:")).toBeInTheDocument();
      });

      const fromInput = screen.getByLabelText("源位置:") as HTMLInputElement;
      const toInput = screen.getByLabelText("目标位置:") as HTMLInputElement;
      const moveButton = screen.getByRole("button", { name: "移动幻灯片" });

      await user.clear(fromInput);
      await user.type(fromInput, "1");
      await user.clear(toInput);
      await user.type(toInput, "3");
      await user.click(moveButton);

      await waitFor(() => {
        // 初始加载 + 操作后刷新 = 至少2次调用
        expect(pptTools.getAllSlidesInfo).toHaveBeenCalledTimes(2);
      });
    });
  });

  describe("错误处理", () => {
    it("应该处理 API 错误", async () => {
      const user = userEvent.setup();
      vi.mocked(pptTools.moveSlide).mockRejectedValue(new Error("API Error"));

      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText("源位置:")).toBeInTheDocument();
      });

      const fromInput = screen.getByLabelText("源位置:") as HTMLInputElement;
      const toInput = screen.getByLabelText("目标位置:") as HTMLInputElement;
      const moveButton = screen.getByRole("button", { name: "移动幻灯片" });

      await user.clear(fromInput);
      await user.type(fromInput, "1");
      await user.clear(toInput);
      await user.type(toInput, "3");
      await user.click(moveButton);

      await waitFor(() => {
        expect(screen.getByText(/移动失败.*API Error/)).toBeInTheDocument();
      });
    });

    it("应该处理加载幻灯片信息失败", async () => {
      vi.mocked(pptTools.getAllSlidesInfo).mockRejectedValue(new Error("Load Error"));

      renderWithProviders(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByText(/加载幻灯片信息失败/)).toBeInTheDocument();
      });
    });
  });
});
