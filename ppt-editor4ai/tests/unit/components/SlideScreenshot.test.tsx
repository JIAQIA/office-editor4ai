/**
 * 文件名: SlideScreenshot.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: vitest, @testing-library/react
 * 描述: SlideScreenshot 组件单元测试 | SlideScreenshot component unit tests
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { screen, waitFor, cleanup } from "@testing-library/react";
import { renderWithProviders, userEvent } from "../../utils/test-utils";
import SlideScreenshot from "../../../src/taskpane/components/tools/SlideScreenshot";
import * as pptTools from "../../../src/ppt-tools";

// Mock ppt-tools module
vi.mock("../../../src/ppt-tools", () => ({
  getSlideScreenshot: vi.fn().mockResolvedValue({
    imageBase64: "mockBase64ImageData",
    slideIndex: 0,
    slideId: "slide-123",
    width: undefined,
    height: undefined,
  }),
  getCurrentSlideScreenshot: vi.fn().mockResolvedValue({
    imageBase64: "mockBase64CurrentSlide",
    slideIndex: 2,
    slideId: "slide-current",
    width: undefined,
    height: undefined,
  }),
  getSlideScreenshotByPageNumber: vi.fn().mockResolvedValue({
    imageBase64: "mockBase64SpecificSlide",
    slideIndex: 4,
    slideId: "slide-specific",
    width: undefined,
    height: undefined,
  }),
  getAllSlidesScreenshots: vi.fn().mockResolvedValue([
    {
      imageBase64: "mockBase64Slide1",
      slideIndex: 0,
      slideId: "slide-1",
      width: undefined,
      height: undefined,
    },
    {
      imageBase64: "mockBase64Slide2",
      slideIndex: 1,
      slideId: "slide-2",
      width: undefined,
      height: undefined,
    },
    {
      imageBase64: "mockBase64Slide3",
      slideIndex: 2,
      slideId: "slide-3",
      width: undefined,
      height: undefined,
    },
  ]),
}));

// Mock clipboard API
Object.assign(navigator, {
  clipboard: {
    writeText: vi.fn().mockResolvedValue(undefined),
  },
});

describe("SlideScreenshot 组件单元测试 | SlideScreenshot Component Unit Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    // Mock window.alert
    global.alert = vi.fn();
    // Mock console methods
    global.console.log = vi.fn();
    global.console.error = vi.fn();
  });

  afterEach(() => {
    cleanup();
  });

  it("应该正确渲染组件 | should render component correctly", () => {
    renderWithProviders(<SlideScreenshot />);

    // 验证模式选择存在 | Verify mode selection exists
    expect(screen.getByText("选择截图模式")).toBeInTheDocument();
    expect(screen.getByLabelText("当前幻灯片")).toBeInTheDocument();
    expect(screen.getByLabelText("指定页码")).toBeInTheDocument();
    expect(screen.getByLabelText("所有幻灯片")).toBeInTheDocument();

    // 验证尺寸输入框存在 | Verify dimension inputs exist
    expect(screen.getByLabelText("宽度（像素）")).toBeInTheDocument();
    expect(screen.getByLabelText("高度（像素）")).toBeInTheDocument();

    // 验证截图按钮存在 | Verify capture button exists
    expect(screen.getByRole("button", { name: "开始截图" })).toBeInTheDocument();

    // 验证提示信息存在 | Verify hint text exists
    expect(screen.getByText(/提示: 如果不指定尺寸/)).toBeInTheDocument();
  });

  it('应该默认选择"当前幻灯片"模式 | should default to "current slide" mode', () => {
    renderWithProviders(<SlideScreenshot />);

    const currentRadio = screen.getByLabelText("当前幻灯片") as HTMLInputElement;
    expect(currentRadio.checked).toBe(true);

    // 不应该显示页码输入框 | Should not show page number input
    expect(screen.queryByLabelText("页码（从 1 开始）")).not.toBeInTheDocument();
  });

  it('应该能够切换到"指定页码"模式 | should be able to switch to "specific page" mode', async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideScreenshot />);

    const specificRadio = screen.getByLabelText("指定页码");
    await user.click(specificRadio);

    // 应该显示页码输入框 | Should show page number input
    expect(screen.getByLabelText("页码（从 1 开始）")).toBeInTheDocument();
  });

  it('应该能够切换到"所有幻灯片"模式 | should be able to switch to "all slides" mode', async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideScreenshot />);

    const allRadio = screen.getByLabelText("所有幻灯片");
    await user.click(allRadio);

    const allRadioChecked = screen.getByLabelText("所有幻灯片") as HTMLInputElement;
    expect(allRadioChecked.checked).toBe(true);
  });

  it("应该能够输入尺寸值 | should be able to input dimension values", async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideScreenshot />);

    const widthInput = screen.getByLabelText("宽度（像素）") as HTMLInputElement;
    const heightInput = screen.getByLabelText("高度（像素）") as HTMLInputElement;

    await user.type(widthInput, "800");
    await user.type(heightInput, "600");

    expect(widthInput.value).toBe("800");
    expect(heightInput.value).toBe("600");
  });

  it("应该能够输入页码 | should be able to input page number", async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideScreenshot />);

    // 切换到指定页码模式 | Switch to specific page mode
    const specificRadio = screen.getByLabelText("指定页码");
    await user.click(specificRadio);

    const pageInput = screen.getByLabelText("页码（从 1 开始）") as HTMLInputElement;
    await user.clear(pageInput);
    await user.type(pageInput, "5");

    expect(pageInput.value).toBe("5");
  });

  it("应该能够获取当前幻灯片截图 | should be able to capture current slide screenshot", async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideScreenshot />);

    const captureButton = screen.getByRole("button", { name: "开始截图" });
    await user.click(captureButton);

    await waitFor(() => {
      expect(pptTools.getCurrentSlideScreenshot).toHaveBeenCalledWith(undefined, undefined);
    });

    // 应该显示截图预览 | Should show screenshot preview
    await waitFor(() => {
      expect(screen.getByText("幻灯片截图")).toBeInTheDocument();
      expect(screen.getByText(/页码: 3/)).toBeInTheDocument(); // slideIndex 2 + 1
      expect(screen.getByText(/ID: slide-current/)).toBeInTheDocument();
    });
  });

  it("应该能够获取当前幻灯片截图（带尺寸）| should be able to capture current slide with dimensions", async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideScreenshot />);

    // 输入尺寸 | Input dimensions
    const widthInput = screen.getByLabelText("宽度（像素）");
    const heightInput = screen.getByLabelText("高度（像素）");
    await user.type(widthInput, "1024");
    await user.type(heightInput, "768");

    const captureButton = screen.getByRole("button", { name: "开始截图" });
    await user.click(captureButton);

    await waitFor(() => {
      expect(pptTools.getCurrentSlideScreenshot).toHaveBeenCalledWith(1024, 768);
    });
  });

  it("应该能够获取指定页码的截图 | should be able to capture specific page screenshot", async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideScreenshot />);

    // 切换到指定页码模式 | Switch to specific page mode
    const specificRadio = screen.getByLabelText("指定页码");
    await user.click(specificRadio);

    // 输入页码 | Input page number
    const pageInput = screen.getByLabelText("页码（从 1 开始）");
    await user.clear(pageInput);
    await user.type(pageInput, "3");

    const captureButton = screen.getByRole("button", { name: "开始截图" });
    await user.click(captureButton);

    await waitFor(() => {
      expect(pptTools.getSlideScreenshotByPageNumber).toHaveBeenCalledWith(3, undefined, undefined);
    });

    // 应该显示截图预览 | Should show screenshot preview
    await waitFor(() => {
      expect(screen.getByText(/页码: 5/)).toBeInTheDocument(); // slideIndex 4 + 1
      expect(screen.getByText(/ID: slide-specific/)).toBeInTheDocument();
    });
  });

  it("应该在输入无效页码时显示警告 | should show alert when entering invalid page number", async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideScreenshot />);

    // 切换到指定页码模式 | Switch to specific page mode
    const specificRadio = screen.getByLabelText("指定页码");
    await user.click(specificRadio);

    // 输入无效页码 | Input invalid page number
    const pageInput = screen.getByLabelText("页码（从 1 开始）");
    await user.clear(pageInput);
    await user.type(pageInput, "0");

    const captureButton = screen.getByRole("button", { name: "开始截图" });
    await user.click(captureButton);

    await waitFor(() => {
      expect(global.alert).toHaveBeenCalledWith("请输入有效的页码（从 1 开始）");
    });

    expect(pptTools.getSlideScreenshotByPageNumber).not.toHaveBeenCalled();
  });

  it("应该能够获取所有幻灯片截图 | should be able to capture all slides screenshots", async () => {
    const user = userEvent.setup();
    renderWithProviders(<SlideScreenshot />);

    // 切换到所有幻灯片模式 | Switch to all slides mode
    const allRadio = screen.getByLabelText("所有幻灯片");
    await user.click(allRadio);

    const captureButton = screen.getByRole("button", { name: "开始截图" });
    await user.click(captureButton);

    await waitFor(() => {
      expect(pptTools.getAllSlidesScreenshots).toHaveBeenCalledWith(undefined, undefined);
    });

    // 应该显示所有截图 | Should show all screenshots
    await waitFor(() => {
      expect(screen.getByText("共 3 张幻灯片")).toBeInTheDocument();
      expect(screen.getByText("幻灯片 1")).toBeInTheDocument();
      expect(screen.getByText("幻灯片 2")).toBeInTheDocument();
      expect(screen.getByText("幻灯片 3")).toBeInTheDocument();
    });
  });

  it("应该在截图过程中显示加载状态 | should show loading state during capture", async () => {
    const user = userEvent.setup();

    // 让截图操作延迟完成 | Make capture delayed
    (pptTools.getCurrentSlideScreenshot as any).mockImplementation(
      () =>
        new Promise((resolve) =>
          setTimeout(
            () =>
              resolve({
                imageBase64: "delayed",
                slideIndex: 0,
                slideId: "delayed-slide",
              }),
            100
          )
        )
    );

    renderWithProviders(<SlideScreenshot />);

    const captureButton = screen.getByRole("button", { name: "开始截图" });
    await user.click(captureButton);

    // 应该显示加载状态 | Should show loading state
    expect(screen.getByRole("button", { name: "截图中..." })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "截图中..." })).toBeDisabled();
    expect(screen.getByText("正在获取截图...")).toBeInTheDocument();
  });

  it("应该在获取所有幻灯片时显示特定加载文本 | should show specific loading text when capturing all slides", async () => {
    const user = userEvent.setup();

    // 让截图操作延迟完成 | Make capture delayed
    (pptTools.getAllSlidesScreenshots as any).mockImplementation(
      () => new Promise((resolve) => setTimeout(() => resolve([]), 100))
    );

    renderWithProviders(<SlideScreenshot />);

    // 切换到所有幻灯片模式 | Switch to all slides mode
    const allRadio = screen.getByLabelText("所有幻灯片");
    await user.click(allRadio);

    const captureButton = screen.getByRole("button", { name: "开始截图" });
    await user.click(captureButton);

    // 应该显示特定加载文本 | Should show specific loading text
    await waitFor(() => {
      expect(screen.getByText("正在获取所有幻灯片截图，请稍候...")).toBeInTheDocument();
    });
  });

});
