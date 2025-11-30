/**
 * 文件名: ImageReplace.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: ImageReplace 组件的单元测试 | ImageReplace component unit tests
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import ImageReplace from "../../../src/taskpane/components/tools/ImageReplace";
import * as imageReplaceModule from "../../../src/ppt-tools/imageReplace";
import * as imageInsertionModule from "../../../src/ppt-tools/imageInsertion";

// Mock the ppt-tools modules
vi.mock("../../../src/ppt-tools/imageReplace");
vi.mock("../../../src/ppt-tools/imageInsertion");

describe("ImageReplace Component", () => {
  const mockImageInfo: imageReplaceModule.ImageElementInfo = {
    elementId: "shape-1",
    elementType: "Picture",
    name: "Test Image",
    left: 100,
    top: 100,
    width: 200,
    height: 150,
    isPlaceholder: false,
  };

  const mockPlaceholderInfo: imageReplaceModule.ImageElementInfo = {
    elementId: "shape-2",
    elementType: "Placeholder",
    name: "Test Placeholder",
    left: 100,
    top: 100,
    width: 300,
    height: 200,
    isPlaceholder: true,
    placeholderType: "Picture",
  };

  beforeEach(() => {
    vi.clearAllMocks();

    // Mock getImageInfo to return a picture by default
    vi.spyOn(imageReplaceModule, "getImageInfo").mockResolvedValue(mockImageInfo);

    // Mock replaceImage to return success
    vi.spyOn(imageReplaceModule, "replaceImage").mockResolvedValue({
      success: true,
      message: "图片替换成功",
      elementId: "new-shape-1",
      elementType: "Picture",
      originalDimensions: {
        left: 100,
        top: 100,
        width: 200,
        height: 150,
      },
    });

    // Mock readImageAsBase64
    vi.spyOn(imageInsertionModule, "readImageAsBase64").mockResolvedValue(
      "data:image/png;base64,iVBORw0KGgoAAAANS..."
    );

    // Mock fetchImageAsBase64
    vi.spyOn(imageInsertionModule, "fetchImageAsBase64").mockResolvedValue(
      "data:image/png;base64,iVBORw0KGgoAAAANS..."
    );

    // Mock alert
    vi.spyOn(window, "alert").mockImplementation(() => {});
  });

  it("应该渲染组件", () => {
    render(<ImageReplace />);
    expect(screen.getByText("选择新图片来源")).toBeInTheDocument();
  });

  it("应该显示当前选中的图片信息", async () => {
    render(<ImageReplace />);

    await waitFor(() => {
      expect(screen.getByText("当前图片信息")).toBeInTheDocument();
      expect(screen.getByText("Test Image")).toBeInTheDocument();
      expect(screen.getByText("Picture")).toBeInTheDocument();
    });
  });

  it("应该显示占位符图片的信息", async () => {
    vi.spyOn(imageReplaceModule, "getImageInfo").mockResolvedValue(mockPlaceholderInfo);

    render(<ImageReplace />);

    await waitFor(() => {
      expect(screen.getByText("Test Placeholder")).toBeInTheDocument();
      expect(screen.getByText(/Placeholder.*Picture/)).toBeInTheDocument();
    });
  });

  it("应该显示未选中元素的提示", async () => {
    vi.spyOn(imageReplaceModule, "getImageInfo").mockResolvedValue(null);

    render(<ImageReplace />);

    await waitFor(() => {
      expect(screen.getByText("使用说明")).toBeInTheDocument();
      expect(screen.getByText(/请先在幻灯片中选中要替换的图片/)).toBeInTheDocument();
    });
  });

  it("应该显示选中非图片元素的警告", async () => {
    vi.spyOn(imageReplaceModule, "getImageInfo").mockResolvedValue({
      ...mockImageInfo,
      elementType: "TextBox",
    });

    render(<ImageReplace />);

    await waitFor(() => {
      expect(screen.getByText("选中的元素不是图片")).toBeInTheDocument();
    });
  });

  it("应该支持切换图片来源类型", async () => {
    const user = userEvent.setup();
    render(<ImageReplace />);

    // 默认是上传本地图片
    expect(screen.getByLabelText("上传本地图片（推荐）")).toBeChecked();

    // 切换到 URL
    await user.click(screen.getByLabelText("使用图片 URL"));
    expect(screen.getByLabelText("图片 URL")).toBeInTheDocument();

    // 切换到 Base64
    await user.click(screen.getByLabelText("粘贴 Base64 数据"));
    expect(screen.getByLabelText("Base64 数据")).toBeInTheDocument();
  });

  it("应该支持切换尺寸设置", async () => {
    const user = userEvent.setup();
    render(<ImageReplace />);

    await waitFor(() => {
      expect(screen.getByText("当前图片信息")).toBeInTheDocument();
    });

    // 默认保持原图片尺寸
    const switchElement = screen.getByRole("switch");
    expect(switchElement).toBeChecked();

    // 切换到自定义尺寸
    await user.click(switchElement);
    expect(switchElement).not.toBeChecked();

    // 应该显示宽度和高度输入框
    expect(screen.getByLabelText("宽度（磅）")).toBeInTheDocument();
    expect(screen.getByLabelText("高度（磅）")).toBeInTheDocument();
  });

  it("应该在未选择图片时禁用替换按钮", async () => {
    render(<ImageReplace />);

    await waitFor(() => {
      expect(screen.getByText("当前图片信息")).toBeInTheDocument();
    });

    const replaceButton = screen.getByRole("button", { name: /确认替换/ });
    expect(replaceButton).toBeDisabled();
  });

  it("应该在选中非图片元素时禁用替换按钮", async () => {
    vi.spyOn(imageReplaceModule, "getImageInfo").mockResolvedValue({
      ...mockImageInfo,
      elementType: "TextBox",
    });

    render(<ImageReplace />);

    await waitFor(() => {
      expect(screen.getByText("选中的元素不是图片")).toBeInTheDocument();
    });

    const replaceButton = screen.getByRole("button", { name: /确认替换/ });
    expect(replaceButton).toBeDisabled();
  });

  it("应该支持刷新选中元素", async () => {
    const user = userEvent.setup();
    const getImageInfoSpy = vi.spyOn(imageReplaceModule, "getImageInfo");

    render(<ImageReplace />);

    await waitFor(() => {
      expect(getImageInfoSpy).toHaveBeenCalledTimes(1);
    });

    // 点击刷新按钮
    const refreshButton = screen.getByRole("button", { name: /刷新选中元素/ });
    await user.click(refreshButton);

    await waitFor(() => {
      expect(getImageInfoSpy).toHaveBeenCalledTimes(2);
    });
  });

  describe("本地上传模式", () => {
    it("应该处理文件上传", async () => {
      const user = userEvent.setup();
      render(<ImageReplace />);

      await waitFor(() => {
        expect(screen.getByText("当前图片信息")).toBeInTheDocument();
      });

      // 创建一个模拟的文件
      const file = new File(["dummy content"], "test.png", { type: "image/png" });

      // 找到隐藏的文件输入框
      const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      expect(fileInput).toBeInTheDocument();

      // 上传文件
      await user.upload(fileInput, file);

      await waitFor(() => {
        expect(imageInsertionModule.readImageAsBase64).toHaveBeenCalledWith(file);
      });
    });

    it("应该在选择图片后启用替换按钮", async () => {
      const user = userEvent.setup();
      render(<ImageReplace />);

      await waitFor(() => {
        expect(screen.getByText("当前图片信息")).toBeInTheDocument();
      });

      const file = new File(["dummy content"], "test.png", { type: "image/png" });
      const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;

      await user.upload(fileInput, file);

      await waitFor(() => {
        const replaceButton = screen.getByRole("button", { name: /确认替换/ });
        expect(replaceButton).not.toBeDisabled();
      });
    });
  });

  describe("URL 模式", () => {
    it("应该在输入 URL 后启用替换按钮", async () => {
      const user = userEvent.setup();
      render(<ImageReplace />);

      await waitFor(() => {
        expect(screen.getByText("当前图片信息")).toBeInTheDocument();
      });

      // 切换到 URL 模式
      await user.click(screen.getByLabelText("使用图片 URL"));

      // 输入 URL
      const urlInput = screen.getByLabelText("图片 URL");
      await user.type(urlInput, "https://example.com/image.png");

      await waitFor(() => {
        const replaceButton = screen.getByRole("button", { name: /确认替换/ });
        expect(replaceButton).not.toBeDisabled();
      });
    });
  });

  describe("Base64 模式", () => {
    it("应该在输入 Base64 后启用替换按钮", async () => {
      const user = userEvent.setup();
      render(<ImageReplace />);

      await waitFor(() => {
        expect(screen.getByText("当前图片信息")).toBeInTheDocument();
      });

      // 切换到 Base64 模式
      await user.click(screen.getByLabelText("粘贴 Base64 数据"));

      // 输入 Base64
      const base64Input = screen.getByLabelText("Base64 数据");
      await user.type(base64Input, "iVBORw0KGgoAAAANS...");

      await waitFor(() => {
        const replaceButton = screen.getByRole("button", { name: /确认替换/ });
        expect(replaceButton).not.toBeDisabled();
      });
    });
  });

  describe("替换操作", () => {
    it("应该成功替换图片", async () => {
      const user = userEvent.setup();
      render(<ImageReplace />);

      await waitFor(() => {
        expect(screen.getByText("当前图片信息")).toBeInTheDocument();
      });

      // 上传文件
      const file = new File(["dummy content"], "test.png", { type: "image/png" });
      const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      await user.upload(fileInput, file);

      await waitFor(() => {
        const replaceButton = screen.getByRole("button", { name: /确认替换/ });
        expect(replaceButton).not.toBeDisabled();
      });

      // 点击替换按钮
      const replaceButton = screen.getByRole("button", { name: /确认替换/ });
      await user.click(replaceButton);

      await waitFor(() => {
        expect(imageReplaceModule.replaceImage).toHaveBeenCalledWith({
          imageSource: "data:image/png;base64,iVBORw0KGgoAAAANS...",
          keepDimensions: true,
          width: undefined,
          height: undefined,
        });
        expect(window.alert).toHaveBeenCalledWith(expect.stringContaining("图片替换成功"));
      });
    });

    it("应该支持自定义尺寸替换", async () => {
      const user = userEvent.setup();
      render(<ImageReplace />);

      await waitFor(() => {
        expect(screen.getByText("当前图片信息")).toBeInTheDocument();
      });

      // 上传文件
      const file = new File(["dummy content"], "test.png", { type: "image/png" });
      const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      await user.upload(fileInput, file);

      // 切换到自定义尺寸
      const switchElement = screen.getByRole("switch");
      await user.click(switchElement);

      // 输入自定义尺寸
      const widthInput = screen.getByLabelText("宽度（磅）");
      const heightInput = screen.getByLabelText("高度（磅）");
      await user.clear(widthInput);
      await user.type(widthInput, "300");
      await user.clear(heightInput);
      await user.type(heightInput, "250");

      // 点击替换按钮
      const replaceButton = screen.getByRole("button", { name: /确认替换/ });
      await user.click(replaceButton);

      await waitFor(() => {
        expect(imageReplaceModule.replaceImage).toHaveBeenCalledWith({
          imageSource: "data:image/png;base64,iVBORw0KGgoAAAANS...",
          keepDimensions: false,
          width: 300,
          height: 250,
        });
      });
    });

    it("应该处理替换失败的情况", async () => {
      const user = userEvent.setup();

      // Mock replaceImage to return failure
      vi.spyOn(imageReplaceModule, "replaceImage").mockResolvedValue({
        success: false,
        message: "替换失败：未找到元素",
      });

      render(<ImageReplace />);

      await waitFor(() => {
        expect(screen.getByText("当前图片信息")).toBeInTheDocument();
      });

      // 上传文件
      const file = new File(["dummy content"], "test.png", { type: "image/png" });
      const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      await user.upload(fileInput, file);

      // 点击替换按钮
      const replaceButton = screen.getByRole("button", { name: /确认替换/ });
      await user.click(replaceButton);

      await waitFor(() => {
        expect(window.alert).toHaveBeenCalledWith(expect.stringContaining("替换失败"));
      });
    });
  });
});
