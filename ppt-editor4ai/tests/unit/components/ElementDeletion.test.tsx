/**
 * 文件名: ElementDeletion.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: vitest, @testing-library/react
 * 描述: ElementDeletion 组件单元测试 | ElementDeletion component unit tests
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { screen, waitFor } from "@testing-library/react";
import { renderWithProviders, userEvent } from "../../utils/test-utils";
import { ElementDeletion } from "../../../src/taskpane/components/tools/ElementDeletion";
import * as pptTools from "../../../src/ppt-tools";

// Mock ppt-tools module
vi.mock("../../../src/ppt-tools", () => ({
  getCurrentSlideElements: vi.fn(),
  deleteElementById: vi.fn(),
  deleteElementsByIds: vi.fn(),
}));

// Mock PowerPoint API
const mockPowerPoint = {
  run: vi.fn(),
};

global.PowerPoint = mockPowerPoint as any;

describe("ElementDeletion 组件单元测试 | ElementDeletion Component Unit Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("应该正确渲染组件 | should render component correctly", () => {
    renderWithProviders(<ElementDeletion />);

    // 验证标题存在 | Verify title exists
    expect(screen.getByText("元素删除调试工具")).toBeInTheDocument();

    // 验证按钮存在 | Verify buttons exist
    expect(screen.getByRole("button", { name: "获取当前页面元素" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "获取PPT中选中的元素" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "删除选中元素" })).toBeInTheDocument();

    // 验证输入框存在 | Verify textarea exists
    expect(screen.getByPlaceholderText(/输入或从列表中选择元素ID/)).toBeInTheDocument();
  });

  it("应该能够获取当前页面元素列表 | should be able to get current slide elements", async () => {
    const user = userEvent.setup();
    const mockElements = [
      {
        id: "shape-123",
        type: "TextBox",
        name: "Text 1",
        left: 100,
        top: 200,
        width: 300,
        height: 100,
        text: "Sample text",
      },
      {
        id: "shape-456",
        type: "Rectangle",
        name: "Shape 1",
        left: 150,
        top: 250,
        width: 200,
        height: 150,
      },
    ];

    vi.mocked(pptTools.getCurrentSlideElements).mockResolvedValue(mockElements);

    renderWithProviders(<ElementDeletion />);

    const button = screen.getByRole("button", { name: "获取当前页面元素" });
    await user.click(button);

    await waitFor(() => {
      expect(pptTools.getCurrentSlideElements).toHaveBeenCalledTimes(1);
      expect(screen.getByText(/成功获取 2 个元素/)).toBeInTheDocument();
    });

    // 验证元素列表显示 | Verify elements list is displayed
    expect(screen.getByText("TextBox - Text 1")).toBeInTheDocument();
    expect(screen.getByText("Rectangle - Shape 1")).toBeInTheDocument();
  });

  it("获取元素列表失败时应该显示错误信息 | should display error message when getting elements fails", async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.getCurrentSlideElements).mockRejectedValue(new Error("获取失败"));

    renderWithProviders(<ElementDeletion />);

    const button = screen.getByRole("button", { name: "获取当前页面元素" });
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText(/获取元素列表失败: 获取失败/)).toBeInTheDocument();
    });
  });

  it("应该能够从列表中选择元素 | should be able to select element from list", async () => {
    const user = userEvent.setup();
    const mockElements = [
      {
        id: "shape-123",
        type: "TextBox",
        name: "Text 1",
        left: 100,
        top: 200,
        width: 300,
        height: 100,
      },
    ];

    vi.mocked(pptTools.getCurrentSlideElements).mockResolvedValue(mockElements);

    renderWithProviders(<ElementDeletion />);

    // 先获取元素列表 | First get elements list
    const getButton = screen.getByRole("button", { name: "获取当前页面元素" });
    await user.click(getButton);

    await waitFor(() => {
      expect(screen.getByText("TextBox - Text 1")).toBeInTheDocument();
    });

    // 点击元素选择 | Click to select element
    const elementItem = screen.getByText("TextBox - Text 1");
    await user.click(elementItem);

    await waitFor(() => {
      expect(screen.getByText(/已选中元素: shape-123/)).toBeInTheDocument();
    });

    // 验证输入框中有ID | Verify ID in textarea
    const textarea = screen.getByPlaceholderText(/输入或从列表中选择元素ID/) as HTMLTextAreaElement;
    expect(textarea.value).toBe("shape-123");
  });

  it("应该支持多选元素（Ctrl+点击）| should support multi-select elements (Ctrl+click)", async () => {
    const user = userEvent.setup();
    const mockElements = [
      {
        id: "shape-123",
        type: "TextBox",
        name: "Text 1",
        left: 100,
        top: 200,
        width: 300,
        height: 100,
      },
      {
        id: "shape-456",
        type: "Rectangle",
        name: "Shape 1",
        left: 150,
        top: 250,
        width: 200,
        height: 150,
      },
    ];

    vi.mocked(pptTools.getCurrentSlideElements).mockResolvedValue(mockElements);

    renderWithProviders(<ElementDeletion />);

    // 获取元素列表 | Get elements list
    const getButton = screen.getByRole("button", { name: "获取当前页面元素" });
    await user.click(getButton);

    await waitFor(() => {
      expect(screen.getByText("TextBox - Text 1")).toBeInTheDocument();
    });

    // 第一次点击选择第一个元素 | First click to select first element
    const element1 = screen.getByText("TextBox - Text 1");
    await user.click(element1);

    // Ctrl+点击选择第二个元素 | Ctrl+click to select second element
    const element2 = screen.getByText("Rectangle - Shape 1");
    await user.keyboard("{Control>}");
    await user.click(element2);
    await user.keyboard("{/Control}");

    await waitFor(() => {
      expect(screen.getByText(/已选中 2 个元素/)).toBeInTheDocument();
    });

    // 验证输入框中有两个ID | Verify two IDs in textarea
    const textarea = screen.getByPlaceholderText(/输入或从列表中选择元素ID/) as HTMLTextAreaElement;
    expect(textarea.value).toContain("shape-123");
    expect(textarea.value).toContain("shape-456");
  });

  it("应该能够手动输入元素ID | should be able to manually input element ID", async () => {
    const user = userEvent.setup();
    renderWithProviders(<ElementDeletion />);

    const textarea = screen.getByPlaceholderText(/输入或从列表中选择元素ID/);
    await user.type(textarea, "manual-id-123");

    expect((textarea as HTMLTextAreaElement).value).toBe("manual-id-123");
  });

  it("应该能够删除单个元素 | should be able to delete single element", async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.deleteElementById).mockResolvedValue({
      success: true,
      deletedCount: 1,
      message: "成功删除 1 个元素",
    });

    vi.mocked(pptTools.getCurrentSlideElements).mockResolvedValue([]);

    renderWithProviders(<ElementDeletion />);

    // 输入元素ID | Input element ID
    const textarea = screen.getByPlaceholderText(/输入或从列表中选择元素ID/);
    await user.type(textarea, "shape-123");

    // 点击删除按钮 | Click delete button
    const deleteButton = screen.getByRole("button", { name: "删除选中元素" });
    await user.click(deleteButton);

    // 等待删除函数被调用 | Wait for delete function to be called
    await waitFor(() => {
      expect(pptTools.deleteElementById).toHaveBeenCalledWith("shape-123");
    });

    // 等待刷新列表函数被调用 | Wait for refresh list function to be called
    await waitFor(() => {
      expect(pptTools.getCurrentSlideElements).toHaveBeenCalled();
    });

    // 验证输入框被清空 | Verify textarea is cleared
    expect((textarea as HTMLTextAreaElement).value).toBe("");
  });

  it("应该能够批量删除多个元素 | should be able to delete multiple elements", async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.deleteElementsByIds).mockResolvedValue({
      success: true,
      deletedCount: 2,
      message: "成功删除 2 个元素",
    });

    vi.mocked(pptTools.getCurrentSlideElements).mockResolvedValue([]);

    renderWithProviders(<ElementDeletion />);

    // 输入多个元素ID（逗号分隔）| Input multiple element IDs (comma separated)
    const textarea = screen.getByPlaceholderText(/输入或从列表中选择元素ID/);
    await user.type(textarea, "shape-123, shape-456");

    // 点击删除按钮 | Click delete button
    const deleteButton = screen.getByRole("button", { name: "删除选中元素" });
    await user.click(deleteButton);

    // 等待删除函数被调用 | Wait for delete function to be called
    await waitFor(() => {
      expect(pptTools.deleteElementsByIds).toHaveBeenCalledWith(["shape-123", "shape-456"]);
    });

    // 等待刷新列表函数被调用 | Wait for refresh list function to be called
    await waitFor(() => {
      expect(pptTools.getCurrentSlideElements).toHaveBeenCalled();
    });

    // 验证输入框被清空 | Verify textarea is cleared
    expect((textarea as HTMLTextAreaElement).value).toBe("");
  });

  it("删除失败时应该显示错误信息 | should display error message when deletion fails", async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.deleteElementById).mockResolvedValue({
      success: false,
      deletedCount: 0,
      message: "未找到ID为 shape-999 的元素",
    });

    renderWithProviders(<ElementDeletion />);

    const textarea = screen.getByPlaceholderText(/输入或从列表中选择元素ID/);
    await user.type(textarea, "shape-999");

    const deleteButton = screen.getByRole("button", { name: "删除选中元素" });
    await user.click(deleteButton);

    await waitFor(() => {
      expect(screen.getByText(/删除失败: 未找到ID为 shape-999 的元素/)).toBeInTheDocument();
    });
  });

  it("未输入ID时删除按钮应该被禁用 | delete button should be disabled when no ID is entered", () => {
    renderWithProviders(<ElementDeletion />);

    const deleteButton = screen.getByRole("button", { name: "删除选中元素" });
    // 验证按钮被禁用，因此无法点击 | Verify button is disabled, so cannot click
    expect(deleteButton).toBeDisabled();
  });

  it("应该能够获取PPT中选中的元素 | should be able to get selected shapes from PPT", async () => {
    const user = userEvent.setup();

    const mockShape1 = {
      id: "selected-123",
      type: "TextBox",
      name: "Selected Text",
      load: vi.fn(),
    };

    const mockShapes = {
      items: [mockShape1],
      load: vi.fn(),
    };

    const mockShapeCount = {
      value: 1,
    };

    const mockContext = {
      presentation: {
        getSelectedShapes: vi.fn().mockReturnValue({
          getCount: vi.fn().mockReturnValue(mockShapeCount),
          items: [mockShape1],
          load: vi.fn(),
        }),
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    mockPowerPoint.run.mockImplementation(async (callback: any) => {
      const shapes = mockContext.presentation.getSelectedShapes();
      shapes.items = [mockShape1];
      await callback(mockContext);
    });

    renderWithProviders(<ElementDeletion />);

    const button = screen.getByRole("button", { name: "获取PPT中选中的元素" });
    await user.click(button);

    await waitFor(() => {
      expect(mockPowerPoint.run).toHaveBeenCalled();
    });
  });

  it("PPT中未选中元素时应该显示提示 | should show prompt when no shapes selected in PPT", async () => {
    const user = userEvent.setup();

    const mockShapeCount = {
      value: 0,
    };

    const mockContext = {
      presentation: {
        getSelectedShapes: vi.fn().mockReturnValue({
          getCount: vi.fn().mockReturnValue(mockShapeCount),
          items: [],
          load: vi.fn(),
        }),
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    mockPowerPoint.run.mockImplementation(async (callback: any) => {
      await callback(mockContext);
    });

    renderWithProviders(<ElementDeletion />);

    const button = screen.getByRole("button", { name: "获取PPT中选中的元素" });
    await user.click(button);

    await waitFor(() => {
      expect(screen.getByText("请先在幻灯片中选中至少一个元素")).toBeInTheDocument();
    });
  });

  it("应该显示使用说明 | should display usage instructions", () => {
    renderWithProviders(<ElementDeletion />);

    expect(screen.getByText("使用说明:")).toBeInTheDocument();
    expect(screen.getByText(/方式1: 点击.*获取当前页面元素.*按钮/)).toBeInTheDocument();
    expect(screen.getByText(/方式2: 在PPT中选中元素/)).toBeInTheDocument();
    expect(screen.getByText(/方式3: 手动输入元素ID/)).toBeInTheDocument();
  });

  it("删除按钮在没有ID时应该被禁用 | delete button should be disabled when no ID", () => {
    renderWithProviders(<ElementDeletion />);

    const deleteButton = screen.getByRole("button", { name: "删除选中元素" });
    expect(deleteButton).toBeDisabled();
  });

  it("删除按钮在有ID时应该被启用 | delete button should be enabled when ID exists", async () => {
    const user = userEvent.setup();
    renderWithProviders(<ElementDeletion />);

    const textarea = screen.getByPlaceholderText(/输入或从列表中选择元素ID/);
    await user.type(textarea, "shape-123");

    const deleteButton = screen.getByRole("button", { name: "删除选中元素" });
    expect(deleteButton).not.toBeDisabled();
  });

  it("应该支持多种ID分隔符（逗号、空格、换行）| should support multiple ID separators (comma, space, newline)", async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.deleteElementsByIds).mockResolvedValue({
      success: true,
      deletedCount: 3,
      message: "成功删除 3 个元素",
    });

    vi.mocked(pptTools.getCurrentSlideElements).mockResolvedValue([]);

    renderWithProviders(<ElementDeletion />);

    // 输入多种分隔符的ID | Input IDs with various separators
    const textarea = screen.getByPlaceholderText(/输入或从列表中选择元素ID/);
    await user.type(textarea, "id1, id2 id3");

    const deleteButton = screen.getByRole("button", { name: "删除选中元素" });
    await user.click(deleteButton);

    await waitFor(() => {
      expect(pptTools.deleteElementsByIds).toHaveBeenCalledWith(["id1", "id2", "id3"]);
    });
  });

  it("删除成功后应该清空输入框并刷新列表 | should clear input and refresh list after successful deletion", async () => {
    const user = userEvent.setup();
    vi.mocked(pptTools.deleteElementById).mockResolvedValue({
      success: true,
      deletedCount: 1,
      message: "成功删除 1 个元素",
    });

    vi.mocked(pptTools.getCurrentSlideElements).mockResolvedValue([]);

    renderWithProviders(<ElementDeletion />);

    const textarea = screen.getByPlaceholderText(/输入或从列表中选择元素ID/) as HTMLTextAreaElement;
    await user.type(textarea, "shape-123");

    const deleteButton = screen.getByRole("button", { name: "删除选中元素" });
    await user.click(deleteButton);

    await waitFor(() => {
      expect(textarea.value).toBe("");
      expect(pptTools.getCurrentSlideElements).toHaveBeenCalled();
    });
  });

  it("应该显示元素的详细信息 | should display element details", async () => {
    const user = userEvent.setup();
    const mockElements = [
      {
        id: "shape-123",
        type: "TextBox",
        name: "Text 1",
        left: 100.5,
        top: 200.7,
        width: 300.3,
        height: 100.9,
        text: "Sample text content",
        placeholderType: "Title",
      },
    ];

    vi.mocked(pptTools.getCurrentSlideElements).mockResolvedValue(mockElements);

    renderWithProviders(<ElementDeletion />);

    const button = screen.getByRole("button", { name: "获取当前页面元素" });
    await user.click(button);

    await waitFor(() => {
      // 验证元素类型和名称 | Verify element type and name
      expect(screen.getByText("TextBox - Text 1")).toBeInTheDocument();

      // 验证ID | Verify ID
      expect(screen.getByText(/ID: shape-123/)).toBeInTheDocument();

      // 验证位置和尺寸（四舍五入）| Verify position and size (rounded)
      expect(screen.getByText(/位置: \(101, 201\)/)).toBeInTheDocument();
      expect(screen.getByText(/尺寸: 300 × 101/)).toBeInTheDocument();

      // 验证文本内容 | Verify text content
      expect(screen.getByText(/文本: Sample text content/)).toBeInTheDocument();

      // 验证占位符类型 | Verify placeholder type
      expect(screen.getByText(/占位符类型: Title/)).toBeInTheDocument();
    });
  });
});
