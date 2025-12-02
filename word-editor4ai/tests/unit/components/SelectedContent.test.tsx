/**
 * 文件名: SelectedContent.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: SelectedContent 组件的单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { render, screen, fireEvent, waitFor } from "@testing-library/react";
import SelectedContent from "../../../src/taskpane/components/tools/SelectedContent";
import * as wordTools from "../../../src/word-tools";

// Mock word-tools 模块 / Mock word-tools module
vi.mock("../../../src/word-tools", () => ({
  getSelectedContent: vi.fn(),
}));

describe("SelectedContent 组件 / SelectedContent Component", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("应该正确渲染初始状态 / Should render initial state correctly", () => {
    render(<SelectedContent />);

    expect(screen.getByText(/请先在文档中选中要获取的内容/)).toBeInTheDocument();
    expect(screen.getByText("获取选项")).toBeInTheDocument();
    expect(screen.getByText("包含文本内容")).toBeInTheDocument();
    expect(screen.getByText("包含图片信息")).toBeInTheDocument();
    expect(screen.getByText("包含表格信息")).toBeInTheDocument();
    expect(screen.getByText("包含内容控件")).toBeInTheDocument();
    expect(screen.getByText("详细元数据")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "获取选中内容" })).toBeInTheDocument();
  });

  it("应该显示空状态提示 / Should show empty state message", () => {
    render(<SelectedContent />);

    expect(screen.getByText(/在文档中选中内容后，点击上方按钮获取选中内容信息/)).toBeInTheDocument();
  });

  it("应该能够获取选中内容 / Should get selected content", async () => {
    const mockContentInfo = {
      text: "这是选中的文本",
      elements: [
        {
          id: "sel-para-0",
          type: "Paragraph" as const,
          text: "这是选中的文本",
        },
      ],
      metadata: {
        isEmpty: false,
        characterCount: 8,
        paragraphCount: 1,
        tableCount: 0,
        imageCount: 0,
      },
    };

    vi.mocked(wordTools.getSelectedContent).mockResolvedValue(mockContentInfo);

    render(<SelectedContent />);

    const button = screen.getByRole("button", { name: "获取选中内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/成功获取选中内容/)).toBeInTheDocument();
    });

    expect(screen.getByText("字符数")).toBeInTheDocument();
    expect(screen.getByText("元素总数")).toBeInTheDocument();
    expect(screen.getByText("📄 选中文本预览")).toBeInTheDocument();
    expect(screen.getAllByText("这是选中的文本").length).toBeGreaterThan(0);
  });

  it("应该能够处理空选择 / Should handle empty selection", async () => {
    const mockContentInfo = {
      text: "",
      elements: [],
      metadata: {
        isEmpty: true,
        characterCount: 0,
        paragraphCount: 0,
        tableCount: 0,
        imageCount: 0,
      },
    };

    vi.mocked(wordTools.getSelectedContent).mockResolvedValue(mockContentInfo);

    render(<SelectedContent />);

    const button = screen.getByRole("button", { name: "获取选中内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getAllByText(/当前没有选中任何内容/).length).toBeGreaterThan(0);
    });

    expect(screen.getByText(/请在文档中选中文本、表格或其他内容后重试/)).toBeInTheDocument();
  });

  it("应该能够显示多个元素 / Should display multiple elements", async () => {
    const mockContentInfo = {
      text: "段落1\n段落2",
      elements: [
        {
          id: "sel-para-0",
          type: "Paragraph" as const,
          text: "段落1",
        },
        {
          id: "sel-para-1",
          type: "Paragraph" as const,
          text: "段落2",
        },
        {
          id: "sel-table-2",
          type: "Table" as const,
          rowCount: 2,
          columnCount: 3,
        },
      ],
      metadata: {
        isEmpty: false,
        characterCount: 8,
        paragraphCount: 2,
        tableCount: 1,
        imageCount: 0,
      },
    };

    vi.mocked(wordTools.getSelectedContent).mockResolvedValue(mockContentInfo);

    render(<SelectedContent />);

    const button = screen.getByRole("button", { name: "获取选中内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/包含 3 个元素/)).toBeInTheDocument();
    });

    expect(screen.getByText("📦 内容元素 (3)")).toBeInTheDocument();
    expect(screen.getByText("段落1")).toBeInTheDocument();
    expect(screen.getByText("段落2")).toBeInTheDocument();
  });

  it("应该能够处理错误 / Should handle errors", async () => {
    vi.mocked(wordTools.getSelectedContent).mockRejectedValue(new Error("获取失败"));

    render(<SelectedContent />);

    const button = screen.getByRole("button", { name: "获取选中内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/❌ 获取失败/)).toBeInTheDocument();
    });
  });

  it("应该能够切换选项开关 / Should toggle option switches", () => {
    render(<SelectedContent />);

    const textSwitch = screen.getAllByRole("switch")[0]; // 包含文本内容
    const imagesSwitch = screen.getAllByRole("switch")[1]; // 包含图片信息
    const tablesSwitch = screen.getAllByRole("switch")[2]; // 包含表格信息
    const controlsSwitch = screen.getAllByRole("switch")[3]; // 包含内容控件
    const metadataSwitch = screen.getAllByRole("switch")[4]; // 详细元数据

    // 初始状态检查 / Check initial state
    expect(textSwitch).toBeChecked();
    expect(imagesSwitch).toBeChecked();
    expect(tablesSwitch).toBeChecked();
    expect(controlsSwitch).toBeChecked();
    expect(metadataSwitch).not.toBeChecked();

    // 切换开关 / Toggle switches
    fireEvent.click(textSwitch);
    expect(textSwitch).not.toBeChecked();

    fireEvent.click(metadataSwitch);
    expect(metadataSwitch).toBeChecked();
  });

  it("应该在加载时禁用按钮 / Should disable button during loading", async () => {
    vi.mocked(wordTools.getSelectedContent).mockImplementation(
      () => new Promise((resolve) => setTimeout(resolve, 100))
    );

    render(<SelectedContent />);

    const button = screen.getByRole("button", { name: "获取选中内容" });
    fireEvent.click(button);

    expect(button).toBeDisabled();

    await waitFor(() => {
      expect(button).not.toBeDisabled();
    });
  });

  it("应该显示统计信息 / Should display statistics", async () => {
    const mockContentInfo = {
      text: "测试内容",
      elements: [
        {
          id: "sel-para-0",
          type: "Paragraph" as const,
          text: "段落",
        },
        {
          id: "sel-img-1",
          type: "InlinePicture" as const,
          width: 100,
          height: 100,
        },
      ],
      metadata: {
        isEmpty: false,
        characterCount: 100,
        paragraphCount: 1,
        tableCount: 0,
        imageCount: 1,
      },
    };

    vi.mocked(wordTools.getSelectedContent).mockResolvedValue(mockContentInfo);

    render(<SelectedContent />);

    const button = screen.getByRole("button", { name: "获取选中内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("100")).toBeInTheDocument(); // 字符数
      expect(screen.getByText("字符数")).toBeInTheDocument();
      expect(screen.getByText("元素总数")).toBeInTheDocument();
      expect(screen.getByText("段落数")).toBeInTheDocument();
      expect(screen.getByText("表格数")).toBeInTheDocument();
      expect(screen.getByText("图片数")).toBeInTheDocument();
    });
  });

  it("应该显示元素类型图标 / Should display element type icons", async () => {
    const mockContentInfo = {
      text: "内容",
      elements: [
        {
          id: "sel-para-0",
          type: "Paragraph" as const,
          text: "段落",
        },
        {
          id: "sel-table-1",
          type: "Table" as const,
          rowCount: 2,
          columnCount: 2,
        },
      ],
      metadata: {
        isEmpty: false,
        characterCount: 10,
        paragraphCount: 1,
        tableCount: 1,
        imageCount: 0,
      },
    };

    vi.mocked(wordTools.getSelectedContent).mockResolvedValue(mockContentInfo);

    render(<SelectedContent />);

    const button = screen.getByRole("button", { name: "获取选中内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getAllByText("段落").length).toBeGreaterThan(0);
      expect(screen.getAllByText("表格").length).toBeGreaterThan(0);
    });
  });

  it("应该能够显示详细元数据 / Should display detailed metadata", async () => {
    const mockContentInfo = {
      text: "内容",
      elements: [
        {
          id: "sel-para-0",
          type: "Paragraph" as const,
          text: "段落",
          style: "Heading1",
          alignment: "Left",
          isListItem: true,
        },
        {
          id: "sel-table-1",
          type: "Table" as const,
          rowCount: 3,
          columnCount: 4,
        },
      ],
      metadata: {
        isEmpty: false,
        characterCount: 10,
        paragraphCount: 1,
        tableCount: 1,
        imageCount: 0,
      },
    };

    vi.mocked(wordTools.getSelectedContent).mockResolvedValue(mockContentInfo);

    render(<SelectedContent />);

    // 先开启详细元数据选项 / Enable detailed metadata option
    const metadataSwitch = screen.getAllByRole("switch")[4];
    fireEvent.click(metadataSwitch);

    const button = screen.getByRole("button", { name: "获取选中内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/样式: Heading1/)).toBeInTheDocument();
      expect(screen.getByText(/对齐: Left/)).toBeInTheDocument();
      expect(screen.getByText("列表项")).toBeInTheDocument();
      expect(screen.getByText(/3 行/)).toBeInTheDocument();
      expect(screen.getByText(/4 列/)).toBeInTheDocument();
    });
  });
});
