/**
 * 文件名: TextBoxContent.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: TextBoxContent 组件的单元测试 | Unit tests for TextBoxContent component
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { render, screen, fireEvent, waitFor } from "@testing-library/react";
import TextBoxContent from "../../../src/taskpane/components/tools/TextBoxContent";
import * as wordTools from "../../../src/word-tools";

// Mock word-tools 模块 / Mock word-tools module
vi.mock("../../../src/word-tools", () => ({
  getTextBoxes: vi.fn(),
}));

describe("TextBoxContent 组件 / TextBoxContent Component", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("应该正确渲染初始状态 / Should render initial state correctly", () => {
    render(<TextBoxContent />);

    expect(screen.getByText("获取选项")).toBeInTheDocument();
    expect(screen.getByText("包含文本内容")).toBeInTheDocument();
    expect(screen.getByText("包含段落详情")).toBeInTheDocument();
    expect(screen.getByText("详细元数据")).toBeInTheDocument();
    expect(screen.getByPlaceholderText("不限制")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "获取文本框内容" })).toBeInTheDocument();
  });

  it("应该能够获取文本框内容 / Should get text box content", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
        text: "这是文本框1的内容",
        width: 200,
        height: 100,
        left: 50,
        top: 50,
        rotation: 0,
        visible: true,
        lockAspectRatio: false,
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/找到 1 个文本框/)).toBeInTheDocument();
    });

    expect(screen.getByText(/文本框 1/)).toBeInTheDocument();
    expect(screen.getByText("这是文本框1的内容")).toBeInTheDocument();
  });

  it("应该能够处理空结果 / Should handle empty results", async () => {
    vi.mocked(wordTools.getTextBoxes).mockResolvedValue([]);

    render(<TextBoxContent />);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("未找到文本框")).toBeInTheDocument();
    });
  });

  it("应该能够处理多个文本框 / Should handle multiple text boxes", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
        text: "文本框1",
      },
      {
        id: "textbox-2",
        name: "TextBox2",
        text: "文本框2",
      },
      {
        id: "textbox-3",
        name: "TextBox3",
        text: "文本框3",
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/找到 3 个文本框/)).toBeInTheDocument();
    });

    expect(screen.getByText(/文本框 1/)).toBeInTheDocument();
    expect(screen.getByText(/文本框 2/)).toBeInTheDocument();
    expect(screen.getByText(/文本框 3/)).toBeInTheDocument();
    expect(screen.getByText("文本框1")).toBeInTheDocument();
    expect(screen.getByText("文本框2")).toBeInTheDocument();
    expect(screen.getByText("文本框3")).toBeInTheDocument();
  });

  it("应该能够处理错误 / Should handle errors", async () => {
    vi.mocked(wordTools.getTextBoxes).mockRejectedValue(new Error("获取失败"));

    render(<TextBoxContent />);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/错误: 获取失败/)).toBeInTheDocument();
    });
  });

  it("应该能够切换选项开关 / Should toggle option switches", () => {
    render(<TextBoxContent />);

    const textSwitch = screen.getAllByRole("switch")[0]; // 包含文本内容
    const paragraphsSwitch = screen.getAllByRole("switch")[1]; // 包含段落详情
    const metadataSwitch = screen.getAllByRole("switch")[2]; // 详细元数据

    // 初始状态检查 / Check initial state
    expect(textSwitch).toBeChecked();
    expect(paragraphsSwitch).not.toBeChecked();
    expect(metadataSwitch).not.toBeChecked();

    // 切换开关 / Toggle switches
    fireEvent.click(textSwitch);
    expect(textSwitch).not.toBeChecked();

    fireEvent.click(paragraphsSwitch);
    expect(paragraphsSwitch).toBeChecked();

    fireEvent.click(metadataSwitch);
    expect(metadataSwitch).toBeChecked();
  });

  it("应该在加载时禁用按钮 / Should disable button during loading", async () => {
    vi.mocked(wordTools.getTextBoxes).mockImplementation(
      () => new Promise((resolve) => setTimeout(() => resolve([]), 100))
    );

    render(<TextBoxContent />);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    expect(button).toBeDisabled();

    await waitFor(() => {
      expect(button).not.toBeDisabled();
    });
  });

  it("应该能够显示详细元数据 / Should display detailed metadata", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
        text: "文本框内容",
        width: 200,
        height: 100,
        left: 50,
        top: 50,
        rotation: 0,
        visible: true,
        lockAspectRatio: false,
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    // 先开启详细元数据选项 / Enable detailed metadata option
    const metadataSwitch = screen.getAllByRole("switch")[2];
    fireEvent.click(metadataSwitch);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/找到 1 个文本框/)).toBeInTheDocument();
    });

    expect(screen.getByText("ID:")).toBeInTheDocument();
    expect(screen.getByText("textbox-1")).toBeInTheDocument();
    expect(screen.getByText("宽度:")).toBeInTheDocument();
    expect(screen.getByText("200.00 pt")).toBeInTheDocument();
    expect(screen.getByText("高度:")).toBeInTheDocument();
    expect(screen.getByText("100.00 pt")).toBeInTheDocument();
    expect(screen.getByText("左边距:")).toBeInTheDocument();
    expect(screen.getByText("上边距:")).toBeInTheDocument();
    expect(screen.getAllByText("50.00 pt").length).toBe(2); // 左边距和上边距都是 50.00 pt
    expect(screen.getByText("旋转角度:")).toBeInTheDocument();
    expect(screen.getByText("0°")).toBeInTheDocument();
    expect(screen.getByText("可见性:")).toBeInTheDocument();
    expect(screen.getByText("可见")).toBeInTheDocument();
    expect(screen.getByText("锁定纵横比:")).toBeInTheDocument();
    expect(screen.getByText("否")).toBeInTheDocument();
  });

  it("应该能够显示段落详情 / Should display paragraph details", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
        text: "文本框内容",
        paragraphs: [
          {
            id: "textbox-1-para-0",
            type: "Paragraph" as const,
            text: "段落1",
            style: "Normal",
            alignment: "Left",
            isListItem: false,
          },
          {
            id: "textbox-1-para-1",
            type: "Paragraph" as const,
            text: "段落2",
            style: "Normal",
            alignment: "Center",
            isListItem: true,
          },
        ],
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    // 先开启段落详情选项 / Enable paragraph details option
    const paragraphsSwitch = screen.getAllByRole("switch")[1];
    fireEvent.click(paragraphsSwitch);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/段落详情 \(2 个段落\)/)).toBeInTheDocument();
    });

    expect(screen.getAllByText(/段落1/).length).toBeGreaterThan(0);
    expect(screen.getAllByText(/段落2/).length).toBeGreaterThan(0);
  });

  it("应该能够显示段落详细元数据 / Should display paragraph detailed metadata", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
        text: "文本框内容",
        paragraphs: [
          {
            id: "textbox-1-para-0",
            type: "Paragraph" as const,
            text: "段落1",
            style: "Heading1",
            alignment: "Center",
            isListItem: true,
          },
        ],
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    // 开启段落详情和详细元数据选项 / Enable paragraph details and detailed metadata options
    const paragraphsSwitch = screen.getAllByRole("switch")[1];
    const metadataSwitch = screen.getAllByRole("switch")[2];
    fireEvent.click(paragraphsSwitch);
    fireEvent.click(metadataSwitch);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getAllByText(/样式: Heading1/).length).toBeGreaterThan(0);
      expect(screen.getAllByText(/对齐: Center/).length).toBeGreaterThan(0);
      expect(screen.getAllByText(/列表项: 是/).length).toBeGreaterThan(0);
    });
  });

  it("应该能够设置最大文本长度 / Should set max text length", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        text: "短文本",
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    const input = screen.getByPlaceholderText("不限制");
    fireEvent.change(input, { target: { value: "100" } });

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(wordTools.getTextBoxes).toHaveBeenCalledWith({
        includeText: true,
        includeParagraphs: false,
        detailedMetadata: false,
        maxTextLength: 100,
      });
    });
  });

  it("应该能够显示 JSON 输出 / Should display JSON output", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
        text: "文本框内容",
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("JSON 输出")).toBeInTheDocument();
    });

    const jsonOutput = screen.getByText(/"id": "textbox-1"/);
    expect(jsonOutput).toBeInTheDocument();
  });

  it("应该能够显示文本框名称 / Should display text box name", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "我的文本框",
        text: "内容",
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/文本框 1: 我的文本框/)).toBeInTheDocument();
    });
  });

  it("应该在没有文本时不显示文本内容区域 / Should not display text content when text is not included", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    // 关闭文本内容选项 / Disable text content option
    const textSwitch = screen.getAllByRole("switch")[0];
    fireEvent.click(textSwitch);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/找到 1 个文本框/)).toBeInTheDocument();
    });

    expect(screen.queryByText("文本内容:")).not.toBeInTheDocument();
  });

  it("应该正确处理不可见的文本框 / Should handle invisible text boxes correctly", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
        text: "内容",
        visible: false,
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    // 开启详细元数据 / Enable detailed metadata
    const metadataSwitch = screen.getAllByRole("switch")[2];
    fireEvent.click(metadataSwitch);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("可见性:")).toBeInTheDocument();
      expect(screen.getByText("隐藏")).toBeInTheDocument();
    });
  });

  it("应该正确处理锁定纵横比的文本框 / Should handle locked aspect ratio text boxes correctly", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
        text: "内容",
        lockAspectRatio: true,
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    // 开启详细元数据 / Enable detailed metadata
    const metadataSwitch = screen.getAllByRole("switch")[2];
    fireEvent.click(metadataSwitch);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("锁定纵横比:")).toBeInTheDocument();
      expect(screen.getByText("是")).toBeInTheDocument();
    });
  });

  it("应该正确处理有旋转角度的文本框 / Should handle rotated text boxes correctly", async () => {
    const mockTextBoxes = [
      {
        id: "textbox-1",
        name: "TextBox1",
        text: "内容",
        rotation: 45,
      },
    ];

    vi.mocked(wordTools.getTextBoxes).mockResolvedValue(mockTextBoxes);

    render(<TextBoxContent />);

    // 开启详细元数据 / Enable detailed metadata
    const metadataSwitch = screen.getAllByRole("switch")[2];
    fireEvent.click(metadataSwitch);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("旋转角度:")).toBeInTheDocument();
      expect(screen.getByText("45°")).toBeInTheDocument();
    });
  });

  it("应该正确传递所有选项到 getTextBoxes 函数 / Should pass all options to getTextBoxes function correctly", async () => {
    vi.mocked(wordTools.getTextBoxes).mockResolvedValue([]);

    render(<TextBoxContent />);

    // 设置所有选项 / Set all options
    const textSwitch = screen.getAllByRole("switch")[0];
    const paragraphsSwitch = screen.getAllByRole("switch")[1];
    const metadataSwitch = screen.getAllByRole("switch")[2];
    const input = screen.getByPlaceholderText("不限制");

    fireEvent.click(textSwitch); // 关闭文本内容
    fireEvent.click(paragraphsSwitch); // 开启段落详情
    fireEvent.click(metadataSwitch); // 开启详细元数据
    fireEvent.change(input, { target: { value: "50" } });

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(wordTools.getTextBoxes).toHaveBeenCalledWith({
        includeText: false,
        includeParagraphs: true,
        detailedMetadata: true,
        maxTextLength: 50,
      });
    });
  });

  it("应该在清空最大文本长度输入时不传递该参数 / Should not pass maxTextLength when input is empty", async () => {
    vi.mocked(wordTools.getTextBoxes).mockResolvedValue([]);

    render(<TextBoxContent />);

    const button = screen.getByRole("button", { name: "获取文本框内容" });
    fireEvent.click(button);

    await waitFor(() => {
      expect(wordTools.getTextBoxes).toHaveBeenCalledWith({
        includeText: true,
        includeParagraphs: false,
        detailedMetadata: false,
        maxTextLength: undefined,
      });
    });
  });
});
