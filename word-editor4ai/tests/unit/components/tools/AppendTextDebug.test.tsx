/**
 * 文件名: AppendTextDebug.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: AppendTextDebug组件的测试
 */

import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, vi, beforeEach } from "vitest";
import { AppendTextDebug } from "../../../../src/taskpane/components/tools/AppendTextDebug";

// 模拟appendText函数
const mockAppendText = vi.fn().mockResolvedValue(undefined);

describe("AppendTextDebug", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("应该渲染默认状态", () => {
    render(<AppendTextDebug appendText={mockAppendText} />);

    expect(screen.getByLabelText("要追加的文本内容:")).toBeInTheDocument();
    expect(screen.getByText("追加文本")).toBeInTheDocument();
    expect(screen.getByLabelText("应用文本格式")).not.toBeChecked();
  });

  it("应该允许输入文本", async () => {
    render(<AppendTextDebug appendText={mockAppendText} />);

    const textarea = screen.getByLabelText("要追加的文本内容:");
    await userEvent.type(textarea, "新的测试文本");

    expect(textarea).toHaveValue("这是追加到文档末尾的文本新的测试文本");
  });

  it("应该切换格式选项", async () => {
    render(<AppendTextDebug appendText={mockAppendText} />);

    const formatSwitch = screen.getByLabelText("应用文本格式");
    await userEvent.click(formatSwitch);

    expect(formatSwitch).toBeChecked();
    expect(screen.getByLabelText("字体名称:")).toBeInTheDocument();
  });

  it("应该调用appendText函数", async () => {
    render(<AppendTextDebug appendText={mockAppendText} />);

    const button = screen.getByText("追加文本");
    await userEvent.click(button);

    await waitFor(() => {
      expect(mockAppendText).toHaveBeenCalledWith({
        text: "这是追加到文档末尾的文本",
        format: undefined,
        images: undefined,
      });
    });
  });

  it("应该显示成功消息", async () => {
    render(<AppendTextDebug appendText={mockAppendText} />);

    await userEvent.click(screen.getByText("追加文本"));

    await waitFor(() => {
      expect(screen.getByText("内容已成功追加（文本）")).toBeInTheDocument();
    });
  });

  it("应该显示错误消息", async () => {
    mockAppendText.mockRejectedValue(new Error("测试错误"));
    render(<AppendTextDebug appendText={mockAppendText} />);

    await userEvent.click(screen.getByText("追加文本"));

    await waitFor(() => {
      expect(screen.getByText("追加内容失败: 测试错误")).toBeInTheDocument();
    });
  });
});
