/**
 * 文件名: InsertImageDebug.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: InsertImageDebug组件的测试
 */

import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, vi, beforeEach } from "vitest";
import { InsertImageDebug } from "../../../../src/taskpane/components/tools/InsertImageDebug";

// 模拟insertImage函数
vi.mock("../../../../src/word-tools", async () => {
  const actual = await vi.importActual("../../../../src/word-tools");
  return {
    ...actual,
    insertImage: vi.fn().mockResolvedValue({
      success: true,
      imageId: "示例图片", // 使用 altTextTitle 作为标识符
    }),
  };
});

describe("InsertImageDebug", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("应该渲染默认状态", () => {
    render(<InsertImageDebug />);

    expect(screen.getByText("选择图片")).toBeInTheDocument();
    expect(screen.getByText("上传图片")).toBeInTheDocument();
    expect(screen.getByText("基本选项")).toBeInTheDocument();
    expect(screen.getByText("插入图片")).toBeInTheDocument();
  });

  it("应该显示图片上传按钮", () => {
    render(<InsertImageDebug />);

    const uploadButton = screen.getByText("上传图片");
    expect(uploadButton).toBeInTheDocument();
  });

  it("应该有宽度和高度输入框", () => {
    render(<InsertImageDebug />);

    expect(screen.getByLabelText("宽度（磅）")).toBeInTheDocument();
    expect(screen.getByLabelText("高度（磅）")).toBeInTheDocument();
  });

  it("应该有替代文本输入框", () => {
    render(<InsertImageDebug />);

    expect(screen.getByLabelText("替代文本")).toBeInTheDocument();
  });

  it("应该有插入位置下拉框", () => {
    render(<InsertImageDebug />);

    expect(screen.getByLabelText("插入位置")).toBeInTheDocument();
  });

  it("应该有布局类型下拉框", () => {
    render(<InsertImageDebug />);

    expect(screen.getByLabelText("布局类型")).toBeInTheDocument();
  });

  it("应该有保持纵横比开关", () => {
    render(<InsertImageDebug />);

    expect(screen.getByLabelText("保持纵横比")).toBeInTheDocument();
  });

  it("应该允许输入宽度", async () => {
    render(<InsertImageDebug />);

    const widthInput = screen.getByLabelText("宽度（磅）");
    await userEvent.clear(widthInput);
    await userEvent.type(widthInput, "500");

    expect(widthInput).toHaveValue(500);
  });

  it("应该允许输入高度", async () => {
    render(<InsertImageDebug />);

    const heightInput = screen.getByLabelText("高度（磅）");
    await userEvent.clear(heightInput);
    await userEvent.type(heightInput, "400");

    expect(heightInput).toHaveValue(400);
  });

  it("应该允许输入替代文本", async () => {
    render(<InsertImageDebug />);

    const altTextInput = screen.getByLabelText("替代文本");
    await userEvent.clear(altTextInput);
    await userEvent.type(altTextInput, "新的替代文本");

    expect(altTextInput).toHaveValue("新的替代文本");
  });

  it("应该切换保持纵横比", async () => {
    render(<InsertImageDebug />);

    const aspectRatioSwitch = screen.getByLabelText("保持纵横比");
    expect(aspectRatioSwitch).toBeChecked();

    await userEvent.click(aspectRatioSwitch);
    expect(aspectRatioSwitch).not.toBeChecked();
  });

  it("插入图片按钮在没有图片时应该被禁用", () => {
    render(<InsertImageDebug />);

    const insertButton = screen.getByText("插入图片");
    expect(insertButton).toBeDisabled();
  });

  it("应该显示超链接输入框", () => {
    render(<InsertImageDebug />);

    expect(screen.getByLabelText("超链接（可选）")).toBeInTheDocument();
  });

  it("应该显示详细描述输入框", () => {
    render(<InsertImageDebug />);

    expect(screen.getByLabelText("详细描述（可选）")).toBeInTheDocument();
  });
});
