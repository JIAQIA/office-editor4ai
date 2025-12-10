/**
 * 文件名: InsertEquationDebug.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/react, vitest
 * 描述: InsertEquationDebug 组件单元测试
 */

import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, vi, beforeEach } from "vitest";
import { InsertEquationDebug } from "../../../../src/taskpane/components/tools/InsertEquationDebug";

// 模拟 insertEquation 函数 / Mock insertEquation function
vi.mock("../../../../src/word-tools", async () => {
  const actual = await vi.importActual("../../../../src/word-tools");
  return {
    ...actual,
    insertEquation: vi.fn().mockResolvedValue({
      success: true,
      latex: "E = mc^2",
    }),
  };
});

describe("InsertEquationDebug 组件测试 / InsertEquationDebug Component Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("应该正确渲染组件", () => {
    render(<InsertEquationDebug />);
    
    expect(screen.getByText("插入公式")).toBeInTheDocument();
    expect(screen.getByText(/使用 LaTeX 格式插入数学公式/)).toBeInTheDocument();
  });

  it("应该有默认的 LaTeX 值 / Should have default LaTeX value", () => {
    render(<InsertEquationDebug />);
    
    const textarea = screen.getByPlaceholderText(/输入 LaTeX 格式的公式/);
    expect(textarea).toHaveValue("E = mc^2");
  });

  it("应该能够更新 LaTeX 输入 / Should be able to update LaTeX input", async () => {
    const user = userEvent.setup();
    render(<InsertEquationDebug />);
    
    const textarea = screen.getByPlaceholderText(/输入 LaTeX 格式的公式/);
    await user.clear(textarea);
    await user.type(textarea, "\\frac{a}{b}");
    
    expect(textarea).toHaveValue("\\frac{a}{b}");
  });

  it("点击示例应该更新 LaTeX 输入 / Should update LaTeX input when clicking example", async () => {
    const user = userEvent.setup();
    render(<InsertEquationDebug />);
    
    const exampleItem = screen.getByText(/勾股定理/);
    await user.click(exampleItem);
    
    const textarea = screen.getByPlaceholderText(/输入 LaTeX 格式的公式/);
    expect(textarea).toHaveValue("a^2 + b^2 = c^2");
  });

  it("成功插入公式时应该显示成功消息 / Should show success message on successful insert", async () => {
    const user = userEvent.setup();
    const { insertEquation } = await import("../../../../src/word-tools");

    render(<InsertEquationDebug />);
    
    const insertButton = screen.getByRole("button", { name: "插入公式" });
    await user.click(insertButton);

    await waitFor(() => {
      expect(screen.getByText(/成功在.*插入公式/)).toBeInTheDocument();
    });

    expect(insertEquation).toHaveBeenCalledWith("E = mc^2", "End");
  });

  it("插入失败时应该显示错误消息 / Should show error message on failed insert", async () => {
    const user = userEvent.setup();
    const { insertEquation } = await import("../../../../src/word-tools");

    vi.mocked(insertEquation).mockResolvedValueOnce({
      success: false,
      error: "测试错误",
    });

    render(<InsertEquationDebug />);
    
    const insertButton = screen.getByRole("button", { name: "插入公式" });
    await user.click(insertButton);

    await waitFor(() => {
      expect(screen.getByText(/插入失败: 测试错误/)).toBeInTheDocument();
    });
  });

  it("LaTeX 为空时应该显示错误 / Should show error when LaTeX is empty", async () => {
    const user = userEvent.setup();
    render(<InsertEquationDebug />);
    
    const textarea = screen.getByPlaceholderText(/输入 LaTeX 格式的公式/);
    await user.clear(textarea);
    
    const insertButton = screen.getByRole("button", { name: "插入公式" });
    await user.click(insertButton);

    await waitFor(() => {
      expect(screen.getByText("请输入 LaTeX 公式")).toBeInTheDocument();
    });
  });

  it("点击重置按钮应该重置表单 / Should reset form when clicking reset button", async () => {
    const user = userEvent.setup();
    render(<InsertEquationDebug />);
    
    const textarea = screen.getByPlaceholderText(/输入 LaTeX 格式的公式/);
    await user.clear(textarea);
    await user.type(textarea, "\\frac{a}{b}");
    
    const resetButton = screen.getByRole("button", { name: "重置" });
    await user.click(resetButton);

    expect(textarea).toHaveValue("E = mc^2");
  });

  it("应该显示所有示例公式 / Should display all example equations", () => {
    render(<InsertEquationDebug />);
    
    expect(screen.getByText(/质能方程/)).toBeInTheDocument();
    expect(screen.getByText(/勾股定理/)).toBeInTheDocument();
    expect(screen.getByText(/分数/)).toBeInTheDocument();
    expect(screen.getByText(/根号/)).toBeInTheDocument();
    expect(screen.getByText(/求和/)).toBeInTheDocument();
    expect(screen.getByText(/积分/)).toBeInTheDocument();
  });
});
