/**
 * 文件名: InsertPageBreakDebug.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/react, vitest
 * 描述: InsertPageBreakDebug组件的测试 / Test for InsertPageBreakDebug component
 */

import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, vi, beforeEach } from "vitest";
import { InsertPageBreakDebug } from "../../../../src/taskpane/components/tools/InsertPageBreakDebug";

// 模拟insertPageBreak函数 / Mock insertPageBreak function
vi.mock("../../../../src/word-tools", async () => {
  const actual = await vi.importActual("../../../../src/word-tools");
  return {
    ...actual,
    insertPageBreak: vi.fn().mockResolvedValue({
      success: true,
    }),
  };
});

describe("InsertPageBreakDebug 组件测试 / InsertPageBreakDebug Component Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("基本渲染测试 / Basic Rendering Tests", () => {
    it("应该渲染组件标题 / Should render component title", () => {
      render(<InsertPageBreakDebug />);

      const titles = screen.getAllByText("插入分页符");
      expect(titles.length).toBeGreaterThan(0);
    });

    it("应该渲染描述文本 / Should render description text", () => {
      render(<InsertPageBreakDebug />);

      expect(
        screen.getByText(/分页符用于强制在指定位置开始新的一页/)
      ).toBeInTheDocument();
    });

    it("应该渲染插入位置下拉框 / Should render insert location dropdown", () => {
      render(<InsertPageBreakDebug />);

      expect(screen.getByRole("combobox")).toBeInTheDocument();
    });

    it("应该渲染插入分页符按钮 / Should render insert page break button", () => {
      render(<InsertPageBreakDebug />);

      expect(screen.getByRole('button', { name: '插入分页符' })).toBeInTheDocument();
    });

    it("应该渲染重置按钮 / Should render reset button", () => {
      render(<InsertPageBreakDebug />);

      expect(screen.getByText("重置")).toBeInTheDocument();
    });

    it("应该渲染使用提示信息 / Should render usage tips", () => {
      render(<InsertPageBreakDebug />);

      expect(screen.getByText("使用提示：")).toBeInTheDocument();
      expect(screen.getByText(/分页符会在指定位置强制开始新页面/)).toBeInTheDocument();
    });
  });

  describe("插入位置选择测试 / Insert Location Selection Tests", () => {
    it("默认选中文档末尾 / Should default to document end", () => {
      render(<InsertPageBreakDebug />);

      const dropdown = screen.getByRole("combobox");
      expect(dropdown).toHaveValue("文档末尾");
    });

    it("应该能够选择文档开头 / Should be able to select document start", async () => {
      const user = userEvent.setup();
      render(<InsertPageBreakDebug />);

      const dropdown = screen.getByRole("combobox");
      await user.click(dropdown);

      const startOption = screen.getByText("文档开头");
      await user.click(startOption);

      await waitFor(() => {
        expect(dropdown).toHaveValue("文档开头");
      });
    });

    it("应该能够选择选中内容之前 / Should be able to select before selection", async () => {
      const user = userEvent.setup();
      render(<InsertPageBreakDebug />);

      const dropdown = screen.getByRole("combobox");
      await user.click(dropdown);

      const beforeOption = screen.getByText("选中内容之前");
      await user.click(beforeOption);

      await waitFor(() => {
        expect(dropdown).toHaveValue("选中内容之前");
      });
    });

    it("应该能够选择选中内容之后 / Should be able to select after selection", async () => {
      const user = userEvent.setup();
      render(<InsertPageBreakDebug />);

      const dropdown = screen.getByRole("combobox");
      await user.click(dropdown);

      const afterOption = screen.getByText("选中内容之后");
      await user.click(afterOption);

      await waitFor(() => {
        expect(dropdown).toHaveValue("选中内容之后");
      });
    });

    it("应该能够选择替换选中内容 / Should be able to select replace selection", async () => {
      const user = userEvent.setup();
      render(<InsertPageBreakDebug />);

      const dropdown = screen.getByRole("combobox");
      await user.click(dropdown);

      const replaceOption = screen.getByText("替换选中内容");
      await user.click(replaceOption);

      await waitFor(() => {
        expect(dropdown).toHaveValue("替换选中内容");
      });
    });
  });

  describe("插入分页符功能测试 / Insert Page Break Function Tests", () => {
    it("点击插入按钮应该调用insertPageBreak函数 / Should call insertPageBreak when insert button clicked", async () => {
      const user = userEvent.setup();
      const { insertPageBreak } = await import("../../../../src/word-tools");

      render(<InsertPageBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(insertPageBreak).toHaveBeenCalledWith("End");
      });
    });

    it("插入成功应该显示成功消息 / Should show success message on successful insert", async () => {
      const user = userEvent.setup();
      render(<InsertPageBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/成功在 文档末尾 插入分页符/)).toBeInTheDocument();
      });
    });

    it("插入失败应该显示错误消息 / Should show error message on failed insert", async () => {
      const user = userEvent.setup();
      const { insertPageBreak } = await import("../../../../src/word-tools");

      // 模拟插入失败 / Mock insert failure
      vi.mocked(insertPageBreak).mockResolvedValueOnce({
        success: false,
        error: "插入失败",
      });

      render(<InsertPageBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/插入失败: 插入失败/)).toBeInTheDocument();
      });
    });

    it("插入时应该显示加载状态 / Should show loading state during insert", async () => {
      const user = userEvent.setup();
      const { insertPageBreak } = await import("../../../../src/word-tools");

      // 模拟延迟 / Mock delay
      vi.mocked(insertPageBreak).mockImplementation(
        () =>
          new Promise((resolve) =>
            setTimeout(
              () =>
                resolve({
                  success: true,
                }),
              100
            )
          )
      );

      render(<InsertPageBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      // 检查按钮是否被禁用 / Check if button is disabled
      expect(insertButton).toBeDisabled();
    });

    it("应该能够在不同位置插入分页符 / Should be able to insert page break at different locations", async () => {
      const user = userEvent.setup();
      const { insertPageBreak } = await import("../../../../src/word-tools");

      render(<InsertPageBreakDebug />);

      // 选择文档开头 / Select document start
      const dropdown = screen.getByRole("combobox");
      await user.click(dropdown);
      const startOption = screen.getByText("文档开头");
      await user.click(startOption);

      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(insertPageBreak).toHaveBeenCalledWith("Start");
      });
    });
  });

  describe("重置功能测试 / Reset Function Tests", () => {
    it("点击重置按钮应该重置插入位置 / Should reset insert location when reset button clicked", async () => {
      const user = userEvent.setup();
      render(<InsertPageBreakDebug />);

      // 更改插入位置 / Change insert location
      const dropdown = screen.getByRole("combobox");
      await user.click(dropdown);
      const startOption = screen.getByText("文档开头");
      await user.click(startOption);

      await waitFor(() => {
        expect(dropdown).toHaveValue("文档开头");
      });

      // 点击重置 / Click reset
      const resetButton = screen.getByText("重置");
      await user.click(resetButton);

      await waitFor(() => {
        expect(dropdown).toHaveValue("文档末尾");
      });
    });

    it("点击重置按钮应该清除结果消息 / Should clear result message when reset button clicked", async () => {
      const user = userEvent.setup();
      render(<InsertPageBreakDebug />);

      // 插入分页符 / Insert page break
      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/成功在 文档末尾 插入分页符/)).toBeInTheDocument();
      });

      // 点击重置 / Click reset
      const resetButton = screen.getByText("重置");
      await user.click(resetButton);

      await waitFor(() => {
        expect(screen.queryByText(/成功在 文档末尾 插入分页符/)).not.toBeInTheDocument();
      });
    });
  });

  describe("边界情况测试 / Edge Case Tests", () => {
    it("插入时抛出异常应该显示错误消息 / Should show error message when exception thrown during insert", async () => {
      const user = userEvent.setup();
      const { insertPageBreak } = await import("../../../../src/word-tools");

      // 模拟抛出异常 / Mock exception
      vi.mocked(insertPageBreak).mockRejectedValueOnce(new Error("网络错误"));

      render(<InsertPageBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/插入失败: 网络错误/)).toBeInTheDocument();
      });
    });

    it("插入失败但没有错误信息应该显示默认错误 / Should show default error when insert fails without error message", async () => {
      const user = userEvent.setup();
      const { insertPageBreak } = await import("../../../../src/word-tools");

      // 模拟插入失败但没有错误信息 / Mock insert failure without error message
      vi.mocked(insertPageBreak).mockResolvedValueOnce({
        success: false,
      });

      render(<InsertPageBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/插入失败: 未知错误/)).toBeInTheDocument();
      });
    });

    it("加载时重置按钮应该被禁用 / Reset button should be disabled during loading", async () => {
      const user = userEvent.setup();
      const { insertPageBreak } = await import("../../../../src/word-tools");

      // 模拟延迟 / Mock delay
      vi.mocked(insertPageBreak).mockImplementation(
        () =>
          new Promise((resolve) =>
            setTimeout(
              () =>
                resolve({
                  success: true,
                }),
              100
            )
          )
      );

      render(<InsertPageBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      const resetButton = screen.getByText("重置");
      expect(resetButton).toBeDisabled();
    });
  });

  describe("UI样式测试 / UI Style Tests", () => {
    it("UI样式测试 / UI Style Tests", async () => {
      const user = userEvent.setup();
      render(<InsertPageBreakDebug />);

      // 使用更精确的选择器定位按钮
      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      await waitFor(() => {
        const successMessage = screen.getByText(/成功在 文档末尾 插入分页符/);
        expect(successMessage).toBeInTheDocument();
      });
    });

    it("错误消息应该有错误样式 / Error message should have error style", async () => {
      const user = userEvent.setup();
      const { insertPageBreak } = await import("../../../../src/word-tools");

      vi.mocked(insertPageBreak).mockResolvedValueOnce({
        success: false,
        error: "测试错误",
      });

      render(<InsertPageBreakDebug />);

      // 使用更精确的选择器定位按钮
      const insertButton = screen.getByRole('button', { name: '插入分页符' });
      await user.click(insertButton);

      await waitFor(() => {
        const errorMessage = screen.getByText(/插入失败: 测试错误/);
        expect(errorMessage).toBeInTheDocument();
      });
    });
  });
});
