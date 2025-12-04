/**
 * 文件名: InsertSectionBreakDebug.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/react, vitest
 * 描述: InsertSectionBreakDebug组件的测试 / Test for InsertSectionBreakDebug component
 */

import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, vi, beforeEach } from "vitest";
import { InsertSectionBreakDebug } from "../../../../src/taskpane/components/tools/InsertSectionBreakDebug";

// 模拟insertSectionBreak函数 / Mock insertSectionBreak function
vi.mock("../../../../src/word-tools", async () => {
  const actual = await vi.importActual("../../../../src/word-tools");
  return {
    ...actual,
    insertSectionBreak: vi.fn().mockResolvedValue({
      success: true,
      sectionIndex: 1,
    }),
  };
});

describe("InsertSectionBreakDebug 组件测试 / InsertSectionBreakDebug Component Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("基本渲染测试 / Basic Rendering Tests", () => {
    it("应该渲染组件标题 / Should render component title", () => {
      render(<InsertSectionBreakDebug />);

      const titles = screen.getAllByText("插入分节符");
      expect(titles.length).toBeGreaterThan(0);
    });

    it("应该渲染描述文本 / Should render description text", () => {
      render(<InsertSectionBreakDebug />);

      expect(
        screen.getByText(/分节符用于将文档分成不同的节/)
      ).toBeInTheDocument();
    });

    it("应该渲染分节符类型下拉框 / Should render section break type dropdown", () => {
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      expect(dropdowns.length).toBeGreaterThanOrEqual(2);
    });

    it("应该渲染插入位置下拉框 / Should render insert location dropdown", () => {
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      expect(dropdowns.length).toBeGreaterThanOrEqual(2);
    });

    it("应该渲染插入分节符按钮 / Should render insert section break button", () => {
      render(<InsertSectionBreakDebug />);

      expect(screen.getByRole('button', { name: '插入分节符' })).toBeInTheDocument();
    });

    it("应该渲染重置按钮 / Should render reset button", () => {
      render(<InsertSectionBreakDebug />);

      expect(screen.getByText("重置")).toBeInTheDocument();
    });

    it("应该渲染使用提示信息 / Should render usage tips", () => {
      render(<InsertSectionBreakDebug />);

      expect(screen.getByText("使用提示：")).toBeInTheDocument();
      expect(screen.getByText(/连续分节符/)).toBeInTheDocument();
      expect(screen.getByText(/下一页分节符/)).toBeInTheDocument();
    });
  });

  describe("分节符类型选择测试 / Section Break Type Selection Tests", () => {
    it("默认选中下一页分节符 / Should default to next page section break", () => {
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      const typeDropdown = dropdowns[0];
      expect(typeDropdown).toHaveValue("下一页");
    });

    it("应该能够选择连续分节符 / Should be able to select continuous section break", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      const typeDropdown = dropdowns[0];
      await user.click(typeDropdown);

      const continuousOption = screen.getByText("连续");
      await user.click(continuousOption);

      await waitFor(() => {
        expect(typeDropdown).toHaveValue("连续");
      });
    });

    it("应该能够选择奇数页分节符 / Should be able to select odd page section break", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      const typeDropdown = dropdowns[0];
      await user.click(typeDropdown);

      const oddPageOption = screen.getByText("奇数页");
      await user.click(oddPageOption);

      await waitFor(() => {
        expect(typeDropdown).toHaveValue("奇数页");
      });
    });

    it("应该能够选择偶数页分节符 / Should be able to select even page section break", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      const typeDropdown = dropdowns[0];
      await user.click(typeDropdown);

      const evenPageOption = screen.getByText("偶数页");
      await user.click(evenPageOption);

      await waitFor(() => {
        expect(typeDropdown).toHaveValue("偶数页");
      });
    });
  });

  describe("插入位置选择测试 / Insert Location Selection Tests", () => {
    it("默认选中文档末尾 / Should default to document end", () => {
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      const locationDropdown = dropdowns[1];
      expect(locationDropdown).toHaveValue("文档末尾");
    });

    it("应该能够选择文档开头 / Should be able to select document start", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      const locationDropdown = dropdowns[1];
      await user.click(locationDropdown);

      const startOption = screen.getByText("文档开头");
      await user.click(startOption);

      await waitFor(() => {
        expect(locationDropdown).toHaveValue("文档开头");
      });
    });

    it("应该能够选择选中内容之前 / Should be able to select before selection", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      const locationDropdown = dropdowns[1];
      await user.click(locationDropdown);

      const beforeOption = screen.getByText("选中内容之前");
      await user.click(beforeOption);

      await waitFor(() => {
        expect(locationDropdown).toHaveValue("选中内容之前");
      });
    });

    it("应该能够选择选中内容之后 / Should be able to select after selection", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      const locationDropdown = dropdowns[1];
      await user.click(locationDropdown);

      const afterOption = screen.getByText("选中内容之后");
      await user.click(afterOption);

      await waitFor(() => {
        expect(locationDropdown).toHaveValue("选中内容之后");
      });
    });

    it("应该能够选择替换选中内容 / Should be able to select replace selection", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const dropdowns = screen.getAllByRole("combobox");
      const locationDropdown = dropdowns[1];
      await user.click(locationDropdown);

      const replaceOption = screen.getByText("替换选中内容");
      await user.click(replaceOption);

      await waitFor(() => {
        expect(locationDropdown).toHaveValue("替换选中内容");
      });
    });
  });

  describe("插入分节符功能测试 / Insert Section Break Function Tests", () => {
    it("点击插入按钮应该调用insertSectionBreak函数 / Should call insertSectionBreak when insert button clicked", async () => {
      const user = userEvent.setup();
      const { insertSectionBreak } = await import("../../../../src/word-tools");

      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(insertSectionBreak).toHaveBeenCalledWith("NextPage", "End");
      });
    });

    it("插入成功应该显示成功消息 / Should show success message on successful insert", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/成功在 文档末尾 插入下一页分节符/)).toBeInTheDocument();
      });
    });

    it("插入成功应该显示新节索引 / Should show new section index on successful insert", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/新节索引: 1/)).toBeInTheDocument();
      });
    });

    it("插入失败应该显示错误消息 / Should show error message on failed insert", async () => {
      const user = userEvent.setup();
      const { insertSectionBreak } = await import("../../../../src/word-tools");

      // 模拟插入失败 / Mock insert failure
      vi.mocked(insertSectionBreak).mockResolvedValueOnce({
        success: false,
        error: "插入失败",
      });

      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/插入失败: 插入失败/)).toBeInTheDocument();
      });
    });

    it("插入时应该显示加载状态 / Should show loading state during insert", async () => {
      const user = userEvent.setup();
      const { insertSectionBreak } = await import("../../../../src/word-tools");

      // 模拟延迟 / Mock delay
      vi.mocked(insertSectionBreak).mockImplementation(
        () =>
          new Promise((resolve) =>
            setTimeout(
              () =>
                resolve({
                  success: true,
                  sectionIndex: 1,
                }),
              100
            )
          )
      );

      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      // 检查按钮是否被禁用 / Check if button is disabled
      expect(insertButton).toBeDisabled();
    });

    it("应该能够在不同位置插入不同类型的分节符 / Should be able to insert different types at different locations", async () => {
      const user = userEvent.setup();
      const { insertSectionBreak } = await import("../../../../src/word-tools");

      render(<InsertSectionBreakDebug />);

      // 选择连续分节符 / Select continuous section break
      const dropdowns = screen.getAllByRole("combobox");
      const typeDropdown = dropdowns[0];
      await user.click(typeDropdown);
      const continuousOption = screen.getByText("连续");
      await user.click(continuousOption);

      // 选择文档开头 / Select document start
      const locationDropdown = dropdowns[1];
      await user.click(locationDropdown);
      const startOption = screen.getByText("文档开头");
      await user.click(startOption);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(insertSectionBreak).toHaveBeenCalledWith("Continuous", "Start");
      });
    });
  });

  describe("重置功能测试 / Reset Function Tests", () => {
    it("点击重置按钮应该重置分节符类型和插入位置 / Should reset section break type and location when reset button clicked", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      // 更改分节符类型和插入位置 / Change section break type and location
      const dropdowns = screen.getAllByRole("combobox");
      const typeDropdown = dropdowns[0];
      await user.click(typeDropdown);
      const continuousOption = screen.getByText("连续");
      await user.click(continuousOption);

      const locationDropdown = dropdowns[1];
      await user.click(locationDropdown);
      const startOption = screen.getByText("文档开头");
      await user.click(startOption);

      await waitFor(() => {
        expect(typeDropdown).toHaveValue("连续");
        expect(locationDropdown).toHaveValue("文档开头");
      });

      // 点击重置 / Click reset
      const resetButton = screen.getByText("重置");
      await user.click(resetButton);

      await waitFor(() => {
        expect(typeDropdown).toHaveValue("下一页");
        expect(locationDropdown).toHaveValue("文档末尾");
      });
    });

    it("点击重置按钮应该清除结果消息 / Should clear result message when reset button clicked", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      // 插入分节符 / Insert section break
      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/成功在 文档末尾 插入下一页分节符/)).toBeInTheDocument();
      });

      // 点击重置 / Click reset
      const resetButton = screen.getByText("重置");
      await user.click(resetButton);

      await waitFor(() => {
        expect(screen.queryByText(/成功在 文档末尾 插入下一页分节符/)).not.toBeInTheDocument();
      });
    });
  });

  describe("边界情况测试 / Edge Case Tests", () => {
    it("插入时抛出异常应该显示错误消息 / Should show error message when exception thrown during insert", async () => {
      const user = userEvent.setup();
      const { insertSectionBreak } = await import("../../../../src/word-tools");

      // 模拟抛出异常 / Mock exception
      vi.mocked(insertSectionBreak).mockRejectedValueOnce(new Error("网络错误"));

      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/插入失败: 网络错误/)).toBeInTheDocument();
      });
    });

    it("插入失败但没有错误信息应该显示默认错误 / Should show default error when insert fails without error message", async () => {
      const user = userEvent.setup();
      const { insertSectionBreak } = await import("../../../../src/word-tools");

      // 模拟插入失败但没有错误信息 / Mock insert failure without error message
      vi.mocked(insertSectionBreak).mockResolvedValueOnce({
        success: false,
      });

      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        expect(screen.getByText(/插入失败: 未知错误/)).toBeInTheDocument();
      });
    });

    it("加载时重置按钮应该被禁用 / Reset button should be disabled during loading", async () => {
      const user = userEvent.setup();
      const { insertSectionBreak } = await import("../../../../src/word-tools");

      // 模拟延迟 / Mock delay
      vi.mocked(insertSectionBreak).mockImplementation(
        () =>
          new Promise((resolve) =>
            setTimeout(
              () =>
                resolve({
                  success: true,
                  sectionIndex: 1,
                }),
              100
            )
          )
      );

      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      const resetButton = screen.getByText("重置");
      expect(resetButton).toBeDisabled();
    });
  });

  describe("UI样式测试 / UI Style Tests", () => {
    it("成功消息应该有成功样式 / Success message should have success style", async () => {
      const user = userEvent.setup();
      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        const successMessage = screen.getByText(/成功在 文档末尾 插入下一页分节符/);
        expect(successMessage).toBeInTheDocument();
      });
    });

    it("错误消息应该有错误样式 / Error message should have error style", async () => {
      const user = userEvent.setup();
      const { insertSectionBreak } = await import("../../../../src/word-tools");

      vi.mocked(insertSectionBreak).mockResolvedValueOnce({
        success: false,
        error: "测试错误",
      });

      render(<InsertSectionBreakDebug />);

      const insertButton = screen.getByRole('button', { name: '插入分节符' });
      await user.click(insertButton);

      await waitFor(() => {
        const errorMessage = screen.getByText(/插入失败: 测试错误/);
        expect(errorMessage).toBeInTheDocument();
      });
    });
  });
});
