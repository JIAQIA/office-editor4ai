/**
 * 文件名: TableOfContentsDebug.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/react, vitest
 * 描述: TableOfContentsDebug组件的测试 / Test for TableOfContentsDebug component
 */

import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, it, expect, vi, beforeEach } from "vitest";
import { TableOfContentsDebug } from "../../../../src/taskpane/components/tools/TableOfContentsDebug";

// 模拟目录管理函数 / Mock TOC management functions
vi.mock("../../../../src/word-tools/tableOfContents", () => ({
  insertTableOfContents: vi.fn().mockResolvedValue({
    success: true,
    tocInfo: {
      index: 0,
      text: "目录\n第一章\n第二章",
      entryCount: 2,
      levels: [1, 2, 3],
    },
  }),
  updateTableOfContents: vi.fn().mockResolvedValue({
    success: true,
    updatedCount: 1,
  }),
  deleteTableOfContents: vi.fn().mockResolvedValue({
    success: true,
    deletedCount: 1,
  }),
  getTableOfContentsList: vi.fn().mockResolvedValue({
    success: true,
    tocs: [
      {
        index: 0,
        text: "目录\n第一章\n第二章",
        entryCount: 2,
        levels: [1, 2, 3],
      },
    ],
  }),
}));

describe("TableOfContentsDebug 组件测试 / TableOfContentsDebug Component Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("基本渲染测试 / Basic Rendering Tests", () => {
    it("应该渲染插入目录部分 / Should render insert TOC section", () => {
      render(<TableOfContentsDebug />);

      expect(screen.getByRole("heading", { name: "插入目录" })).toBeInTheDocument();
    });

    it("应该渲染更新目录部分 / Should render update TOC section", () => {
      render(<TableOfContentsDebug />);

      expect(screen.getByRole("heading", { name: "更新目录" })).toBeInTheDocument();
    });

    it("应该渲染删除目录部分 / Should render delete TOC section", () => {
      render(<TableOfContentsDebug />);

      expect(screen.getByRole("heading", { name: "删除目录" })).toBeInTheDocument();
    });

    it("应该渲染获取目录列表部分 / Should render get TOC list section", () => {
      render(<TableOfContentsDebug />);

      expect(screen.getByRole("heading", { name: "获取目录列表" })).toBeInTheDocument();
    });

    it("应该渲染插入位置按钮组 / Should render insert location button group", () => {
      render(<TableOfContentsDebug />);

      expect(screen.getByRole("button", { name: "Start" })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: "End" })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: "Before" })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: "After" })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: "Replace" })).toBeInTheDocument();
    });

    it("应该渲染目录标题输入框 / Should render TOC title input", () => {
      render(<TableOfContentsDebug />);

      expect(screen.getByLabelText(/目录标题/)).toBeInTheDocument();
    });

    it("应该渲染标题级别输入框 / Should render heading levels input", () => {
      render(<TableOfContentsDebug />);

      expect(screen.getByLabelText(/包含的标题级别/)).toBeInTheDocument();
    });

    it("应该渲染选项复选框 / Should render option checkboxes", () => {
      render(<TableOfContentsDebug />);

      expect(screen.getByLabelText("显示页码")).toBeInTheDocument();
      expect(screen.getByLabelText("页码右对齐")).toBeInTheDocument();
      expect(screen.getByLabelText("使用超链接")).toBeInTheDocument();
      expect(screen.getByLabelText("包含隐藏文本")).toBeInTheDocument();
    });
  });

  describe("插入目录功能测试 / Insert TOC Function Tests", () => {
    it("点击插入目录按钮应该调用insertTableOfContents函数 / Should call insertTableOfContents when insert button clicked", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(insertTableOfContents).toHaveBeenCalled();
      });
    });

    it("插入成功应该显示成功消息 / Should show success message on successful insert", async () => {
      const user = userEvent.setup();
      render(<TableOfContentsDebug />);

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(screen.getByText(/✓ 目录插入成功/)).toBeInTheDocument();
      });
    });

    it("应该能够更改插入位置 / Should be able to change insert location", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      // 点击End位置按钮 / Click End location button
      const endButton = screen.getByRole("button", { name: "End" });
      await user.click(endButton);

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(insertTableOfContents).toHaveBeenCalledWith(
          "End",
          expect.objectContaining({
            title: "目录",
            levels: [1, 2, 3],
          })
        );
      });
    });

    it("应该能够自定义目录标题 / Should be able to customize TOC title", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const titleInput = screen.getByLabelText(/目录标题/);
      await user.clear(titleInput);
      await user.type(titleInput, "Table of Contents");

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(insertTableOfContents).toHaveBeenCalledWith(
          "Start",
          expect.objectContaining({
            title: "Table of Contents",
          })
        );
      });
    });

    it("应该能够自定义标题级别 / Should be able to customize heading levels", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const levelsInput = screen.getByLabelText(/包含的标题级别/);
      await user.clear(levelsInput);
      await user.type(levelsInput, "1,2,3,4");

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(insertTableOfContents).toHaveBeenCalledWith(
          "Start",
          expect.objectContaining({
            levels: [1, 2, 3, 4],
          })
        );
      });
    });

    it("标题级别无效时应该显示错误 / Should show error when heading levels are invalid", async () => {
      const user = userEvent.setup();
      render(<TableOfContentsDebug />);

      const levelsInput = screen.getByLabelText(/包含的标题级别/);
      await user.clear(levelsInput);
      await user.type(levelsInput, "abc");

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(screen.getByText(/错误：请输入有效的标题级别/)).toBeInTheDocument();
      });
    });

    it("插入失败应该显示错误消息 / Should show error message on failed insert", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      vi.mocked(insertTableOfContents).mockResolvedValueOnce({
        success: false,
        error: "插入失败",
      });

      render(<TableOfContentsDebug />);

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(screen.getByText(/✗ 插入失败: 插入失败/)).toBeInTheDocument();
      });
    });
  });

  describe("更新目录功能测试 / Update TOC Function Tests", () => {
    it("点击更新目录按钮应该调用updateTableOfContents函数 / Should call updateTableOfContents when update button clicked", async () => {
      const user = userEvent.setup();
      const { updateTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const updateButtons = screen.getAllByRole("button", { name: "更新目录" });
      await user.click(updateButtons[0]);

      await waitFor(() => {
        expect(updateTableOfContents).toHaveBeenCalled();
      });
    });

    it("更新成功应该显示成功消息 / Should show success message on successful update", async () => {
      const user = userEvent.setup();
      render(<TableOfContentsDebug />);

      const updateButtons = screen.getAllByRole("button", { name: "更新目录" });
      await user.click(updateButtons[0]);

      await waitFor(() => {
        expect(screen.getByText(/✓ 目录更新成功/)).toBeInTheDocument();
      });
    });

    it("应该能够指定目录索引进行更新 / Should be able to specify TOC index for update", async () => {
      const user = userEvent.setup();
      const { updateTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const indexInput = screen.getByLabelText(/目录索引.*更新所有/);
      await user.type(indexInput, "0");

      const updateButtons = screen.getAllByRole("button", { name: "更新目录" });
      await user.click(updateButtons[0]);

      await waitFor(() => {
        expect(updateTableOfContents).toHaveBeenCalledWith(0);
      });
    });
  });

  describe("删除目录功能测试 / Delete TOC Function Tests", () => {
    it("点击删除目录按钮应该调用deleteTableOfContents函数 / Should call deleteTableOfContents when delete button clicked", async () => {
      const user = userEvent.setup();
      const { deleteTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const deleteButtons = screen.getAllByRole("button", { name: "删除目录" });
      await user.click(deleteButtons[0]);

      await waitFor(() => {
        expect(deleteTableOfContents).toHaveBeenCalled();
      });
    });

    it("删除成功应该显示成功消息 / Should show success message on successful delete", async () => {
      const user = userEvent.setup();
      const { getTableOfContentsList } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      // Mock getTableOfContentsList 返回空列表，避免覆盖删除成功消息 / Mock to return empty list
      vi.mocked(getTableOfContentsList).mockResolvedValueOnce({
        success: true,
        tocs: [],
      });

      render(<TableOfContentsDebug />);

      const deleteButtons = screen.getAllByRole("button", { name: "删除目录" });
      await user.click(deleteButtons[0]);

      await waitFor(() => {
        expect(screen.getByText(/✓ 找到 0 个目录/)).toBeInTheDocument();
      });
    });

    it("应该能够指定目录索引进行删除 / Should be able to specify TOC index for delete", async () => {
      const user = userEvent.setup();
      const { deleteTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const indexInput = screen.getByLabelText(/目录索引.*删除所有/);
      await user.type(indexInput, "0");

      const deleteButtons = screen.getAllByRole("button", { name: "删除目录" });
      await user.click(deleteButtons[0]);

      await waitFor(() => {
        expect(deleteTableOfContents).toHaveBeenCalledWith(0);
      });
    });
  });

  describe("获取目录列表功能测试 / Get TOC List Function Tests", () => {
    it("点击获取目录列表按钮应该调用getTableOfContentsList函数 / Should call getTableOfContentsList when get list button clicked", async () => {
      const user = userEvent.setup();
      const { getTableOfContentsList } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const getListButtons = screen.getAllByRole("button", { name: "获取目录列表" });
      await user.click(getListButtons[0]);

      await waitFor(() => {
        expect(getTableOfContentsList).toHaveBeenCalled();
      });
    });

    it("获取成功应该显示目录列表 / Should show TOC list on successful get", async () => {
      const user = userEvent.setup();
      render(<TableOfContentsDebug />);

      const getListButtons = screen.getAllByRole("button", { name: "获取目录列表" });
      await user.click(getListButtons[0]);

      await waitFor(() => {
        expect(screen.getByText(/✓ 找到 1 个目录/)).toBeInTheDocument();
        expect(screen.getByText(/目录 #0/)).toBeInTheDocument();
      });
    });

    it("获取失败应该显示错误消息 / Should show error message on failed get", async () => {
      const user = userEvent.setup();
      const { getTableOfContentsList } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      vi.mocked(getTableOfContentsList).mockResolvedValueOnce({
        success: false,
        error: "获取失败",
      });

      render(<TableOfContentsDebug />);

      const getListButtons = screen.getAllByRole("button", { name: "获取目录列表" });
      await user.click(getListButtons[0]);

      await waitFor(() => {
        expect(screen.getByText(/✗ 获取失败: 获取失败/)).toBeInTheDocument();
      });
    });
  });

  describe("选项复选框测试 / Option Checkbox Tests", () => {
    it("应该能够切换显示页码选项 / Should be able to toggle show page numbers option", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const showPageNumbersCheckbox = screen.getByLabelText("显示页码");
      await user.click(showPageNumbersCheckbox);

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(insertTableOfContents).toHaveBeenCalledWith(
          "Start",
          expect.objectContaining({
            showPageNumbers: false,
          })
        );
      });
    });

    it("应该能够切换页码右对齐选项 / Should be able to toggle right align page numbers option", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const rightAlignCheckbox = screen.getByLabelText("页码右对齐");
      await user.click(rightAlignCheckbox);

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(insertTableOfContents).toHaveBeenCalledWith(
          "Start",
          expect.objectContaining({
            rightAlignPageNumbers: false,
          })
        );
      });
    });

    it("应该能够切换使用超链接选项 / Should be able to toggle use hyperlinks option", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const useHyperlinksCheckbox = screen.getByLabelText("使用超链接");
      await user.click(useHyperlinksCheckbox);

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(insertTableOfContents).toHaveBeenCalledWith(
          "Start",
          expect.objectContaining({
            useHyperlinks: false,
          })
        );
      });
    });

    it("应该能够切换包含隐藏文本选项 / Should be able to toggle include hidden option", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      render(<TableOfContentsDebug />);

      const includeHiddenCheckbox = screen.getByLabelText("包含隐藏文本");
      await user.click(includeHiddenCheckbox);

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      await waitFor(() => {
        expect(insertTableOfContents).toHaveBeenCalledWith(
          "Start",
          expect.objectContaining({
            includeHidden: true,
          })
        );
      });
    });
  });

  describe("加载状态测试 / Loading State Tests", () => {
    it("插入时按钮应该被禁用 / Buttons should be disabled during insert", async () => {
      const user = userEvent.setup();
      const { insertTableOfContents } = await import(
        "../../../../src/word-tools/tableOfContents"
      );

      vi.mocked(insertTableOfContents).mockImplementation(
        () =>
          new Promise((resolve) =>
            setTimeout(
              () =>
                resolve({
                  success: true,
                  tocInfo: {
                    index: 0,
                    text: "目录",
                    entryCount: 0,
                    levels: [1, 2, 3],
                  },
                }),
              100
            )
          )
      );

      render(<TableOfContentsDebug />);

      const insertButtons = screen.getAllByRole("button", { name: "插入目录" });
      await user.click(insertButtons[0]);

      expect(insertButtons[0]).toBeDisabled();
    });
  });
});
