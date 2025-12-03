/**
 * 文件名: InsertTableDebug.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/react, @testing-library/user-event, @fluentui/react-components
 * 描述: InsertTableDebug组件的Vitest单元测试
 */

import { describe, test, expect, vi, beforeEach } from "vitest";
import { render, screen, waitFor, fireEvent } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { InsertTableDebug } from "../../../../src/taskpane/components/tools/InsertTableDebug";
import { mockWordRun } from "../../../utils/test-utils";

// Mock Word.js API
vi.mock("../../../../src/word-tools/table", () => ({
  insertTable: vi
    .fn()
    .mockResolvedValueOnce({ success: true, tableIndex: 0 }) // 第一次成功
    .mockRejectedValueOnce(new Error("插入失败")) // 第二次失败
    .mockResolvedValue({ success: true, tableIndex: 1 }), // 后续默认

  updateTable: vi.fn().mockResolvedValue({ success: true }),
  getTableInfo: vi.fn().mockResolvedValue({
    index: 0,
    rowCount: 3,
    columnCount: 3,
    data: [
      ["1", "2", "3"],
      ["4", "5", "6"],
      ["7", "8", "9"],
    ],
  }),
  getAllTablesInfo: vi.fn().mockResolvedValue([{ index: 0, rowCount: 3, columnCount: 3 }]),
}));

describe("InsertTableDebug 组件测试", () => {
  const user = userEvent.setup();

  beforeEach(() => {
    mockWordRun();
    vi.clearAllMocks();
  });

  test("1. 初始渲染正确", async () => {
    render(<InsertTableDebug />);

    // 验证标签页
    expect(screen.getByRole("tab", { name: "插入表格" })).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "更新表格" })).toBeInTheDocument();

    // 验证默认表单字段
    expect(screen.getByLabelText("行数")).toHaveValue("3");
    expect(screen.getByLabelText("列数")).toHaveValue("3");
  });

  test("2. 插入表格 - 成功场景", async () => {
    render(<InsertTableDebug />);

    await user.type(screen.getByLabelText("行数"), "{backspace}5");
    await user.type(screen.getByLabelText("列数"), "{backspace}4");
    await user.click(screen.getByRole("button", { name: "插入表格" }));

    await waitFor(() => {
      expect(screen.getByText(/表格插入成功！索引: 0/)).toBeInTheDocument();
    });
  });

  test("3. 插入表格 - 失败场景", async () => {
    render(<InsertTableDebug />);

    await user.click(screen.getByRole("button", { name: "插入表格" }));

    await waitFor(() => {
      expect(screen.getByText(/插入表格失败: 插入失败/)).toBeInTheDocument();
    });
  });

  test("4. 切换标签页功能", async () => {
    render(<InsertTableDebug />);

    // 切换到更新表格标签页
    await user.click(screen.getByText("更新表格"));
    expect(screen.getByLabelText("表格索引")).toBeInTheDocument();

    // 切换到查询表格标签页
    await user.click(screen.getByText("查询表格"));
    expect(screen.getByRole("button", { name: "查询表格信息" })).toBeInTheDocument();
  });

  test("5. 查询表格信息", async () => {
    render(<InsertTableDebug />);
    await user.click(screen.getByText("查询表格"));

    await user.type(screen.getByLabelText("表格索引"), "0");
    await user.click(screen.getByRole("button", { name: "查询表格信息" }));

    await waitFor(() => {
      expect(screen.getByText(/表格信息获取成功/)).toBeInTheDocument();
    });
  });

  test("6. 样式设置交互", async () => {
    render(<InsertTableDebug />);

    const styleDropdown = screen.getByLabelText("表格样式");
    fireEvent.click(styleDropdown);

    await user.click(screen.getByText("网格2"));
    expect(styleDropdown).toHaveTextContent("网格2");

    const switchElement = screen.getByLabelText("首行特殊格式");
    await user.click(switchElement);
    expect(switchElement).toBeChecked();
  });
});
