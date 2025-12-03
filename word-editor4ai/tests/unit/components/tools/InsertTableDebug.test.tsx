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
import { screen, waitFor, fireEvent, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { InsertTableDebug } from "../../../../src/taskpane/components/tools/InsertTableDebug";
import { mockWordRun, renderWithProviders } from "../../../utils/test-utils";

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
    renderWithProviders(<InsertTableDebug />);

    // 验证标签切换按钮存在（每个标签有1个按钮）
    const insertButtons = screen.getAllByRole("button", { name: "插入表格" });
    expect(insertButtons).toHaveLength(2); // 1个标签按钮 + 1个操作按钮
    
    expect(screen.getByRole("button", { name: "更新表格" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "查询表格" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "删除操作" })).toBeInTheDocument();

    // 验证默认表单字段 - 使用 placeholder 来定位输入框
    const inputs = screen.getAllByRole("spinbutton");
    // 第一个是行数输入框，第二个是列数输入框
    expect(inputs[0]).toHaveValue(3);
    expect(inputs[1]).toHaveValue(3);
  });

  test("2. 插入表格 - 成功场景", async () => {
    renderWithProviders(<InsertTableDebug />);

    await user.type(screen.getByRole("spinbutton", { name: /行数/i }), "{backspace}5");
    await user.type(screen.getByRole("spinbutton", { name: /列数/i }), "{backspace}4");

    const insertButtons = screen.getAllByRole("button", { name: "插入表格" });
    await user.click(insertButtons[1]);

    await waitFor(() => {
      expect(screen.getByText(/表格插入成功！索引: 0/)).toBeInTheDocument();
    });
  });

  test("3. 插入表格 - 失败场景", async () => {
    renderWithProviders(<InsertTableDebug />);

    // 获取所有"插入表格"按钮，点击第二个（操作按钮）
    const insertButtons = screen.getAllByRole("button", { name: "插入表格" });
    await user.click(insertButtons[1]);

    await waitFor(() => {
      expect(screen.getByText(/插入表格失败: 插入失败/)).toBeInTheDocument();
    });
  });

  test("4. 切换标签页功能", async () => {
    renderWithProviders(<InsertTableDebug />);

    // 切换到更新表格标签页
    await user.click(screen.getByText("更新表格"));

    // 等待内容渲染并验证 - 更新表格标签页有多个"表格索引"字段
    // 使用 findAllByPlaceholderText 来定位输入框
    await waitFor(() => {
      const tableIndexInputs = screen.getAllByPlaceholderText("0");
      // 更新表格标签页应该有多个索引为0的占位符字段
      expect(tableIndexInputs.length).toBeGreaterThan(0);
    });

    // 切换到查询表格标签页
    await user.click(screen.getByText("查询表格"));
    // 查询表格标签页有2个按钮：标签按钮和操作按钮
    const queryButtons = screen.getAllByRole("button", { name: "查询表格" });
    expect(queryButtons.length).toBeGreaterThan(0);
  });

  test("5. 查询表格信息", async () => {
    renderWithProviders(<InsertTableDebug />);
    await user.click(screen.getByText("查询表格"));

    // 使用 placeholder 来定位表格索引输入框
    const tableIndexInput = screen.getByPlaceholderText("0");
    await user.type(tableIndexInput, "0");
    const queryButtons = screen.getAllByRole("button", { name: "查询表格" });
    await user.click(queryButtons[1]);

    await waitFor(() => {
      expect(screen.getByText(/表格信息获取成功/)).toBeInTheDocument();
    });
  });

  test("6. 样式设置交互", async () => {
    renderWithProviders(<InsertTableDebug />);

    // 通过 label 文本定位表格样式下拉框 / Locate style dropdown by label text
    const styleLabel = screen.getByText("表格样式");
    const styleField = styleLabel.closest(".fui-Field");
    const styleDropdown = within(styleField as HTMLElement).getByRole("combobox");
    fireEvent.click(styleDropdown);

    await user.click(screen.getByText("网格2"));
    // Dropdown 显示的是 value 而不是 label，所以检查 value
    expect(styleDropdown).toHaveValue("GridTable2");

    // 使用 role 和 name 来定位 switch
    const switchElement = screen.getByRole("switch", { name: /首行特殊格式/i });
    // 初始状态是 checked (firstRow 默认为 true)
    expect(switchElement).toBeChecked();
    // 点击后变为 unchecked
    await user.click(switchElement);
    expect(switchElement).not.toBeChecked();
  });
});
