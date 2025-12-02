/**
 * 文件名: RangeContent.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: RangeContent 组件的单元测试
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { RangeContent } from "../../../src/taskpane/components/tools/RangeContent";
import * as wordTools from "../../../src/word-tools";

// 模拟 word-tools 模块 / Mock word-tools module
vi.mock("../../../src/word-tools", () => ({
  getRangeContent: vi.fn(),
}));

describe("RangeContent 组件 / RangeContent Component", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("应该正确渲染组件 / Should render component correctly", () => {
    render(<RangeContent />);

    expect(screen.getByText("范围定位方式")).toBeInTheDocument();
    expect(screen.getByText("获取选项")).toBeInTheDocument();
    expect(screen.getByText("获取范围内容")).toBeInTheDocument();
  });

  it("应该显示所有定位器选项 / Should display all locator options", async () => {
    const user = userEvent.setup();
    render(<RangeContent />);

    // 点击下拉框 / Click dropdown
    const dropdown = screen.getByRole("combobox");
    await user.click(dropdown);

    // 验证所有选项都存在 / Verify all options exist (使用 getAllByText 因为下拉框展开后会有重复)
    expect(screen.getAllByText("书签").length).toBeGreaterThan(0);
    expect(screen.getByText("标题")).toBeInTheDocument();
    expect(screen.getByText("段落索引")).toBeInTheDocument();
    expect(screen.getByText("节")).toBeInTheDocument();
    expect(screen.getByText("内容控件")).toBeInTheDocument();
  });

  it("应该在选择书签定位器时显示书签名称输入框 / Should show bookmark name input when bookmark locator selected", async () => {
    const user = userEvent.setup();
    render(<RangeContent />);

    // 默认应该显示书签输入 / Should show bookmark input by default
    expect(screen.getByPlaceholderText("输入书签名称")).toBeInTheDocument();
  });

  it("应该在选择标题定位器时显示标题相关输入框 / Should show heading inputs when heading locator selected", async () => {
    const user = userEvent.setup();
    render(<RangeContent />);

    // 选择标题定位器 / Select heading locator
    const dropdown = screen.getByRole("combobox");
    await user.click(dropdown);
    await user.click(screen.getByText("标题"));

    // 验证标题相关输入框 / Verify heading inputs
    expect(screen.getByPlaceholderText("输入标题文本")).toBeInTheDocument();
    expect(screen.getByPlaceholderText("输入标题级别")).toBeInTheDocument();
  });

  it("应该在选择段落索引定位器时显示段落索引输入框 / Should show paragraph index inputs when paragraph locator selected", async () => {
    const user = userEvent.setup();
    render(<RangeContent />);

    // 选择段落索引定位器 / Select paragraph locator
    const dropdown = screen.getByRole("combobox");
    await user.click(dropdown);
    await user.click(screen.getByText("段落索引"));

    // 验证段落索引输入框 / Verify paragraph index inputs
    expect(screen.getByPlaceholderText("0")).toBeInTheDocument();
    expect(screen.getByPlaceholderText("留空则只获取单个段落")).toBeInTheDocument();
  });

  it("应该能够切换获取选项 / Should be able to toggle get options", async () => {
    const user = userEvent.setup();
    render(<RangeContent />);

    // 查找所有开关 / Find all switches
    const switches = screen.getAllByRole("switch");

    // 默认应该都是打开的（除了详细元数据）/ Should all be on by default (except detailed metadata)
    expect(switches[0]).toBeChecked(); // 包含文本内容
    expect(switches[1]).toBeChecked(); // 包含图片信息
    expect(switches[2]).toBeChecked(); // 包含表格信息
    expect(switches[3]).toBeChecked(); // 包含内容控件
    expect(switches[4]).not.toBeChecked(); // 详细元数据

    // 切换详细元数据 / Toggle detailed metadata
    await user.click(switches[4]);
    expect(switches[4]).toBeChecked();
  });

  it("应该在书签名称为空时显示错误 / Should show error when bookmark name is empty", async () => {
    const user = userEvent.setup();
    render(<RangeContent />);

    // 点击获取按钮 / Click get button
    const getButton = screen.getByText("获取范围内容");
    await user.click(getButton);

    // 验证错误信息 / Verify error message
    expect(screen.getByText(/请输入书签名称/)).toBeInTheDocument();
  });

  it("应该能够成功获取范围内容 / Should successfully get range content", async () => {
    const user = userEvent.setup();
    const mockContentInfo = {
      text: "测试内容",
      elements: [
        {
          id: "para-1",
          type: "Paragraph" as const,
          text: "测试段落",
        },
      ],
      metadata: {
        isEmpty: false,
        characterCount: 4,
        paragraphCount: 1,
        tableCount: 0,
        imageCount: 0,
        locatorType: "bookmark",
      },
    };

    vi.mocked(wordTools.getRangeContent).mockResolvedValue(mockContentInfo);

    render(<RangeContent />);

    // 输入书签名称 / Input bookmark name
    const bookmarkInput = screen.getByPlaceholderText("输入书签名称");
    await user.type(bookmarkInput, "测试书签");

    // 点击获取按钮 / Click get button
    const getButton = screen.getByText("获取范围内容");
    await user.click(getButton);

    // 等待结果显示 / Wait for results to display
    await screen.findByText("范围统计信息");

    // 验证统计信息 / Verify statistics
    expect(screen.getByText("bookmark")).toBeInTheDocument();
    expect(screen.getByText("范围统计信息")).toBeInTheDocument();

    // 验证元素列表 / Verify elements list
    expect(screen.getByText(/范围内容元素/)).toBeInTheDocument();
    expect(screen.getByText(/段落 #1/)).toBeInTheDocument();
  });

  it("应该能够处理 API 错误 / Should handle API errors", async () => {
    const user = userEvent.setup();
    vi.mocked(wordTools.getRangeContent).mockRejectedValue(new Error("API 调用失败"));

    render(<RangeContent />);

    // 输入书签名称 / Input bookmark name
    const bookmarkInput = screen.getByPlaceholderText("输入书签名称");
    await user.type(bookmarkInput, "测试书签");

    // 点击获取按钮 / Click get button
    const getButton = screen.getByText("获取范围内容");
    await user.click(getButton);

    // 验证错误信息 / Verify error message
    await screen.findByText(/API 调用失败/);
  });

  it("应该能够清空结果 / Should be able to clear results", async () => {
    const user = userEvent.setup();
    const mockContentInfo = {
      text: "测试内容",
      elements: [],
      metadata: {
        isEmpty: false,
        characterCount: 4,
        paragraphCount: 0,
        tableCount: 0,
        imageCount: 0,
        locatorType: "bookmark",
      },
    };

    vi.mocked(wordTools.getRangeContent).mockResolvedValue(mockContentInfo);

    render(<RangeContent />);

    // 输入书签名称并获取 / Input bookmark name and get
    const bookmarkInput = screen.getByPlaceholderText("输入书签名称");
    await user.type(bookmarkInput, "测试书签");

    const getButton = screen.getByText("获取范围内容");
    await user.click(getButton);

    // 等待结果显示 / Wait for results
    await screen.findByText("范围统计信息");

    // 点击清空按钮 / Click clear button
    const clearButton = screen.getByText("清空");
    await user.click(clearButton);

    // 验证结果已清空 / Verify results cleared
    expect(screen.queryByText("范围统计信息")).not.toBeInTheDocument();
  });

  it("应该在加载时禁用按钮 / Should disable buttons during loading", async () => {
    const user = userEvent.setup();
    vi.mocked(wordTools.getRangeContent).mockImplementation(
      () => new Promise((resolve) => setTimeout(resolve, 1000))
    );

    render(<RangeContent />);

    // 输入书签名称 / Input bookmark name
    const bookmarkInput = screen.getByPlaceholderText("输入书签名称");
    await user.type(bookmarkInput, "测试书签");

    // 点击获取按钮 / Click get button
    const getButton = screen.getByText("获取范围内容");
    await user.click(getButton);

    // 验证按钮被禁用 / Verify buttons are disabled
    expect(getButton).toBeDisabled();
    expect(screen.getByText("清空")).toBeDisabled();
  });
});
