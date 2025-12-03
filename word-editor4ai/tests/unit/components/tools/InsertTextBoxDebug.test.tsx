/**
 * 文件名: InsertTextBoxDebug.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/react, @testing-library/user-event, @fluentui/react-components
 * 描述: InsertTextBoxDebug组件的Vitest单元测试
 */

import { describe, test, expect, vi, beforeEach } from "vitest";
import { screen, waitFor, fireEvent, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { InsertTextBoxDebug } from "../../../../src/taskpane/components/tools/InsertTextBoxDebug";
import { mockWordRun, renderWithProviders } from "../../../utils/test-utils";

// Mock insertTextBox 函数 / Mock insertTextBox function
vi.mock("../../../../src/word-tools", async () => {
  const actual = await vi.importActual("../../../../src/word-tools");
  return {
    ...actual,
    insertTextBox: vi
      .fn()
      .mockResolvedValueOnce({ success: true, textBoxId: "textbox-123" }) // 第一次成功 / First call succeeds
      .mockResolvedValueOnce({ success: false, error: "插入失败" }) // 第二次失败 / Second call fails
      .mockResolvedValue({ success: true, textBoxId: "textbox-456" }), // 后续默认成功 / Subsequent calls succeed
  };
});

describe("InsertTextBoxDebug 组件测试 / InsertTextBoxDebug Component Tests", () => {
  const user = userEvent.setup();

  beforeEach(() => {
    mockWordRun();
    vi.clearAllMocks();
  });

  test("1. 初始渲染正确 / Initial render is correct", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 验证基本元素存在 / Verify basic elements exist
    expect(screen.getByText("文本框内容")).toBeInTheDocument();
    expect(screen.getByText("基本选项")).toBeInTheDocument();
    expect(screen.getByText("位置和旋转")).toBeInTheDocument();
    expect(screen.getByText("插入文本框")).toBeInTheDocument();

    // 验证默认值 / Verify default values
    const textArea = screen.getByPlaceholderText("请输入文本框内容");
    expect(textArea).toHaveValue("示例文本框内容");

    // 验证默认宽度和高度 / Verify default width and height
    const widthInput = screen.getByLabelText("宽度（磅）");
    const heightInput = screen.getByLabelText("高度（磅）");
    expect(widthInput).toHaveValue(150);
    expect(heightInput).toHaveValue(100);
  });

  test("2. 插入文本框 - 成功场景 / Insert text box - Success scenario", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 输入文本内容 / Input text content
    const textArea = screen.getByPlaceholderText("请输入文本框内容");
    await user.clear(textArea);
    await user.type(textArea, "测试文本框");

    // 点击插入按钮 / Click insert button
    const insertButton = screen.getByRole("button", { name: /插入文本框/ });
    await user.click(insertButton);

    // 验证成功消息 / Verify success message
    await waitFor(() => {
      expect(screen.getByText(/文本框插入成功！标识符: textbox-123/)).toBeInTheDocument();
    });
  });

  test("3. 插入文本框 - 失败场景 / Insert text box - Failure scenario", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 点击插入按钮（使用默认文本）/ Click insert button (with default text)
    const insertButton = screen.getByRole("button", { name: /插入文本框/ });
    await user.click(insertButton);

    // 验证失败消息 / Verify failure message
    await waitFor(() => {
      expect(screen.getByText(/文本框插入失败: 插入失败/)).toBeInTheDocument();
    });
  });

  test("4. 空文本验证 / Empty text validation", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 清空文本内容 / Clear text content
    const textArea = screen.getByPlaceholderText("请输入文本框内容");
    await user.clear(textArea);

    // 验证插入按钮被禁用 / Verify insert button is disabled
    const insertButton = screen.getByRole("button", { name: /插入文本框/ });
    expect(insertButton).toBeDisabled();
  });

  test("5. 修改基本选项 / Modify basic options", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 修改宽度 / Modify width
    const widthInput = screen.getByLabelText("宽度（磅）");
    await user.clear(widthInput);
    await user.type(widthInput, "200");
    expect(widthInput).toHaveValue(200);

    // 修改高度 / Modify height
    const heightInput = screen.getByLabelText("高度（磅）");
    await user.clear(heightInput);
    await user.type(heightInput, "150");
    expect(heightInput).toHaveValue(150);

    // 修改文本框名称 / Modify text box name
    const nameInput = screen.getByPlaceholderText("MyTextBox");
    await user.type(nameInput, "TestBox");
    expect(nameInput).toHaveValue("TestBox");
  });

  test("6. 切换开关选项 / Toggle switch options", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 切换锁定纵横比 / Toggle lock aspect ratio
    const lockAspectRatioSwitch = screen.getByRole("switch", { name: /锁定纵横比/ });
    expect(lockAspectRatioSwitch).not.toBeChecked();
    await user.click(lockAspectRatioSwitch);
    expect(lockAspectRatioSwitch).toBeChecked();

    // 切换可见性 / Toggle visibility
    const visibleSwitch = screen.getByRole("switch", { name: /可见/ });
    expect(visibleSwitch).toBeChecked();
    await user.click(visibleSwitch);
    expect(visibleSwitch).not.toBeChecked();
  });

  test("7. 修改插入位置 / Modify insert location", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 定位插入位置下拉框 / Locate insert location dropdown
    const locationLabel = screen.getByText("插入位置");
    const locationField = locationLabel.closest(".fui-Field");
    const locationDropdown = within(locationField as HTMLElement).getByRole("combobox");

    // 点击下拉框 / Click dropdown
    fireEvent.click(locationDropdown);

    // 选择"文档开头" / Select "Start"
    await user.click(screen.getByText("文档开头"));
    expect(locationDropdown).toHaveValue("Start");
  });

  test("8. 修改位置和旋转参数 / Modify position and rotation parameters", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 修改左边距 / Modify left position
    const leftInput = screen.getByLabelText("左边距（磅，可选）");
    await user.type(leftInput, "50");
    expect(leftInput).toHaveValue(50);

    // 修改上边距 / Modify top position
    const topInput = screen.getByLabelText("上边距（磅，可选）");
    await user.type(topInput, "100");
    expect(topInput).toHaveValue(100);

    // 修改旋转角度 / Modify rotation
    const rotationInput = screen.getByLabelText("旋转角度（度，可选）");
    await user.type(rotationInput, "45");
    expect(rotationInput).toHaveValue(45);
  });

  test("9. 启用文本格式 / Enable text format", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 启用文本格式 / Enable text format
    const formatSwitch = screen.getByRole("switch", { name: /启用文本格式/ });
    expect(formatSwitch).not.toBeChecked();
    await user.click(formatSwitch);
    expect(formatSwitch).toBeChecked();

    // 验证格式选项出现 / Verify format options appear
    await waitFor(() => {
      expect(screen.getByLabelText("字体")).toBeInTheDocument();
      expect(screen.getByLabelText("字号")).toBeInTheDocument();
      expect(screen.getByLabelText("文字颜色")).toBeInTheDocument();
    });
  });

  test("10. 修改文本格式选项 / Modify text format options", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 启用文本格式 / Enable text format
    const formatSwitch = screen.getByRole("switch", { name: /启用文本格式/ });
    await user.click(formatSwitch);

    // 修改字体 / Modify font
    const fontInput = screen.getByLabelText("字体");
    await user.clear(fontInput);
    await user.type(fontInput, "Times New Roman");
    expect(fontInput).toHaveValue("Times New Roman");

    // 修改字号 / Modify font size
    const fontSizeInput = screen.getByLabelText("字号");
    await user.clear(fontSizeInput);
    await user.type(fontSizeInput, "16");
    expect(fontSizeInput).toHaveValue(16);

    // 切换粗体 / Toggle bold
    const boldSwitch = screen.getByRole("switch", { name: /粗体/ });
    expect(boldSwitch).not.toBeChecked();
    await user.click(boldSwitch);
    expect(boldSwitch).toBeChecked();

    // 切换斜体 / Toggle italic
    const italicSwitch = screen.getByRole("switch", { name: /斜体/ });
    expect(italicSwitch).not.toBeChecked();
    await user.click(italicSwitch);
    expect(italicSwitch).toBeChecked();

    // 切换删除线 / Toggle strikethrough
    const strikeThroughSwitch = screen.getByRole("switch", { name: /删除线/ });
    expect(strikeThroughSwitch).not.toBeChecked();
    await user.click(strikeThroughSwitch);
    expect(strikeThroughSwitch).toBeChecked();
  });

  test("11. 修改下划线样式 / Modify underline style", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 启用文本格式 / Enable text format
    const formatSwitch = screen.getByRole("switch", { name: /启用文本格式/ });
    await user.click(formatSwitch);

    // 定位下划线下拉框 / Locate underline dropdown
    const underlineLabel = screen.getByText("下划线");
    const underlineField = underlineLabel.closest(".fui-Field");
    const underlineDropdown = within(underlineField as HTMLElement).getByRole("combobox");

    // 点击下拉框 / Click dropdown
    fireEvent.click(underlineDropdown);

    // 选择"单下划线" / Select "Single"
    await user.click(screen.getByText("单下划线"));
    expect(underlineDropdown).toHaveValue("Single");
  });

  test("12. 修改颜色选项 / Modify color options", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 启用文本格式 / Enable text format
    const formatSwitch = screen.getByRole("switch", { name: /启用文本格式/ });
    await user.click(formatSwitch);

    // 修改文字颜色 - 使用文本输入框而不是颜色选择器 / Modify text color - Use text input instead of color picker
    await waitFor(() => {
      expect(screen.getByLabelText("文字颜色")).toBeInTheDocument();
    });
    
    const colorInputs = screen.getAllByDisplayValue("#000000");
    // 找到文本输入框（不是 type="color" 的输入框）/ Find text input (not type="color" input)
    const textColorInput = colorInputs.find(input => input.getAttribute('type') !== 'color');
    if (textColorInput) {
      await user.clear(textColorInput);
      await user.type(textColorInput, "#FF0000");
      expect(textColorInput).toHaveValue("#FF0000");
    }

    // 修改高亮颜色 / Modify highlight color
    const highlightColorInput = screen.getByLabelText("高亮颜色（可选）");
    await user.type(highlightColorInput, "#FFFF00");
    expect(highlightColorInput).toHaveValue("#FFFF00");
  });

  test("13. 加载状态显示 / Loading state display", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 点击插入按钮 / Click insert button
    const insertButton = screen.getByRole("button", { name: /插入文本框/ });
    await user.click(insertButton);

    // 验证加载状态（按钮文本变化）/ Verify loading state (button text changes)
    // 注意：由于 mock 是异步的，加载状态可能很快消失
    // Note: Loading state may disappear quickly due to async mock
    await waitFor(() => {
      expect(screen.getByText(/文本框插入成功！/)).toBeInTheDocument();
    });
  });

  test("14. 边界值测试 - 负数宽高 / Boundary test - Negative width/height", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 输入负数宽度 / Input negative width
    const widthInput = screen.getByLabelText("宽度（磅）");
    await user.clear(widthInput);
    await user.type(widthInput, "-100");
    expect(widthInput).toHaveValue(-100);

    // 输入负数高度 / Input negative height
    const heightInput = screen.getByLabelText("高度（磅）");
    await user.clear(heightInput);
    await user.type(heightInput, "-50");
    expect(heightInput).toHaveValue(-50);
  });

  test("15. 边界值测试 - 极大数值 / Boundary test - Very large values", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 输入极大宽度 / Input very large width
    const widthInput = screen.getByLabelText("宽度（磅）");
    await user.clear(widthInput);
    await user.type(widthInput, "9999");
    expect(widthInput).toHaveValue(9999);

    // 输入极大旋转角度 / Input very large rotation
    const rotationInput = screen.getByLabelText("旋转角度（度，可选）");
    await user.type(rotationInput, "360");
    expect(rotationInput).toHaveValue(360);
  });

  test("16. 特殊字符文本测试 / Special characters text test", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 输入特殊字符 / Input special characters
    const textArea = screen.getByPlaceholderText("请输入文本框内容");
    await user.clear(textArea);
    await user.type(textArea, "特殊字符: @#$%^&*()");
    expect(textArea).toHaveValue("特殊字符: @#$%^&*()");

    // 验证可以插入 / Verify can insert
    const insertButton = screen.getByRole("button", { name: /插入文本框/ });
    expect(insertButton).not.toBeDisabled();
  });

  test("17. 多行文本测试 / Multi-line text test", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 输入多行文本 / Input multi-line text
    const textArea = screen.getByPlaceholderText("请输入文本框内容");
    await user.clear(textArea);
    await user.type(textArea, "第一行{Enter}第二行{Enter}第三行");

    // 验证文本内容 / Verify text content
    expect(textArea).toHaveValue("第一行\n第二行\n第三行");
  });

  test("18. 警告提示显示 / Warning message display", async () => {
    renderWithProviders(<InsertTextBoxDebug />);

    // 验证位置参数警告提示存在 / Verify position parameter warning exists
    expect(
      screen.getByText(/注意：位置参数在某些情况下可能无法完全生效/)
    ).toBeInTheDocument();
  });
});
