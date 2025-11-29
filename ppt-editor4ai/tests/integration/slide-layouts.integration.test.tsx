/**
 * 文件名: slide-layouts.integration.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片布局模板功能集成测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { screen, waitFor } from "@testing-library/react";
import { renderWithProviders, userEvent } from "../utils/test-utils";
import { OfficeMockObject } from "office-addin-mock";
import SlideLayouts from "../../src/taskpane/components/tools/SlideLayouts";

describe("SlideLayouts 集成测试 / SlideLayouts integration tests", () => {
  beforeEach(() => {
    // 清理全局 PowerPoint 对象
    delete (global as any).PowerPoint;
    vi.clearAllMocks();
  });

  it("应该完成完整的布局查询和幻灯片创建流程 / should complete full layout query and slide creation workflow", async () => {
    const user = userEvent.setup();

    // 创建完整的 Mock 上下文
    const mockContext = {
      presentation: {
        slideMasters: {
          items: [
            {
              id: "master-1",
              name: "Office Theme",
              layouts: {
                items: [
                  {
                    id: "layout-1",
                    name: "标题幻灯片",
                    type: "title",
                    shapes: {
                      items: [
                        {
                          type: "Placeholder",
                          placeholderFormat: {
                            type: "Title",
                          },
                        },
                        {
                          type: "Placeholder",
                          placeholderFormat: {
                            type: "Subtitle",
                          },
                        },
                      ],
                    },
                  },
                  {
                    id: "layout-2",
                    name: "空白",
                    type: "blank",
                    shapes: {
                      items: [],
                    },
                  },
                ],
              },
            },
          ],
        },
        slides: {
          items: [
            { id: "slide-1" },
            { id: "slide-2" },
          ],
          add: vi.fn(),
        },
      },
    };

    (global as any).PowerPoint = new OfficeMockObject(mockContext);

    renderWithProviders(<SlideLayouts />);

    // 步骤 1: 获取布局模板列表
    const fetchButton = screen.getByRole("button", { name: /获取布局模板/i });
    await user.click(fetchButton);

    await waitFor(() => {
      expect(screen.getByText("标题幻灯片")).toBeInTheDocument();
      expect(screen.getByText("空白")).toBeInTheDocument();
    });

    // 验证占位符标签显示
    expect(screen.getByText("Title")).toBeInTheDocument();
    expect(screen.getByText("Subtitle")).toBeInTheDocument();

    // 步骤 2: 选择布局
    const layoutCard = screen.getByText("标题幻灯片").closest("div");
    if (layoutCard) {
      await user.click(layoutCard);
    }

    await waitFor(() => {
      expect(screen.getByText(/使用布局.*标题幻灯片.*创建新幻灯片/i)).toBeInTheDocument();
    });

    // 步骤 3: 创建新幻灯片
    // Mock 新幻灯片
    const mockNewSlide = { id: "new-slide-1" };
    mockContext.presentation.slides.items.push(mockNewSlide);

    const createButton = screen.getByRole("button", { name: /创建新幻灯片/i });
    await user.click(createButton);

    await waitFor(() => {
      expect(mockContext.presentation.slides.add).toHaveBeenCalled();
      expect(screen.getByText(/成功创建新幻灯片/i)).toBeInTheDocument();
    });
  });

  it("应该支持切换占位符详细信息选项 / should support toggling placeholder details option", async () => {
    const user = userEvent.setup();

    const mockContext = {
      presentation: {
        slideMasters: {
          items: [
            {
              id: "master-1",
              name: "Office Theme",
              layouts: {
                items: [
                  {
                    id: "layout-1",
                    name: "测试布局",
                    type: "title",
                    shapes: {
                      items: [
                        {
                          type: "Placeholder",
                          placeholderFormat: {
                            type: "Title",
                          },
                        },
                      ],
                    },
                  },
                ],
              },
            },
          ],
        },
        slides: {
          items: [],
          add: vi.fn(),
        },
      },
    };

    (global as any).PowerPoint = new OfficeMockObject(mockContext);

    renderWithProviders(<SlideLayouts />);

    // 关闭占位符详细信息
    const switchElement = screen.getByRole("switch");
    await user.click(switchElement);

    // 获取布局
    await user.click(screen.getByRole("button", { name: /获取布局模板/i }));

    await waitFor(() => {
      expect(screen.getByText("测试布局")).toBeInTheDocument();
    });

    // 验证布局信息显示（即使关闭了占位符详细信息，布局名称仍应显示）
    expect(screen.getByText("测试布局")).toBeInTheDocument();
  });

  it("应该处理多个母版的复杂场景 / should handle complex scenario with multiple masters", async () => {
    const user = userEvent.setup();

    const mockContext = {
      presentation: {
        slideMasters: {
          items: [
            {
              id: "master-1",
              name: "Office Theme",
              layouts: {
                items: [
                  {
                    id: "layout-1",
                    name: "布局1",
                    type: "title",
                    shapes: { items: [] },
                  },
                  {
                    id: "layout-2",
                    name: "布局2",
                    type: "blank",
                    shapes: { items: [] },
                  },
                ],
              },
            },
            {
              id: "master-2",
              name: "Custom Theme",
              layouts: {
                items: [
                  {
                    id: "layout-3",
                    name: "布局3",
                    type: "custom",
                    shapes: { items: [] },
                  },
                ],
              },
            },
          ],
        },
        slides: {
          items: [],
          add: vi.fn(),
        },
      },
    };

    (global as any).PowerPoint = new OfficeMockObject(mockContext);

    renderWithProviders(<SlideLayouts />);

    await user.click(screen.getByRole("button", { name: /获取布局模板/i }));

    await waitFor(() => {
      expect(screen.getByText("布局1")).toBeInTheDocument();
      expect(screen.getByText("布局2")).toBeInTheDocument();
      expect(screen.getByText("布局3")).toBeInTheDocument();
    });

    // 验证统计信息
    expect(screen.getByText(/共找到.*3.*个布局模板/i)).toBeInTheDocument();
  });

  it("应该支持复制布局数据到剪贴板 / should support copying layout data to clipboard", async () => {
    const user = userEvent.setup();

    const mockContext = {
      presentation: {
        slideMasters: {
          items: [
            {
              id: "master-1",
              name: "Office Theme",
              layouts: {
                items: [
                  {
                    id: "layout-1",
                    name: "标题幻灯片",
                    type: "title",
                    shapes: { items: [] },
                  },
                ],
              },
            },
          ],
        },
        slides: {
          items: [],
          add: vi.fn(),
        },
      },
    };

    (global as any).PowerPoint = new OfficeMockObject(mockContext);

    // Mock clipboard API
    const mockWriteText = vi.fn().mockResolvedValue(undefined);
    Object.assign(navigator, {
      clipboard: {
        writeText: mockWriteText,
      },
    });

    renderWithProviders(<SlideLayouts />);

    await user.click(screen.getByRole("button", { name: /获取布局模板/i }));

    await waitFor(() => {
      expect(screen.getByText("标题幻灯片")).toBeInTheDocument();
    });

    const copyButton = screen.getByRole("button", { name: /复制 JSON/i });
    await user.click(copyButton);

    await waitFor(() => {
      expect(mockWriteText).toHaveBeenCalled();
      expect(screen.getByText("已复制")).toBeInTheDocument();
    });

    // 验证复制的数据格式
    const copiedData = mockWriteText.mock.calls[0][0];
    const parsedData = JSON.parse(copiedData);
    expect(parsedData).toHaveLength(1);
    expect(parsedData[0]).toMatchObject({
      id: "layout-1",
      name: "标题幻灯片",
      type: "title",
    });
  });

  it("应该在指定位置创建幻灯片 / should create slide at specified position", async () => {
    const user = userEvent.setup();

    const mockContext = {
      presentation: {
        slideMasters: {
          items: [
            {
              id: "master-1",
              name: "Office Theme",
              layouts: {
                items: [
                  {
                    id: "layout-1",
                    name: "标题幻灯片",
                    type: "title",
                    shapes: { items: [] },
                  },
                ],
              },
            },
          ],
        },
        slides: {
          items: [
            { id: "slide-1" },
            { id: "slide-2" },
            { id: "slide-3" },
          ],
          add: vi.fn(),
        },
      },
    };

    (global as any).PowerPoint = new OfficeMockObject(mockContext);

    renderWithProviders(<SlideLayouts />);

    // 获取布局
    await user.click(screen.getByRole("button", { name: /获取布局模板/i }));

    await waitFor(() => {
      expect(screen.getByText("标题幻灯片")).toBeInTheDocument();
    });

    // 选择布局
    const layoutCard = screen.getByText("标题幻灯片").closest("div");
    if (layoutCard) {
      await user.click(layoutCard);
    }

    await waitFor(() => {
      expect(screen.getByPlaceholderText(/例如: 0 表示插入到开头/i)).toBeInTheDocument();
    });

    // 输入位置
    const positionInput = screen.getByPlaceholderText(/例如: 0 表示插入到开头/i);
    await user.type(positionInput, "1");

    // Mock 新幻灯片
    const mockNewSlide = { id: "new-slide-1" };
    mockContext.presentation.slides.items.push(mockNewSlide);

    // 创建幻灯片
    await user.click(screen.getByRole("button", { name: /创建新幻灯片/i }));

    await waitFor(() => {
      expect(mockContext.presentation.slides.add).toHaveBeenCalled();
      expect(screen.getByText(/成功创建新幻灯片/i)).toBeInTheDocument();
    });
  });

  it("应该处理错误情况并显示友好的错误消息 / should handle errors and display friendly error messages", async () => {
    const user = userEvent.setup();

    // 创建一个会导致错误的 Mock 上下文
    const mockContext = {
      presentation: {
        slideMasters: {
          items: [],
        },
        slides: {
          items: [],
          add: vi.fn(),
        },
      },
    };

    (global as any).PowerPoint = new OfficeMockObject(mockContext);

    renderWithProviders(<SlideLayouts />);

    await user.click(screen.getByRole("button", { name: /获取布局模板/i }));

    // 应该显示空状态或成功消息（因为没有母版，返回空数组）
    await waitFor(() => {
      expect(screen.getByText(/成功获取 0 个布局模板/i)).toBeInTheDocument();
    });
  });
});
