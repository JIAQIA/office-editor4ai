/**
 * 文件名: slideLayouts.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片布局模板工具单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import {
  getAvailableSlideLayouts,
  createSlideWithLayout,
  getLayoutDescription,
  type SlideLayoutTemplate,
} from "../../../src/ppt-tools";

describe("slideLayouts 工具函数测试 / slideLayouts utility functions tests", () => {
  beforeEach(() => {
    // 清理全局 PowerPoint 对象
    delete (global as any).PowerPoint;
  });

  /**
   * 为对象添加 load 方法的 mock
   */
  const addLoadMethod = (obj: any) => {
    obj.load = vi.fn(() => obj);
    return obj;
  };

  /**
   * 创建 PowerPoint.run 的 mock 实现
   */
  const createPowerPointRunMock = (mockContext: any) => {
    // 为 context 添加 sync 方法
    mockContext.sync = vi.fn(async () => {});

    // 递归为所有对象添加 load 方法
    const addLoadToObjects = (obj: any) => {
      if (!obj || typeof obj !== "object") return;

      addLoadMethod(obj);

      // 遍历对象的所有属性
      for (const key in obj) {
        if (obj.hasOwnProperty(key) && typeof obj[key] === "object" && obj[key] !== null) {
          addLoadToObjects(obj[key]);
        }
      }
    };

    addLoadToObjects(mockContext.presentation);

    return async (callback: (context: any) => Promise<void>) => {
      await callback(mockContext);
    };
  };

  describe("getAvailableSlideLayouts", () => {
    it("应该成功获取布局模板列表 / should successfully get layout templates list", async () => {
      // 创建 Mock 数据
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
                      name: "标题和内容",
                      type: "titleAndContent",
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
                              type: "Body",
                            },
                          },
                        ],
                      },
                    },
                    {
                      id: "layout-3",
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
        },
      };

      (global as any).PowerPoint = {
        run: createPowerPointRunMock(mockContext),
      };

      const layouts = await getAvailableSlideLayouts({ includePlaceholders: true });

      expect(layouts).toHaveLength(3);
      expect(layouts[0]).toMatchObject({
        id: "layout-1",
        name: "标题幻灯片",
        type: "title",
        placeholderCount: 2,
        placeholderTypes: ["Title", "Subtitle"],
      });
      expect(layouts[1]).toMatchObject({
        id: "layout-2",
        name: "标题和内容",
        type: "titleAndContent",
        placeholderCount: 2,
      });
      expect(layouts[2]).toMatchObject({
        id: "layout-3",
        name: "空白",
        type: "blank",
        placeholderCount: 0,
        placeholderTypes: [],
      });
    });

    it("应该在不包含占位符信息时正确工作 / should work correctly without placeholder info", async () => {
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
                        items: [],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
      };

      (global as any).PowerPoint = {
        run: createPowerPointRunMock(mockContext),
      };

      const layouts = await getAvailableSlideLayouts({ includePlaceholders: false });

      expect(layouts).toHaveLength(1);
      expect(layouts[0]).toMatchObject({
        id: "layout-1",
        name: "标题幻灯片",
        type: "title",
        placeholderCount: 0,
        placeholderTypes: [],
      });
    });

    it("应该处理没有母版的情况 / should handle case with no slide masters", async () => {
      const mockContext = {
        presentation: {
          slideMasters: {
            items: [],
          },
        },
      };

      (global as any).PowerPoint = {
        run: createPowerPointRunMock(mockContext),
      };

      const layouts = await getAvailableSlideLayouts();

      expect(layouts).toHaveLength(0);
    });

    it("应该处理多个母版的情况 / should handle multiple slide masters", async () => {
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
                  ],
                },
              },
              {
                id: "master-2",
                name: "Custom Theme",
                layouts: {
                  items: [
                    {
                      id: "layout-2",
                      name: "布局2",
                      type: "blank",
                      shapes: { items: [] },
                    },
                  ],
                },
              },
            ],
          },
        },
      };

      (global as any).PowerPoint = {
        run: createPowerPointRunMock(mockContext),
      };

      const layouts = await getAvailableSlideLayouts({ includePlaceholders: false });

      expect(layouts).toHaveLength(2);
      expect(layouts[0].name).toBe("布局1");
      expect(layouts[1].name).toBe("布局2");
    });

    it("应该处理占位符加载失败的情况 / should handle placeholder loading failures", async () => {
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
                            placeholderFormat: null, // 模拟加载失败
                          },
                        ],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
      };

      (global as any).PowerPoint = {
        run: createPowerPointRunMock(mockContext),
      };

      const layouts = await getAvailableSlideLayouts({ includePlaceholders: true });

      expect(layouts).toHaveLength(1);
      // 应该继续处理，即使某些占位符加载失败
      expect(layouts[0].id).toBe("layout-1");
    });
  });

  describe("createSlideWithLayout", () => {
    it("应该成功创建新幻灯片 / should successfully create new slide", async () => {
      const mockNewSlide = { id: "new-slide-1" };
      addLoadMethod(mockNewSlide);

      const mockContext = {
        presentation: {
          slideMasters: {
            items: [
              {
                id: "master-1",
                layouts: {
                  items: [
                    {
                      id: "layout-1",
                      name: "标题幻灯片",
                    },
                  ],
                },
              },
            ],
          },
          slides: {
            items: [{ id: "slide-1" }, { id: "slide-2" }],
            add: vi.fn(() => {
              mockContext.presentation.slides.items.push(mockNewSlide);
            }),
          },
        },
      };

      (global as any).PowerPoint = {
        run: createPowerPointRunMock(mockContext),
      };

      const newSlideId = await createSlideWithLayout("layout-1");

      expect(newSlideId).toBe("new-slide-1");
      expect(mockContext.presentation.slides.add).toHaveBeenCalled();
    });

    it("应该在指定位置创建幻灯片 / should create slide at specified position", async () => {
      const mockNewSlide = { id: "new-slide-1" };
      addLoadMethod(mockNewSlide);

      const mockContext = {
        presentation: {
          slideMasters: {
            items: [
              {
                id: "master-1",
                layouts: {
                  items: [
                    {
                      id: "layout-1",
                      name: "标题幻灯片",
                    },
                  ],
                },
              },
            ],
          },
          slides: {
            items: [{ id: "slide-1" }, { id: "slide-2" }],
            add: vi.fn(() => {
              mockContext.presentation.slides.items.push(mockNewSlide);
            }),
          },
        },
      };

      (global as any).PowerPoint = {
        run: createPowerPointRunMock(mockContext),
      };

      const newSlideId = await createSlideWithLayout("layout-1", 0);

      expect(newSlideId).toBe("new-slide-1");
      expect(mockContext.presentation.slides.add).toHaveBeenCalled();
    });

    it("应该在布局ID不存在时抛出错误 / should throw error when layout ID not found", async () => {
      const mockContext = {
        presentation: {
          slideMasters: {
            items: [
              {
                id: "master-1",
                layouts: {
                  items: [
                    {
                      id: "layout-1",
                      name: "标题幻灯片",
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

      (global as any).PowerPoint = {
        run: createPowerPointRunMock(mockContext),
      };

      await expect(createSlideWithLayout("non-existent-layout")).rejects.toThrow(
        "未找到布局ID: non-existent-layout"
      );
    });

    it("应该处理幻灯片创建失败的情况 / should handle slide creation failure", async () => {
      const mockContext = {
        presentation: {
          slideMasters: {
            items: [
              {
                id: "master-1",
                layouts: {
                  items: [
                    {
                      id: "layout-1",
                      name: "标题幻灯片",
                    },
                  ],
                },
              },
            ],
          },
          slides: {
            items: [],
            add: vi.fn(), // 不添加新幻灯片，模拟创建失败
          },
        },
      };

      (global as any).PowerPoint = {
        run: createPowerPointRunMock(mockContext),
      };

      await expect(createSlideWithLayout("layout-1")).rejects.toThrow(
        "创建幻灯片失败：无法获取新创建的幻灯片"
      );
    });
  });

  describe("getLayoutDescription", () => {
    it("应该生成完整的布局描述 / should generate complete layout description", () => {
      const layout: SlideLayoutTemplate = {
        id: "layout-1",
        name: "标题幻灯片",
        type: "title",
        placeholderCount: 2,
        placeholderTypes: ["Title", "Subtitle"],
        isCustom: false,
      };

      const description = getLayoutDescription(layout);

      expect(description).toContain("类型: title");
      expect(description).toContain("2 个占位符");
      expect(description).toContain("Title, Subtitle");
    });

    it("应该处理无占位符的布局 / should handle layout without placeholders", () => {
      const layout: SlideLayoutTemplate = {
        id: "layout-1",
        name: "空白",
        type: "blank",
        placeholderCount: 0,
        placeholderTypes: [],
        isCustom: false,
      };

      const description = getLayoutDescription(layout);

      expect(description).toContain("类型: blank");
      expect(description).toContain("无占位符");
    });

    it("应该处理未知类型的布局 / should handle layout with unknown type", () => {
      const layout: SlideLayoutTemplate = {
        id: "layout-1",
        name: "自定义布局",
        type: "unknown",
        placeholderCount: 1,
        placeholderTypes: ["Body"],
        isCustom: true,
      };

      const description = getLayoutDescription(layout);

      expect(description).not.toContain("类型: unknown");
      expect(description).toContain("1 个占位符");
    });

    it("应该限制占位符类型显示数量 / should limit placeholder types display", () => {
      const layout: SlideLayoutTemplate = {
        id: "layout-1",
        name: "复杂布局",
        type: "custom",
        placeholderCount: 5,
        placeholderTypes: ["Title", "Body", "Picture", "Chart", "Table"],
        isCustom: true,
      };

      const description = getLayoutDescription(layout);

      expect(description).toContain("5 个占位符");
      // 不应该显示所有占位符类型（超过3个）
      expect(description).not.toContain("Title, Body, Picture, Chart, Table");
    });

    it("应该去重占位符类型 / should deduplicate placeholder types", () => {
      const layout: SlideLayoutTemplate = {
        id: "layout-1",
        name: "重复占位符布局",
        type: "custom",
        placeholderCount: 3,
        placeholderTypes: ["Body", "Body", "Picture"],
        isCustom: false,
      };

      const description = getLayoutDescription(layout);

      expect(description).toContain("3 个占位符");
      expect(description).toContain("Body, Picture");
      // 不应该包含重复的 Body
      expect(description).not.toContain("Body, Body");
    });
  });
});
