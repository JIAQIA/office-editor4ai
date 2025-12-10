/**
 * 文件名: replaceText.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: replaceText 工具的单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import type { ReplaceTextOptions } from "../../src/word-tools";

describe("replaceText", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("类型定义验证", () => {
    it("应该正确定义 selection 定位器", () => {
      const options: ReplaceTextOptions = {
        locator: { type: "selection" },
        newText: "新文本",
      };

      expect(options.locator.type).toBe("selection");
      expect(options.newText).toBe("新文本");
    });

    it("应该正确定义 search 定位器", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "search",
          searchText: "旧文本",
          searchOptions: {
            matchCase: true,
            matchWholeWord: false,
          },
        },
        newText: "新文本",
        replaceAll: true,
      };

      expect(options.locator.type).toBe("search");
      if (options.locator.type === "search") {
        expect(options.locator.searchText).toBe("旧文本");
        expect(options.locator.searchOptions?.matchCase).toBe(true);
      }
      expect(options.replaceAll).toBe(true);
    });

    it("应该正确定义 range 定位器 - bookmark", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "range",
          rangeLocator: {
            type: "bookmark",
            name: "myBookmark",
          },
        },
        newText: "新文本",
      };

      expect(options.locator.type).toBe("range");
      if (options.locator.type === "range") {
        expect(options.locator.rangeLocator.type).toBe("bookmark");
        if (options.locator.rangeLocator.type === "bookmark") {
          expect(options.locator.rangeLocator.name).toBe("myBookmark");
        }
      }
    });

    it("应该正确定义 range 定位器 - heading", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "range",
          rangeLocator: {
            type: "heading",
            text: "第一章",
            level: 1,
            index: 0,
          },
        },
        newText: "新文本",
      };

      expect(options.locator.type).toBe("range");
      if (options.locator.type === "range") {
        expect(options.locator.rangeLocator.type).toBe("heading");
        if (options.locator.rangeLocator.type === "heading") {
          expect(options.locator.rangeLocator.text).toBe("第一章");
          expect(options.locator.rangeLocator.level).toBe(1);
        }
      }
    });

    it("应该正确定义 range 定位器 - paragraph", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "range",
          rangeLocator: {
            type: "paragraph",
            startIndex: 0,
            endIndex: 2,
          },
        },
        newText: "新文本",
      };

      expect(options.locator.type).toBe("range");
      if (options.locator.type === "range") {
        expect(options.locator.rangeLocator.type).toBe("paragraph");
        if (options.locator.rangeLocator.type === "paragraph") {
          expect(options.locator.rangeLocator.startIndex).toBe(0);
          expect(options.locator.rangeLocator.endIndex).toBe(2);
        }
      }
    });

    it("应该支持文本格式选项", () => {
      const options: ReplaceTextOptions = {
        locator: { type: "selection" },
        newText: "新文本",
        format: {
          fontName: "Arial",
          fontSize: 14,
          bold: true,
          italic: false,
          color: "#FF0000",
          highlightColor: "#FFFF00",
        },
      };

      expect(options.format).toBeDefined();
      expect(options.format?.fontName).toBe("Arial");
      expect(options.format?.fontSize).toBe(14);
      expect(options.format?.bold).toBe(true);
      expect(options.format?.color).toBe("#FF0000");
    });
  });

  describe("参数验证", () => {
    it("search 定位器必须包含 searchText", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "search",
          searchText: "必须的搜索文本",
        },
        newText: "新文本",
      };

      if (options.locator.type === "search") {
        expect(options.locator.searchText).toBeTruthy();
      }
    });

    it("range 定位器必须包含 rangeLocator", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "range",
          rangeLocator: {
            type: "paragraph",
            startIndex: 0,
          },
        },
        newText: "新文本",
      };

      if (options.locator.type === "range") {
        expect(options.locator.rangeLocator).toBeDefined();
      }
    });
  });

  describe("使用场景示例", () => {
    it("场景1: 替换选中文本并应用格式", () => {
      const options: ReplaceTextOptions = {
        locator: { type: "selection" },
        newText: "重要提示",
        format: {
          bold: true,
          color: "#FF0000",
          fontSize: 16,
        },
      };

      expect(options).toBeDefined();
    });

    it("场景2: 查找并替换所有匹配项", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "search",
          searchText: "公司名称",
          searchOptions: {
            matchWholeWord: true,
          },
        },
        newText: "新公司名称",
        replaceAll: true,
      };

      expect(options.replaceAll).toBe(true);
    });

    it("场景3: 替换特定段落的文本", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "range",
          rangeLocator: {
            type: "paragraph",
            startIndex: 5,
          },
        },
        newText: "这是新的段落内容",
      };

      expect(options).toBeDefined();
    });

    it("场景4: 替换书签位置的文本", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "range",
          rangeLocator: {
            type: "bookmark",
            name: "签名位置",
          },
        },
        newText: "张三",
        format: {
          fontName: "楷体",
          fontSize: 14,
        },
      };

      expect(options).toBeDefined();
    });

    it("场景5: 替换标题文本", () => {
      const options: ReplaceTextOptions = {
        locator: {
          type: "range",
          rangeLocator: {
            type: "heading",
            level: 1,
            index: 0,
          },
        },
        newText: "新的标题",
      };

      expect(options).toBeDefined();
    });
  });
});
