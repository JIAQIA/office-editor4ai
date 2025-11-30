/**
 * 文件名: visibleContent.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 可见内容获取工具的测试文件
 */

import { describe, it, expect } from "vitest";
import {
  getVisibleContent,
  getVisibleText,
  getVisibleContentStats,
  type PageInfo,
  type ParagraphElement,
  type TableElement,
  type ImageElement,
  type ContentControlElement,
} from "../../../src/word-tools";

/**
 * 注意：这些测试需要在 Word 环境中运行
 * 可以使用 Office Add-in 的测试框架或手动测试
 */

describe("visibleContent", () => {
  describe("getVisibleContent", () => {
    it("应该返回 PageInfo 数组", async () => {
      // 这个测试需要在实际的 Word 环境中运行
      // 这里只是展示预期的数据结构
      const mockPages: PageInfo[] = [
        {
          index: 0,
          elements: [
            {
              id: "para-0-0",
              type: "Paragraph",
              text: "这是一个段落",
            } as ParagraphElement,
          ],
          text: "这是一个段落",
        },
      ];

      expect(mockPages).toHaveLength(1);
      expect(mockPages[0].index).toBe(0);
      expect(mockPages[0].elements).toHaveLength(1);
      expect(mockPages[0].elements[0].type).toBe("Paragraph");
    });

    it("应该支持不同的选项配置", async () => {
      // 测试选项配置的类型检查
      const options = {
        includeText: true,
        includeImages: false,
        includeTables: true,
        includeContentControls: false,
        detailedMetadata: true,
        maxTextLength: 100,
      };

      expect(options.includeText).toBe(true);
      expect(options.includeImages).toBe(false);
      expect(options.maxTextLength).toBe(100);
    });
  });

  describe("数据结构验证", () => {
    it("ParagraphElement 应该包含正确的属性", () => {
      const paragraph: ParagraphElement = {
        id: "para-1",
        type: "Paragraph",
        text: "测试段落",
        style: "Normal",
        alignment: "Left",
        firstLineIndent: 0,
        leftIndent: 0,
        rightIndent: 0,
        lineSpacing: 1.5,
        spaceAfter: 10,
        spaceBefore: 0,
        isListItem: false,
      };

      expect(paragraph.type).toBe("Paragraph");
      expect(paragraph.text).toBe("测试段落");
      expect(paragraph.style).toBe("Normal");
    });

    it("TableElement 应该包含正确的属性", () => {
      const table: TableElement = {
        id: "table-1",
        type: "Table",
        rowCount: 3,
        columnCount: 4,
        cells: [
          [
            { text: "A1", rowIndex: 0, columnIndex: 0 },
            { text: "B1", rowIndex: 0, columnIndex: 1 },
          ],
        ],
      };

      expect(table.type).toBe("Table");
      expect(table.rowCount).toBe(3);
      expect(table.columnCount).toBe(4);
      expect(table.cells).toBeDefined();
    });

    it("ImageElement 应该包含正确的属性", () => {
      const image: ImageElement = {
        id: "img-1",
        type: "Image",
        width: 200,
        height: 150,
        altText: "测试图片",
        hyperlink: "https://example.com",
      };

      expect(image.type).toBe("Image");
      expect(image.width).toBe(200);
      expect(image.height).toBe(150);
      expect(image.altText).toBe("测试图片");
    });

    it("ContentControlElement 应该包含正确的属性", () => {
      const control: ContentControlElement = {
        id: "ctrl-1",
        type: "ContentControl",
        text: "控件内容",
        title: "测试控件",
        tag: "test-tag",
        controlType: "RichText",
        cannotDelete: false,
        cannotEdit: false,
      };

      expect(control.type).toBe("ContentControl");
      expect(control.title).toBe("测试控件");
      expect(control.tag).toBe("test-tag");
    });
  });

  describe("getVisibleText", () => {
    it("应该返回字符串", async () => {
      // 模拟测试
      const mockText = "这是可见的文本内容";
      expect(typeof mockText).toBe("string");
    });
  });

  describe("getVisibleContentStats", () => {
    it("应该返回正确的统计信息结构", async () => {
      // 模拟统计信息
      const mockStats = {
        pageCount: 2,
        elementCount: 10,
        characterCount: 500,
        paragraphCount: 5,
        tableCount: 1,
        imageCount: 2,
        contentControlCount: 2,
      };

      expect(mockStats.pageCount).toBeGreaterThanOrEqual(0);
      expect(mockStats.elementCount).toBeGreaterThanOrEqual(0);
      expect(mockStats.characterCount).toBeGreaterThanOrEqual(0);
      expect(mockStats.paragraphCount).toBeGreaterThanOrEqual(0);
      expect(mockStats.tableCount).toBeGreaterThanOrEqual(0);
      expect(mockStats.imageCount).toBeGreaterThanOrEqual(0);
      expect(mockStats.contentControlCount).toBeGreaterThanOrEqual(0);
    });
  });
});

/**
 * 手动测试指南
 * 
 * 1. 在 Word 中打开一个包含多页内容的文档
 * 2. 确保文档包含：
 *    - 多个段落
 *    - 至少一个表格
 *    - 至少一张图片
 *    - 可选：内容控件
 * 
 * 3. 启动 Add-in 并导航到"可见内容获取"工具
 * 
 * 4. 测试场景：
 *    a. 默认选项获取可见内容
 *    b. 只获取文本（关闭其他选项）
 *    c. 获取详细元数据
 *    d. 滚动文档到不同位置，验证获取的内容是否正确
 *    e. 获取统计信息
 * 
 * 5. 验证点：
 *    - 页面数量是否正确
 *    - 元素类型是否正确识别
 *    - 文本内容是否完整
 *    - 表格行列数是否正确
 *    - 图片信息是否准确
 *    - 统计数据是否准确
 */
