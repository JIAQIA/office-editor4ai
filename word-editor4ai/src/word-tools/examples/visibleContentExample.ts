/**
 * 文件名: visibleContentExample.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 可见内容获取工具的使用示例
 */

import {
  getVisibleContent,
  getVisibleText,
  getVisibleContentStats,
} from "../visibleContent";
import type { PageInfo, AnyContentElement } from "../types";

/**
 * 示例 1: 获取可见内容的基本用法
 */
export async function example1_BasicUsage() {
  console.log("=== 示例 1: 基本用法 ===");

  try {
    const pages = await getVisibleContent({
      includeText: true,
      includeImages: true,
      includeTables: true,
      includeContentControls: true,
    });

    console.log(`找到 ${pages.length} 个可见页面`);

    for (const page of pages) {
      console.log(`\n页面 ${page.index + 1}:`);
      console.log(`- 元素数量: ${page.elements.length}`);
      console.log(`- 文本长度: ${page.text?.length || 0} 字符`);
    }
  } catch (error) {
    console.error("获取可见内容失败:", error);
  }
}

/**
 * 示例 2: 只获取文本内容
 */
export async function example2_TextOnly() {
  console.log("=== 示例 2: 只获取文本 ===");

  try {
    const text = await getVisibleText();
    console.log("可见文本内容:");
    console.log(text);
    console.log(`\n总字符数: ${text.length}`);
  } catch (error) {
    console.error("获取可见文本失败:", error);
  }
}

/**
 * 示例 3: 获取统计信息
 */
export async function example3_Statistics() {
  console.log("=== 示例 3: 统计信息 ===");

  try {
    const stats = await getVisibleContentStats();

    console.log("可见内容统计:");
    console.log(`- 页面数: ${stats.pageCount}`);
    console.log(`- 元素总数: ${stats.elementCount}`);
    console.log(`- 字符数: ${stats.characterCount}`);
    console.log(`- 段落数: ${stats.paragraphCount}`);
    console.log(`- 表格数: ${stats.tableCount}`);
    console.log(`- 图片数: ${stats.imageCount}`);
    console.log(`- 内容控件数: ${stats.contentControlCount}`);
  } catch (error) {
    console.error("获取统计信息失败:", error);
  }
}

/**
 * 示例 4: 获取详细元数据
 */
export async function example4_DetailedMetadata() {
  console.log("=== 示例 4: 详细元数据 ===");

  try {
    const pages = await getVisibleContent({
      includeText: true,
      includeImages: true,
      includeTables: true,
      includeContentControls: true,
      detailedMetadata: true, // 启用详细元数据
    });

    for (const page of pages) {
      console.log(`\n页面 ${page.index + 1}:`);

      for (const element of page.elements) {
        console.log(`\n元素类型: ${element.type}`);

        if (element.type === "Paragraph") {
          const para = element as any;
          console.log(`- 样式: ${para.style || "无"}`);
          console.log(`- 对齐: ${para.alignment || "无"}`);
          console.log(`- 是否列表项: ${para.isListItem ? "是" : "否"}`);
        } else if (element.type === "Table") {
          const table = element as any;
          console.log(`- 行数: ${table.rowCount}`);
          console.log(`- 列数: ${table.columnCount}`);
        } else if (element.type === "Image" || element.type === "InlinePicture") {
          const img = element as any;
          console.log(`- 尺寸: ${img.width}×${img.height}`);
          console.log(`- 替代文本: ${img.altText || "无"}`);
        } else if (element.type === "ContentControl") {
          const ctrl = element as any;
          console.log(`- 标题: ${ctrl.title || "无"}`);
          console.log(`- 标签: ${ctrl.tag || "无"}`);
          console.log(`- 类型: ${ctrl.controlType || "无"}`);
        }
      }
    }
  } catch (error) {
    console.error("获取详细元数据失败:", error);
  }
}

/**
 * 示例 5: 限制文本长度
 */
export async function example5_LimitTextLength() {
  console.log("=== 示例 5: 限制文本长度 ===");

  try {
    const pages = await getVisibleContent({
      includeText: true,
      maxTextLength: 100, // 限制每个文本元素最多 100 字符
    });

    for (const page of pages) {
      console.log(`\n页面 ${page.index + 1}:`);

      for (const element of page.elements) {
        if (element.text) {
          console.log(`- ${element.type}: ${element.text}`);
        }
      }
    }
  } catch (error) {
    console.error("获取内容失败:", error);
  }
}

/**
 * 示例 6: 按元素类型分组
 */
export async function example6_GroupByType() {
  console.log("=== 示例 6: 按元素类型分组 ===");

  try {
    const pages = await getVisibleContent({
      includeText: true,
      includeImages: true,
      includeTables: true,
      includeContentControls: true,
    });

    const grouped: Record<string, AnyContentElement[]> = {};

    for (const page of pages) {
      for (const element of page.elements) {
        if (!grouped[element.type]) {
          grouped[element.type] = [];
        }
        grouped[element.type].push(element);
      }
    }

    console.log("\n按类型分组的元素:");
    for (const [type, elements] of Object.entries(grouped)) {
      console.log(`\n${type}: ${elements.length} 个`);

      // 显示前 3 个元素的摘要
      elements.slice(0, 3).forEach((element, index) => {
        const preview = element.text
          ? element.text.substring(0, 50) + (element.text.length > 50 ? "..." : "")
          : "(无文本)";
        console.log(`  ${index + 1}. ${preview}`);
      });

      if (elements.length > 3) {
        console.log(`  ... 还有 ${elements.length - 3} 个`);
      }
    }
  } catch (error) {
    console.error("分组失败:", error);
  }
}

/**
 * 示例 7: 提取表格数据
 */
export async function example7_ExtractTableData() {
  console.log("=== 示例 7: 提取表格数据 ===");

  try {
    const pages = await getVisibleContent({
      includeText: true,
      includeTables: true,
      detailedMetadata: true,
    });

    for (const page of pages) {
      const tables = page.elements.filter((e) => e.type === "Table");

      if (tables.length > 0) {
        console.log(`\n页面 ${page.index + 1} 包含 ${tables.length} 个表格:`);

        tables.forEach((table: any, index) => {
          console.log(`\n表格 ${index + 1}:`);
          console.log(`- 尺寸: ${table.rowCount}×${table.columnCount}`);

          if (table.cells && table.cells.length > 0) {
            console.log("- 内容预览:");
            // 显示前 2 行
            table.cells.slice(0, 2).forEach((row: any[], rowIndex: number) => {
              const rowText = row.map((cell) => cell.text).join(" | ");
              console.log(`  行 ${rowIndex + 1}: ${rowText}`);
            });
          }
        });
      }
    }
  } catch (error) {
    console.error("提取表格数据失败:", error);
  }
}

/**
 * 示例 8: 查找特定内容
 */
export async function example8_SearchContent() {
  console.log("=== 示例 8: 查找特定内容 ===");

  const searchTerm = "重要"; // 要搜索的关键词

  try {
    const pages = await getVisibleContent({
      includeText: true,
    });

    console.log(`搜索关键词: "${searchTerm}"`);

    let foundCount = 0;

    for (const page of pages) {
      for (const element of page.elements) {
        if (element.text && element.text.includes(searchTerm)) {
          foundCount++;
          console.log(`\n找到匹配 (页面 ${page.index + 1}, ${element.type}):`);
          console.log(`  ${element.text.substring(0, 100)}...`);
        }
      }
    }

    console.log(`\n共找到 ${foundCount} 个匹配项`);
  } catch (error) {
    console.error("搜索失败:", error);
  }
}

/**
 * 运行所有示例
 */
export async function runAllExamples() {
  console.log("开始运行所有示例...\n");

  await example1_BasicUsage();
  await example2_TextOnly();
  await example3_Statistics();
  await example4_DetailedMetadata();
  await example5_LimitTextLength();
  await example6_GroupByType();
  await example7_ExtractTableData();
  await example8_SearchContent();

  console.log("\n所有示例运行完成!");
}
