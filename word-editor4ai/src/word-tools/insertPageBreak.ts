/**
 * 文件名: insertPageBreak.ts
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: None
 * 描述: 插入分页符工具核心逻辑 / Insert Page Break Tool Core Logic
 */

/* global Word, console */

import type { InsertLocation } from "./types";

// 重新导出以保持向后兼容性 / Re-export for backward compatibility
export type { InsertLocation };

/**
 * 插入分页符结果 / Insert Page Break Result
 */
export interface InsertPageBreakResult {
  /** 是否成功 / Success */
  success: boolean;
  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;
}

/**
 * 在文档中插入分页符 / Insert page break in document
 *
 * @remarks
 * - 分页符会强制在指定位置开始新页面
 * - 支持在文档开头、结尾、选中内容前后插入
 * - Page break forces a new page at the specified location
 * - Supports insertion at document start, end, before/after selection
 *
 * @param location - 插入位置 / Insert location
 *   - "Start": 文档开头 / Document start
 *   - "End": 文档结尾 / Document end
 *   - "Before": 选中内容之前 / Before selection
 *   - "After": 选中内容之后 / After selection
 *   - "Replace": 替换选中内容 / Replace selection
 *
 * @example
 * ```typescript
 * // 在文档末尾插入分页符 / Insert page break at document end
 * await insertPageBreak("End");
 *
 * // 在选中内容之后插入分页符 / Insert page break after selection
 * await insertPageBreak("After");
 * ```
 */
export async function insertPageBreak(location: InsertLocation): Promise<InsertPageBreakResult> {
  try {
    await Word.run(async (context) => {
      // 获取插入范围 / Get insert range
      let insertRange: Word.Range;
      let breakLocation: "Before" | "After";

      switch (location) {
        case "Start":
          // 在文档开头插入分页符 / Insert page break at document start
          insertRange = context.document.body.getRange("Start");
          breakLocation = "Before";
          break;
        case "End":
          // 在文档末尾插入分页符 / Insert page break at document end
          insertRange = context.document.body.getRange("End");
          breakLocation = "After";
          break;
        case "Before":
          // 在选中内容之前插入分页符 / Insert page break before selection
          insertRange = context.document.getSelection();
          breakLocation = "Before";
          break;
        case "After":
          // 在选中内容之后插入分页符 / Insert page break after selection
          insertRange = context.document.getSelection();
          breakLocation = "After";
          break;
        case "Replace":
          // 替换选中内容为分页符 / Replace selection with page break
          insertRange = context.document.getSelection();
          breakLocation = "Before";
          // 先删除选中内容 / Delete selection first
          insertRange.delete();
          break;
        default:
          insertRange = context.document.body.getRange("End");
          breakLocation = "After";
      }

      // 插入分页符 / Insert page break
      // Word.BreakType.page 表示分页符
      // Word.BreakType.page represents page break
      insertRange.insertBreak(Word.BreakType.page, breakLocation);

      await context.sync();
    });

    return {
      success: true,
    };
  } catch (error) {
    console.error("插入分页符失败 / Insert page break failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
