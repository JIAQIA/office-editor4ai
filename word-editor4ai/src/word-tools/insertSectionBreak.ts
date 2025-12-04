/**
 * 文件名: insertSectionBreak.ts
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: None
 * 描述: 插入分节符工具核心逻辑 / Insert Section Break Tool Core Logic
 */

/* global Word, console */

import type { InsertLocation } from "./types";

// 重新导出以保持向后兼容性 / Re-export for backward compatibility
export type { InsertLocation };

/**
 * 分节符类型 / Section Break Type
 *
 * @remarks
 * - Continuous: 连续分节符，新节从当前页继续
 * - NextPage: 下一页分节符，新节从下一页开始
 * - OddPage: 奇数页分节符，新节从下一个奇数页开始
 * - EvenPage: 偶数页分节符，新节从下一个偶数页开始
 *
 * - Continuous: Continuous section break, new section continues on current page
 * - NextPage: Next page section break, new section starts on next page
 * - OddPage: Odd page section break, new section starts on next odd page
 * - EvenPage: Even page section break, new section starts on next even page
 */
export type SectionBreakType = "Continuous" | "NextPage" | "OddPage" | "EvenPage";

/**
 * 插入分节符结果 / Insert Section Break Result
 */
export interface InsertSectionBreakResult {
  /** 是否成功 / Success */
  success: boolean;
  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;
  /** 新创建的节索引 / New section index */
  sectionIndex?: number;
}

/**
 * 将自定义分节符类型映射到 Word.BreakType / Map custom section break type to Word.BreakType
 */
function mapSectionBreakType(breakType: SectionBreakType): Word.BreakType {
  const typeMap: Record<SectionBreakType, Word.BreakType> = {
    Continuous: Word.BreakType.sectionContinuous,
    NextPage: Word.BreakType.sectionNext,
    OddPage: Word.BreakType.sectionOdd,
    EvenPage: Word.BreakType.sectionEven,
  };
  return typeMap[breakType];
}

/**
 * 在文档中插入分节符 / Insert section break in document
 *
 * @remarks
 * - 分节符用于将文档分成不同的节，每节可以有独立的页面设置
 * - 支持四种分节符类型：连续、下一页、奇数页、偶数页
 * - 支持在文档开头、结尾、选中内容前后插入
 * - Section breaks divide document into sections with independent page settings
 * - Supports four types: continuous, next page, odd page, even page
 * - Supports insertion at document start, end, before/after selection
 *
 * @param breakType - 分节符类型 / Section break type
 *   - "Continuous": 连续分节符 / Continuous section break
 *   - "NextPage": 下一页分节符 / Next page section break
 *   - "OddPage": 奇数页分节符 / Odd page section break
 *   - "EvenPage": 偶数页分节符 / Even page section break
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
 * // 在文档末尾插入下一页分节符 / Insert next page section break at document end
 * await insertSectionBreak("NextPage", "End");
 *
 * // 在选中内容之后插入连续分节符 / Insert continuous section break after selection
 * await insertSectionBreak("Continuous", "After");
 *
 * // 在文档开头插入奇数页分节符 / Insert odd page section break at document start
 * await insertSectionBreak("OddPage", "Start");
 * ```
 */
export async function insertSectionBreak(
  breakType: SectionBreakType,
  location: InsertLocation
): Promise<InsertSectionBreakResult> {
  try {
    let newSectionIndex: number | undefined;

    await Word.run(async (context) => {
      // 获取插入范围 / Get insert range
      let insertRange: Word.Range;
      let breakLocation: "Before" | "After";

      switch (location) {
        case "Start":
          // 在文档开头插入分节符 / Insert section break at document start
          insertRange = context.document.body.getRange("Start");
          breakLocation = "Before";
          break;
        case "End":
          // 在文档末尾插入分节符 / Insert section break at document end
          insertRange = context.document.body.getRange("End");
          breakLocation = "After";
          break;
        case "Before":
          // 在选中内容之前插入分节符 / Insert section break before selection
          insertRange = context.document.getSelection();
          breakLocation = "Before";
          break;
        case "After":
          // 在选中内容之后插入分节符 / Insert section break after selection
          insertRange = context.document.getSelection();
          breakLocation = "After";
          break;
        case "Replace":
          // 替换选中内容为分节符 / Replace selection with section break
          insertRange = context.document.getSelection();
          breakLocation = "Before";
          // 先删除选中内容 / Delete selection first
          insertRange.delete();
          break;
        default:
          insertRange = context.document.body.getRange("End");
          breakLocation = "After";
      }

      // 插入分节符 / Insert section break
      // 使用 insertBreak 方法插入分节符
      // Use insertBreak method to insert section break
      const wordBreakType = mapSectionBreakType(breakType);
      insertRange.insertBreak(wordBreakType, breakLocation);

      // 获取新创建的节索引 / Get new section index
      // 注意：分节符会创建新的节，我们需要获取节的数量
      // Note: Section break creates a new section, we need to get section count
      const sections = context.document.sections;
      sections.load("items");

      await context.sync();

      // 新节的索引是总节数减1（因为索引从0开始）
      // New section index is total sections minus 1 (because index starts from 0)
      newSectionIndex = sections.items.length - 1;
    });

    return {
      success: true,
      sectionIndex: newSectionIndex,
    };
  } catch (error) {
    console.error("插入分节符失败 / Insert section break failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
