/**
 * 文件名: replaceText.ts
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 统一的文本替换工具，支持选区、搜索、范围三种定位方式
 */

/* global Word, console */

import type { TextFormat, RangeLocator } from "./types";

/**
 * 查找选项 / Search Options
 */
export interface SearchOptions {
  /** 是否区分大小写 / Match case */
  matchCase?: boolean;
  /** 是否全字匹配 / Match whole word */
  matchWholeWord?: boolean;
  /** 是否使用通配符 / Use wildcards */
  matchWildcards?: boolean;
}

/**
 * 替换文本定位器 / Replace Text Locator
 */
export type ReplaceTextLocator =
  | {
      /** 定位类型：当前选区 / Locator type: current selection */
      type: "selection";
    }
  | {
      /** 定位类型：搜索匹配 / Locator type: search match */
      type: "search";
      /** 搜索文本 / Search text */
      searchText: string;
      /** 搜索选项 / Search options */
      searchOptions?: SearchOptions;
    }
  | {
      /** 定位类型：指定范围 / Locator type: specific range */
      type: "range";
      /** 范围定位器 / Range locator */
      rangeLocator: RangeLocator;
    };

/**
 * 替换文本选项 / Replace Text Options
 */
export interface ReplaceTextOptions {
  /** 定位方式 / Locator */
  locator: ReplaceTextLocator;
  /** 新文本内容 / New text content */
  newText: string;
  /** 文本格式（可选）/ Text format (optional) */
  format?: TextFormat;
  /** 是否替换所有匹配项（仅 search 模式）/ Replace all matches (search mode only) */
  replaceAll?: boolean;
}

/**
 * 替换结果 / Replace Result
 */
export interface ReplaceResult {
  /** 替换的数量 / Number of replacements */
  count: number;
  /** 是否成功 / Success */
  success: boolean;
  /** 错误信息（如果有）/ Error message (if any) */
  error?: string;
}

/**
 * 应用文本格式到范围 / Apply text format to range
 */
function applyTextFormat(range: Word.Range, format: TextFormat): void {
  const font = range.font;

  if (format.fontName !== undefined) {
    font.name = format.fontName;
  }
  if (format.fontSize !== undefined) {
    font.size = format.fontSize;
  }
  if (format.bold !== undefined) {
    font.bold = format.bold;
  }
  if (format.italic !== undefined) {
    font.italic = format.italic;
  }
  if (format.underline !== undefined) {
    font.underline = format.underline as Word.UnderlineType;
  }
  if (format.color !== undefined) {
    font.color = format.color;
  }
  if (format.highlightColor !== undefined) {
    font.highlightColor = format.highlightColor;
  }
  if (format.strikeThrough !== undefined) {
    font.strikeThrough = format.strikeThrough;
  }
  if (format.superscript !== undefined) {
    font.superscript = format.superscript;
  }
  if (format.subscript !== undefined) {
    font.subscript = format.subscript;
  }
}

/**
 * 根据范围定位器获取范围 / Get range by range locator
 */
async function getRangeByLocator(
  context: Word.RequestContext,
  locator: RangeLocator
): Promise<Word.Range> {
  switch (locator.type) {
    case "bookmark": {
      const bookmark = context.document.getBookmarkRangeOrNullObject(locator.name);
      bookmark.load("text");
      await context.sync();
      if (bookmark.isNullObject) {
        throw new Error(
          `书签 "${locator.name}" 不存在 / Bookmark "${locator.name}" does not exist`
        );
      }
      return bookmark;
    }

    case "heading": {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const matchedParagraphs: Word.Paragraph[] = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load(["style", "text"]);
      }
      await context.sync();

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const style = para.style.toLowerCase();

        if (style.includes("heading") || style.includes("标题")) {
          if (locator.level) {
            const levelMatch = style.match(/\d+/);
            if (levelMatch && parseInt(levelMatch[0]) === locator.level) {
              if (!locator.text || para.text.includes(locator.text)) {
                matchedParagraphs.push(para);
              }
            }
          } else if (!locator.text || para.text.includes(locator.text)) {
            matchedParagraphs.push(para);
          }
        }
      }

      if (matchedParagraphs.length === 0) {
        throw new Error("未找到匹配的标题 / No matching heading found");
      }

      const index = locator.index || 0;
      if (index >= matchedParagraphs.length) {
        throw new Error(`标题索引 ${index} 超出范围 / Heading index ${index} out of range`);
      }

      return matchedParagraphs[index].getRange();
    }

    case "paragraph": {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      if (locator.startIndex >= paragraphs.items.length) {
        throw new Error(
          `段落索引 ${locator.startIndex} 超出范围 / Paragraph index ${locator.startIndex} out of range`
        );
      }

      const startPara = paragraphs.items[locator.startIndex];
      if (locator.endIndex !== undefined) {
        if (locator.endIndex >= paragraphs.items.length) {
          throw new Error(
            `段落索引 ${locator.endIndex} 超出范围 / Paragraph index ${locator.endIndex} out of range`
          );
        }
        const endPara = paragraphs.items[locator.endIndex];
        const startRange = startPara.getRange("Start");
        const endRange = endPara.getRange("End");
        return startRange.expandTo(endRange);
      } else {
        return startPara.getRange();
      }
    }

    case "section": {
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      if (locator.index >= sections.items.length) {
        throw new Error(
          `节索引 ${locator.index} 超出范围 / Section index ${locator.index} out of range`
        );
      }

      const section = sections.items[locator.index];
      return section.body.getRange();
    }

    case "contentControl": {
      const contentControls = context.document.contentControls;
      contentControls.load("items");
      await context.sync();

      const matchedControls: Word.ContentControl[] = [];

      for (let i = 0; i < contentControls.items.length; i++) {
        const control = contentControls.items[i];
        control.load(["title", "tag"]);
      }
      await context.sync();

      for (let i = 0; i < contentControls.items.length; i++) {
        const control = contentControls.items[i];
        if (locator.title && control.title === locator.title) {
          matchedControls.push(control);
        } else if (locator.tag && control.tag === locator.tag) {
          matchedControls.push(control);
        }
      }

      if (matchedControls.length === 0) {
        throw new Error("未找到匹配的内容控件 / No matching content control found");
      }

      const index = locator.index || 0;
      if (index >= matchedControls.length) {
        throw new Error(
          `内容控件索引 ${index} 超出范围 / Content control index ${index} out of range`
        );
      }

      return matchedControls[index].getRange();
    }

    default:
      throw new Error(`不支持的定位器类型 / Unsupported locator type`);
  }
}

/**
 * 替换文本 / Replace Text
 *
 * @param options - 替换选项 / Replace options
 * @returns Promise<ReplaceResult> 替换结果 / Replace result
 *
 * @remarks
 * 此函数提供统一的文本替换能力，支持三种定位方式：
 * 1. selection - 替换当前选中的文本
 * 2. search - 查找并替换匹配的文本
 * 3. range - 替换指定范围的文本
 *
 * This function provides unified text replacement capability with three locator types:
 * 1. selection - Replace currently selected text
 * 2. search - Find and replace matching text
 * 3. range - Replace text in specific range
 *
 * @example
 * ```typescript
 * // 替换选中文本
 * await replaceText({
 *   locator: { type: "selection" },
 *   newText: "新文本",
 *   format: { bold: true }
 * });
 *
 * // 查找并替换
 * await replaceText({
 *   locator: {
 *     type: "search",
 *     searchText: "旧文本",
 *     searchOptions: { matchCase: true }
 *   },
 *   newText: "新文本",
 *   replaceAll: true
 * });
 *
 * // 替换指定范围
 * await replaceText({
 *   locator: {
 *     type: "range",
 *     rangeLocator: { type: "paragraph", startIndex: 0 }
 *   },
 *   newText: "新文本"
 * });
 * ```
 */
export async function replaceText(options: ReplaceTextOptions): Promise<ReplaceResult> {
  const { locator, newText, format, replaceAll = false } = options;

  try {
    return await Word.run(async (context) => {
      let count = 0;

      switch (locator.type) {
        case "selection": {
          const selection = context.document.getSelection();
          // 抑制 isEmpty 属性的导航 load 性能告警 / Suppress performance warning for navigation load on isEmpty
          // eslint-disable-next-line office-addins/no-navigational-load
          selection.load(["text", "isEmpty"]);
          await context.sync();

          if (selection.isEmpty) {
            return {
              count: 0,
              success: false,
              error: "没有选中的内容 / No selection",
            };
          }

          const newRange = selection.insertText(newText, "Replace");
          if (format) {
            applyTextFormat(newRange, format);
          }
          await context.sync();

          count = 1;
          break;
        }

        case "search": {
          const searchResults = context.document.body.search(locator.searchText, {
            matchCase: locator.searchOptions?.matchCase ?? false,
            matchWholeWord: locator.searchOptions?.matchWholeWord ?? false,
            matchWildcards: locator.searchOptions?.matchWildcards ?? false,
          });
          searchResults.load("items");
          await context.sync();

          if (searchResults.items.length === 0) {
            return {
              count: 0,
              success: false,
              error: `未找到匹配的文本 "${locator.searchText}" / No matching text found "${locator.searchText}"`,
            };
          }

          const itemsToReplace = replaceAll ? searchResults.items : [searchResults.items[0]];

          for (const result of itemsToReplace) {
            const newRange = result.insertText(newText, "Replace");
            if (format) {
              applyTextFormat(newRange, format);
            }
            count++;
          }

          await context.sync();
          break;
        }

        case "range": {
          const range = await getRangeByLocator(context, locator.rangeLocator);
          const newRange = range.insertText(newText, "Replace");
          if (format) {
            applyTextFormat(newRange, format);
          }
          await context.sync();

          count = 1;
          break;
        }

        default:
          return {
            count: 0,
            success: false,
            error: `不支持的定位器类型 / Unsupported locator type`,
          };
      }

      console.log(`成功替换 ${count} 处文本 / Successfully replaced ${count} text(s)`);
      return {
        count,
        success: true,
      };
    });
  } catch (error) {
    console.error("替换文本失败 / Failed to replace text:", error);
    return {
      count: 0,
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
