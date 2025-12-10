/**
 * 文件名: tableOfContents.ts
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: None
 * 描述: 目录管理工具核心逻辑 / Table of Contents Management Tool Core Logic
 */

/* global Word, console */

import type { InsertLocation } from "./types";

// 重新导出以保持向后兼容性 / Re-export for backward compatibility
export type { InsertLocation };

/**
 * 目录选项 / Table of Contents Options
 */
export interface TOCOptions {
  /** 目录标题，默认为"目录" / TOC title, default is "目录" */
  title?: string;

  /** 包含的标题级别（1-9），默认 [1, 2, 3] / Heading levels to include, default [1, 2, 3] */
  levels?: number[];

  /** 是否显示页码，默认为 true / Whether to show page numbers, default true */
  showPageNumbers?: boolean;

  /** 页码是否右对齐，默认为 true / Whether to right-align page numbers, default true */
  rightAlignPageNumbers?: boolean;

  /** 是否使用超链接，默认为 true / Whether to use hyperlinks, default true */
  useHyperlinks?: boolean;

  /** 是否包含隐藏文本，默认为 false / Whether to include hidden text, default false */
  includeHidden?: boolean;
}

/**
 * 目录信息 / Table of Contents Info
 */
export interface TOCInfo {
  /** 目录索引（从0开始）/ TOC index (0-based) */
  index: number;

  /** 目录范围文本 / TOC range text */
  text: string;

  /** 目录条目数量 / Number of TOC entries */
  entryCount: number;

  /** 包含的标题级别 / Included heading levels */
  levels: number[];
}

/**
 * 插入目录结果 / Insert TOC Result
 */
export interface InsertTOCResult {
  /** 是否成功 / Success */
  success: boolean;

  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;

  /** 目录信息 / TOC info */
  tocInfo?: TOCInfo;
}

/**
 * 更新目录结果 / Update TOC Result
 */
export interface UpdateTOCResult {
  /** 是否成功 / Success */
  success: boolean;

  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;

  /** 更新的目录数量 / Number of TOCs updated */
  updatedCount?: number;
}

/**
 * 删除目录结果 / Delete TOC Result
 */
export interface DeleteTOCResult {
  /** 是否成功 / Success */
  success: boolean;

  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;

  /** 删除的目录数量 / Number of TOCs deleted */
  deletedCount?: number;
}

/**
 * 获取目录列表结果 / Get TOC List Result
 */
export interface GetTOCListResult {
  /** 是否成功 / Success */
  success: boolean;

  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;

  /** 目录列表 / TOC list */
  tocs?: TOCInfo[];
}

/**
 * 在文档中插入目录 / Insert table of contents in document
 *
 * @remarks
 * - 目录会自动提取文档中的标题并生成链接
 * - 支持自定义标题级别、页码显示等选项
 * - 支持在文档开头、结尾、选中内容前后插入
 * - TOC automatically extracts headings and generates links
 * - Supports custom heading levels, page number display, etc.
 * - Supports insertion at document start, end, before/after selection
 *
 * @param location - 插入位置 / Insert location
 *   - "Start": 文档开头 / Document start
 *   - "End": 文档结尾 / Document end
 *   - "Before": 选中内容之前 / Before selection
 *   - "After": 选中内容之后 / After selection
 *   - "Replace": 替换选中内容 / Replace selection
 *
 * @param options - 目录选项 / TOC options
 *
 * @example
 * ```typescript
 * // 在文档开头插入默认目录 / Insert default TOC at document start
 * await insertTableOfContents("Start");
 *
 * // 在文档末尾插入自定义目录 / Insert custom TOC at document end
 * await insertTableOfContents("End", {
 *   title: "Table of Contents",
 *   levels: [1, 2, 3, 4],
 *   showPageNumbers: true,
 *   rightAlignPageNumbers: true,
 *   useHyperlinks: true
 * });
 * ```
 */
export async function insertTableOfContents(
  location: InsertLocation = "Start",
  options?: TOCOptions
): Promise<InsertTOCResult> {
  try {
    // 设置默认选项 / Set default options
    const defaultOptions: Required<TOCOptions> = {
      title: "目录",
      levels: [1, 2, 3],
      showPageNumbers: true,
      rightAlignPageNumbers: true,
      useHyperlinks: true,
      includeHidden: false,
    };

    const finalOptions = { ...defaultOptions, ...options };

    let tocInfo: TOCInfo | undefined;

    await Word.run(async (context) => {
      // 获取插入范围 / Get insert range
      let insertRange: Word.Range;

      switch (location) {
        case "Start":
          // 在文档开头插入 / Insert at document start
          insertRange = context.document.body.getRange("Start");
          break;
        case "End":
          // 在文档末尾插入 / Insert at document end
          insertRange = context.document.body.getRange("End");
          break;
        case "Before":
          // 在选中内容之前插入 / Insert before selection
          insertRange = context.document.getSelection();
          break;
        case "After":
          // 在选中内容之后插入 / Insert after selection
          insertRange = context.document.getSelection();
          break;
        case "Replace":
          // 替换选中内容 / Replace selection
          insertRange = context.document.getSelection();
          // 先删除选中内容 / Delete selection first
          insertRange.delete();
          insertRange = context.document.getSelection();
          break;
        default:
          insertRange = context.document.body.getRange("Start");
      }

      // 插入目录标题（如果提供）/ Insert TOC title (if provided)
      if (finalOptions.title) {
        const titleRange =
          location === "Before" || location === "Start"
            ? insertRange.insertText(finalOptions.title + "\n", "Before")
            : insertRange.insertText(finalOptions.title + "\n", "After");

        // 设置标题格式 / Set title format
        titleRange.font.size = 16;
        titleRange.font.bold = true;
        titleRange.paragraphs.getFirst().alignment = Word.Alignment.centered;

        // 更新插入范围为标题后的位置 / Update insert range to after title
        if (location === "Before" || location === "Start") {
          insertRange = titleRange.getRange("After");
        } else {
          insertRange = titleRange.getRange("After");
        }
      }

      // 计算标题级别范围 / Calculate heading level range
      const minLevel = Math.min(...finalOptions.levels);
      const maxLevel = Math.max(...finalOptions.levels);

      // 插入目录 / Insert TOC
      // Word.js API: insertField(insertLocation, fieldType, text, removeFormatting)
      // 使用 TOC 域代码来插入目录 / Use TOC field code to insert table of contents
      // TOC 域代码格式: TOC \o "minLevel-maxLevel" \h \z \u
      // \o: 指定标题级别范围 / Specify heading level range
      // \h: 使用超链接 / Use hyperlinks
      // \z: 隐藏页码和制表符 / Hide page numbers and tabs (如果不需要页码)
      // \u: 使用超链接而不是页码 / Use hyperlinks instead of page numbers
      let tocFieldCode = `TOC \\o "${minLevel}-${maxLevel}"`;
      if (finalOptions.useHyperlinks) {
        tocFieldCode += " \\h";
      }
      if (!finalOptions.showPageNumbers) {
        tocFieldCode += " \\n";
      }

      const tocField = insertRange.insertField(
        location === "Before" || location === "Start"
          ? Word.InsertLocation.before
          : Word.InsertLocation.after,
        Word.FieldType.toc,
        tocFieldCode,
        false
      );

      // 获取目录的结果范围 / Get TOC result range
      const tocRange = tocField.result;
      tocField.load("code");

      // 加载目录信息 / Load TOC info
      tocRange.load("text");

      await context.sync();

      // 获取所有目录 / Get all TOCs
      const tocs = context.document.contentControls.getByTypes([Word.ContentControlType.richText]);
      // eslint-disable-next-line office-addins/no-navigational-load
      tocs.load("items/length");

      await context.sync();

      // 查找刚插入的目录（最后一个）/ Find the just inserted TOC (last one)
      const tocIndex = tocs.items.length - 1;

      tocInfo = {
        index: tocIndex,
        text: tocRange.text,
        entryCount: tocRange.text.split("\n").filter((line) => line.trim().length > 0).length,
        levels: finalOptions.levels,
      };
    });

    return {
      success: true,
      tocInfo,
    };
  } catch (error) {
    console.error("插入目录失败 / Insert TOC failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 更新文档中的目录 / Update table of contents in document
 *
 * @remarks
 * - 更新目录以反映文档结构的最新变化
 * - 如果不指定索引，则更新所有目录
 * - Update TOC to reflect latest document structure changes
 * - If index is not specified, update all TOCs
 *
 * @param tocIndex - 目录索引（可选，不指定则更新所有）/ TOC index (optional, update all if not specified)
 *
 * @example
 * ```typescript
 * // 更新所有目录 / Update all TOCs
 * await updateTableOfContents();
 *
 * // 更新指定目录 / Update specific TOC
 * await updateTableOfContents(0);
 * ```
 */
export async function updateTableOfContents(tocIndex?: number): Promise<UpdateTOCResult> {
  try {
    let updatedCount = 0;

    await Word.run(async (context) => {
      // 获取文档中的所有域 / Get all fields in document
      const fields = context.document.body.fields;
      fields.load("items");

      await context.sync();

      if (tocIndex !== undefined) {
        // 更新指定目录 / Update specific TOC
        if (tocIndex >= 0 && tocIndex < fields.items.length) {
          const field = fields.items[tocIndex];
          // 更新域结果 / Update field result
          field.updateResult();
          updatedCount = 1;
        } else {
          throw new Error(
            `目录索引 ${tocIndex} 超出范围 (0-${fields.items.length - 1}) / TOC index ${tocIndex} out of range (0-${fields.items.length - 1})`
          );
        }
      } else {
        // 更新所有目录 / Update all TOCs
        // 遍历并更新所有域 / Iterate and update all fields
        fields.items.forEach((field) => {
          field.updateResult();
        });

        updatedCount = fields.items.length;
      }

      // 同步更新操作 / Sync update operations
      // 注意：updateResult() 后的 sync 可能会因为对象引用失效而抛出 ItemNotFound
      // 但这是非阻断性的，更新操作实际上已经成功
      // Note: sync after updateResult() may throw ItemNotFound due to invalid object references
      // But this is non-blocking, the update operation has actually succeeded
      try {
        await context.sync();
      } catch (syncError) {
        // 忽略 ItemNotFound 错误，因为更新已经成功
        // Ignore ItemNotFound error as the update has already succeeded
        const errorMessage = syncError instanceof Error ? syncError.message : String(syncError);
        if (!errorMessage.includes("ItemNotFound")) {
          // 其他错误需要抛出 / Other errors need to be thrown
          throw syncError;
        }
        // ItemNotFound 错误被忽略，更新已成功 / ItemNotFound error ignored, update succeeded
      }
    });

    return {
      success: true,
      updatedCount,
    };
  } catch (error) {
    console.error("更新目录失败 / Update TOC failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 删除文档中的目录 / Delete table of contents in document
 *
 * @remarks
 * - 删除指定索引的目录
 * - 如果不指定索引，则删除所有目录
 * - Delete TOC at specified index
 * - If index is not specified, delete all TOCs
 *
 * @param tocIndex - 目录索引（可选，不指定则删除所有）/ TOC index (optional, delete all if not specified)
 *
 * @example
 * ```typescript
 * // 删除所有目录 / Delete all TOCs
 * await deleteTableOfContents();
 *
 * // 删除指定目录 / Delete specific TOC
 * await deleteTableOfContents(0);
 * ```
 */
export async function deleteTableOfContents(tocIndex?: number): Promise<DeleteTOCResult> {
  try {
    let deletedCount = 0;

    await Word.run(async (context) => {
      // 获取文档中的所有域 / Get all fields in document
      const fields = context.document.body.fields;
      fields.load("items");

      await context.sync();

      if (tocIndex !== undefined) {
        // 删除指定目录 / Delete specific TOC
        if (tocIndex >= 0 && tocIndex < fields.items.length) {
          const field = fields.items[tocIndex];
          field.delete();
          deletedCount = 1;
        }
      } else {
        // 删除所有目录 / Delete all TOCs
        // 从后往前删除以避免索引问题 / Delete from back to front to avoid index issues
        for (let i = fields.items.length - 1; i >= 0; i--) {
          const field = fields.items[i];
          field.delete();
          deletedCount++;
        }
      }

      await context.sync();
    });

    return {
      success: true,
      deletedCount,
    };
  } catch (error) {
    console.error("删除目录失败 / Delete TOC failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 获取文档中的目录列表 / Get list of table of contents in document
 *
 * @remarks
 * - 获取文档中所有目录的信息
 * - Get information of all TOCs in document
 *
 * @example
 * ```typescript
 * // 获取所有目录 / Get all TOCs
 * const result = await getTableOfContentsList();
 * if (result.success) {
 *   console.log(`找到 ${result.tocs?.length} 个目录`);
 * }
 * ```
 */
export async function getTableOfContentsList(): Promise<GetTOCListResult> {
  try {
    const tocs: TOCInfo[] = [];

    await Word.run(async (context) => {
      // 获取文档中的所有域 / Get all fields in document
      const fields = context.document.body.fields;
      fields.load("items");

      await context.sync();

      // 遍历所有域并提取目录信息 / Iterate all fields and extract TOC info
      for (let i = 0; i < fields.items.length; i++) {
        const field = fields.items[i];
        field.load("result");
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();

        const range = field.result;
        range.load("text");

        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();

        // 解析目录信息 / Parse TOC info
        const text = range.text;
        const entryCount = text.split("\n").filter((line) => line.trim().length > 0).length;

        tocs.push({
          index: i,
          text,
          entryCount,
          levels: [1, 2, 3], // 默认级别，实际应该从域代码中解析 / Default levels, should parse from field code
        });
      }
    });

    return {
      success: true,
      tocs,
    };
  } catch (error) {
    console.error("获取目录列表失败 / Get TOC list failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
