/**
 * 文件名: documentStats.ts
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取文档统计信息的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

/**
 * 文档统计信息 / Document Statistics
 */
export interface DocumentStats {
  /** 字符数（包含空格）/ Character count (with spaces) */
  characterCount: number;
  /** 字符数（不含空格）/ Character count (without spaces) */
  characterCountNoSpaces: number;
  /** 单词数 / Word count */
  wordCount: number;
  /** 段落数 / Paragraph count */
  paragraphCount: number;
  /** 页数 / Page count */
  pageCount: number;
  /** 节数 / Section count */
  sectionCount: number;
  /** 表格数 / Table count */
  tableCount: number;
  /** 图片数 / Image count */
  imageCount: number;
  /** 内嵌图片数 / Inline picture count */
  inlinePictureCount: number;
  /** 内容控件数 / Content control count */
  contentControlCount: number;
  /** 列表数 / List count */
  listCount: number;
  /** 脚注数 / Footnote count */
  footnoteCount: number;
  /** 尾注数 / Endnote count */
  endnoteCount: number;
  /** 标题数（按级别统计）/ Heading count by level */
  headingCounts: Record<number, number>;
  /** 总标题数 / Total heading count */
  totalHeadingCount: number;
}

/**
 * 获取文档统计信息选项 / Get Document Stats Options
 */
export interface GetDocumentStatsOptions {
  /** 是否包含页眉页脚内容 / Include header footer content */
  includeHeaderFooter?: boolean;
  /** 是否包含脚注尾注 / Include footnotes and endnotes */
  includeNotes?: boolean;
  /** 是否包含详细的标题统计 / Include detailed heading statistics */
  includeHeadingStats?: boolean;
}

/**
 * 判断段落是否为标题样式
 * Determine if a paragraph is a heading style
 */
function isHeadingStyle(style: string): boolean {
  const headingPattern = /^(Heading|标题)\s*(\d)$/i;
  return headingPattern.test(style);
}

/**
 * 从样式名称中提取标题级别
 * Extract heading level from style name
 */
function extractHeadingLevel(style: string): number {
  const headingPattern = /^(Heading|标题)\s*(\d)$/i;
  const match = style.match(headingPattern);
  if (match && match[2]) {
    return parseInt(match[2], 10);
  }
  return 0;
}

/**
 * 统计文本中的单词数（支持中英文）
 * Count words in text (supports Chinese and English)
 */
function countWords(text: string): number {
  if (!text || text.trim().length === 0) {
    return 0;
  }

  // 移除多余空格 / Remove extra spaces
  const cleanText = text.trim().replace(/\s+/g, " ");

  // 分离中文字符和英文单词 / Separate Chinese characters and English words
  // 中文字符每个算一个词 / Each Chinese character counts as one word
  const chineseChars = cleanText.match(/[\u4e00-\u9fa5]/g) || [];
  
  // 英文单词（连续的字母数字字符）/ English words (consecutive alphanumeric characters)
  const englishWords = cleanText.match(/[a-zA-Z0-9]+/g) || [];

  return chineseChars.length + englishWords.length;
}

/**
 * 获取文档统计信息
 * Get document statistics
 *
 * @param options - 获取选项 / Get options
 * @returns 文档统计信息 / Document statistics
 *
 * @example
 * ```typescript
 * // 获取基本统计信息
 * const stats = await getDocumentStats();
 * console.log(`字符数: ${stats.characterCount}`);
 * console.log(`单词数: ${stats.wordCount}`);
 * console.log(`段落数: ${stats.paragraphCount}`);
 *
 * // 获取包含页眉页脚的完整统计
 * const fullStats = await getDocumentStats({
 *   includeHeaderFooter: true,
 *   includeNotes: true,
 *   includeHeadingStats: true
 * });
 * console.log(`总页数: ${fullStats.pageCount}`);
 * console.log(`标题数: ${fullStats.totalHeadingCount}`);
 * ```
 */
export async function getDocumentStats(
  options: GetDocumentStatsOptions = {}
): Promise<DocumentStats> {
  const {
    includeHeaderFooter = false,
    includeNotes = false,
    includeHeadingStats = true,
  } = options;

  return Word.run(async (context) => {
    try {
      // 初始化统计数据 / Initialize statistics
      const stats: DocumentStats = {
        characterCount: 0,
        characterCountNoSpaces: 0,
        wordCount: 0,
        paragraphCount: 0,
        pageCount: 0,
        sectionCount: 0,
        tableCount: 0,
        imageCount: 0,
        inlinePictureCount: 0,
        contentControlCount: 0,
        listCount: 0,
        footnoteCount: 0,
        endnoteCount: 0,
        headingCounts: {},
        totalHeadingCount: 0,
      };

      // 获取文档主体 / Get document body
      const body = context.document.body;
      
      // 批量加载所有需要的集合 / Batch load all required collections
      const paragraphs = body.paragraphs;
      const tables = body.tables;
      const inlinePictures = body.inlinePictures;
      const contentControls = body.contentControls;
      const sections = context.document.sections;

      // 加载集合 / Load collections
      paragraphs.load("items");
      tables.load("items");
      inlinePictures.load("items");
      contentControls.load("items");
      sections.load("items");
      body.load("text");

      await context.sync();

      // 统计基本文本信息 / Count basic text information
      const bodyText = body.text || "";
      stats.characterCount = bodyText.length;
      stats.characterCountNoSpaces = bodyText.replace(/\s/g, "").length;
      stats.wordCount = countWords(bodyText);

      // 统计段落数 / Count paragraphs
      stats.paragraphCount = paragraphs.items.length;

      // 统计表格数 / Count tables
      stats.tableCount = tables.items.length;

      // 统计内嵌图片数 / Count inline pictures
      stats.inlinePictureCount = inlinePictures.items.length;

      // 统计内容控件数 / Count content controls
      stats.contentControlCount = contentControls.items.length;

      // 统计节数 / Count sections
      stats.sectionCount = sections.items.length;

      // 统计列表和标题 / Count lists and headings
      if (includeHeadingStats || stats.listCount === 0) {
        const listIdSet = new Set<number>();

        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          para.load("style,isListItem");
        }

        await context.sync();

        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];

          // 统计列表 / Count lists
          if (para.isListItem) {
            try {
              // 通过 paragraph 的 listOrNullObject 属性获取列表对象
              // Get list object through paragraph's listOrNullObject property
              const list = para.listOrNullObject;
              list.load("id,isNullObject");
              await context.sync();
              if (!list.isNullObject) {
                listIdSet.add(list.id);
              }
            } catch {
              // 忽略错误 / Ignore errors
            }
          }

          // 统计标题 / Count headings
          if (includeHeadingStats && isHeadingStyle(para.style)) {
            const level = extractHeadingLevel(para.style);
            stats.headingCounts[level] = (stats.headingCounts[level] || 0) + 1;
            stats.totalHeadingCount++;
          }
        }

        stats.listCount = listIdSet.size;
      }

      // 统计图片数（通过形状获取）/ Count images (via shapes)
      try {
        const images = body.getRange().getRange("Whole").getRange().inlinePictures;
        images.load("items");
        await context.sync();
        stats.imageCount = images.items.length;
      } catch {
        // 如果获取失败，使用内嵌图片数作为图片数 / If failed, use inline picture count
        stats.imageCount = stats.inlinePictureCount;
      }

      // 统计页眉页脚内容 / Count header footer content
      if (includeHeaderFooter) {
        let headerFooterText = "";

        for (let i = 0; i < sections.items.length; i++) {
          const section = sections.items[i];

          try {
            // 获取页眉 / Get headers
            const headerPrimary = section.getHeader(Word.HeaderFooterType.primary);
            const headerFirst = section.getHeader(Word.HeaderFooterType.firstPage);
            const headerEven = section.getHeader(Word.HeaderFooterType.evenPages);

            // 获取页脚 / Get footers
            const footerPrimary = section.getFooter(Word.HeaderFooterType.primary);
            const footerFirst = section.getFooter(Word.HeaderFooterType.firstPage);
            const footerEven = section.getFooter(Word.HeaderFooterType.evenPages);

            // 加载文本 / Load text
            headerPrimary.load("text");
            headerFirst.load("text");
            headerEven.load("text");
            footerPrimary.load("text");
            footerFirst.load("text");
            footerEven.load("text");

            await context.sync();

            // 统计页眉页脚文本 / Count header footer text
            const headerFooterTexts = [
              headerPrimary.text,
              headerFirst.text,
              headerEven.text,
              footerPrimary.text,
              footerFirst.text,
              footerEven.text,
            ];

            for (const text of headerFooterTexts) {
              if (text) {
                headerFooterText += text;
              }
            }
          } catch {
            // 忽略页眉页脚获取错误 / Ignore header footer errors
          }
        }

        stats.characterCount += headerFooterText.length;
        stats.characterCountNoSpaces += headerFooterText.replace(/\s/g, "").length;
        stats.wordCount += countWords(headerFooterText);
      }

      // 统计脚注和尾注 / Count footnotes and endnotes
      if (includeNotes) {
        try {
          const footnotes = body.footnotes;
          const endnotes = body.endnotes;

          footnotes.load("items");
          endnotes.load("items");

          await context.sync();

          stats.footnoteCount = footnotes.items.length;
          stats.endnoteCount = endnotes.items.length;

          // 统计脚注尾注的文本 / Count footnote endnote text
          let notesText = "";

          for (const footnote of footnotes.items) {
            footnote.body.load("text");
          }
          for (const endnote of endnotes.items) {
            endnote.body.load("text");
          }

          await context.sync();

          for (const footnote of footnotes.items) {
            const text = footnote.body.text || "";
            notesText += text;
          }
          for (const endnote of endnotes.items) {
            const text = endnote.body.text || "";
            notesText += text;
          }

          stats.characterCount += notesText.length;
          stats.characterCountNoSpaces += notesText.replace(/\s/g, "").length;
          stats.wordCount += countWords(notesText);
        } catch (error) {
          console.warn("获取脚注尾注失败 / Failed to get footnotes/endnotes:", error);
          // 继续执行，不影响其他统计 / Continue, don't affect other statistics
        }
      }

      // 估算页数（基于字符数，每页约1800字符）
      // Estimate page count (based on character count, ~1800 chars per page)
      stats.pageCount = Math.max(1, Math.ceil(stats.characterCount / 1800));

      return stats;
    } catch (error) {
      console.error("获取文档统计信息失败 / Failed to get document stats:", error);
      throw new Error(
        `获取文档统计信息失败 / Failed to get document stats: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  });
}

/**
 * 获取简化的文档统计信息（仅包含基本统计）
 * Get simplified document statistics (basic stats only)
 *
 * @returns 简化的文档统计信息 / Simplified document statistics
 */
export async function getBasicDocumentStats(): Promise<{
  characterCount: number;
  wordCount: number;
  paragraphCount: number;
  pageCount: number;
}> {
  return Word.run(async (context) => {
    try {
      const body = context.document.body;
      const paragraphs = body.paragraphs;

      body.load("text");
      paragraphs.load("items");

      await context.sync();

      const bodyText = body.text || "";
      const characterCount = bodyText.length;
      const wordCount = countWords(bodyText);
      const paragraphCount = paragraphs.items.length;
      const pageCount = Math.max(1, Math.ceil(characterCount / 1800));

      return {
        characterCount,
        wordCount,
        paragraphCount,
        pageCount,
      };
    } catch (error) {
      console.error("获取基本统计信息失败 / Failed to get basic stats:", error);
      throw new Error(
        `获取基本统计信息失败 / Failed to get basic stats: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  });
}

/**
 * 格式化文档统计信息为可读文本
 * Format document statistics as readable text
 *
 * @param stats - 文档统计信息 / Document statistics
 * @returns 格式化的统计文本 / Formatted statistics text
 */
export function formatDocumentStats(stats: DocumentStats): string {
  const lines: string[] = [
    `文档统计信息 / Document Statistics`,
    `${"=".repeat(50)}`,
    ``,
    `基本统计 / Basic Statistics:`,
    `  字符数（含空格）/ Characters (with spaces): ${stats.characterCount.toLocaleString()}`,
    `  字符数（不含空格）/ Characters (no spaces): ${stats.characterCountNoSpaces.toLocaleString()}`,
    `  单词数 / Words: ${stats.wordCount.toLocaleString()}`,
    `  段落数 / Paragraphs: ${stats.paragraphCount.toLocaleString()}`,
    `  页数（估算）/ Pages (estimated): ${stats.pageCount}`,
    ``,
    `结构统计 / Structure Statistics:`,
    `  节数 / Sections: ${stats.sectionCount}`,
    `  表格数 / Tables: ${stats.tableCount}`,
    `  图片数 / Images: ${stats.imageCount}`,
    `  内容控件数 / Content Controls: ${stats.contentControlCount}`,
    `  列表数 / Lists: ${stats.listCount}`,
  ];

  if (stats.footnoteCount > 0 || stats.endnoteCount > 0) {
    lines.push(``);
    lines.push(`注释统计 / Notes Statistics:`);
    lines.push(`  脚注数 / Footnotes: ${stats.footnoteCount}`);
    lines.push(`  尾注数 / Endnotes: ${stats.endnoteCount}`);
  }

  if (stats.totalHeadingCount > 0) {
    lines.push(``);
    lines.push(`标题统计 / Heading Statistics:`);
    lines.push(`  总标题数 / Total Headings: ${stats.totalHeadingCount}`);
    
    const levels = Object.keys(stats.headingCounts)
      .map(Number)
      .sort((a, b) => a - b);
    
    for (const level of levels) {
      lines.push(`  标题 ${level} / Heading ${level}: ${stats.headingCounts[level]}`);
    }
  }

  return lines.join("\n");
}
