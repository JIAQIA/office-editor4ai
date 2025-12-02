/**
 * 文件名: rangeContent.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取指定范围内容的工具核心逻辑，支持通过书签、标题、段落索引等方式定位
 */

/* global Word, console */

import type {
  AnyContentElement,
  ContentControlElement,
  ContentInfo,
  GetContentOptions,
  InlinePictureElement,
  ParagraphElement,
  RangeLocator,
  TableCellInfo,
  TableElement,
} from "./types";

/**
 * 获取范围内容的选项 / Get Range Content Options
 */
export type GetRangeContentOptions = GetContentOptions;

/**
 * 获取指定范围的内容
 * Get content from specified range
 *
 * @param locator - 范围定位器 / Range locator
 * @param options - 获取选项 / Get options
 * @returns Promise<ContentInfo> 范围内容信息 / Range content information
 *
 * @remarks
 * 此函数根据不同的定位方式（书签、标题、段落索引等）获取文档中指定范围的内容。
 * This function gets content from a specified range in the document using different locator types (bookmark, heading, paragraph index, etc.).
 *
 * @example
 * ```typescript
 * // 通过书签获取内容
 * const bookmarkContent = await getRangeContent(
 *   { type: 'bookmark', name: 'MyBookmark' },
 *   { includeText: true, detailedMetadata: true }
 * );
 *
 * // 通过标题获取内容
 * const headingContent = await getRangeContent(
 *   { type: 'heading', text: '第一章', level: 1 },
 *   { includeText: true }
 * );
 *
 * // 通过段落索引获取内容
 * const paragraphContent = await getRangeContent(
 *   { type: 'paragraph', startIndex: 0, endIndex: 5 },
 *   { includeText: true, includeTables: true }
 * );
 * ```
 */
export async function getRangeContent(
  locator: RangeLocator,
  options: GetRangeContentOptions = {}
): Promise<ContentInfo> {
  const {
    includeText = true,
    includeImages = true,
    includeTables = true,
    includeContentControls = true,
    detailedMetadata = false,
    maxTextLength,
  } = options;

  try {
    return await Word.run(async (context) => {
      let range: Word.Range | null = null;

      // 根据定位器类型获取范围 / Get range based on locator type
      switch (locator.type) {
        case "bookmark":
          range = await getBookmarkRange(context, locator.name);
          break;
        case "heading":
          range = await getHeadingRange(context, locator);
          break;
        case "paragraph":
          range = await getParagraphRange(context, locator);
          break;
        case "section":
          range = await getSectionRange(context, locator.index);
          break;
        case "contentControl":
          range = await getContentControlRange(context, locator);
          break;
        default:
          throw new Error(
            `不支持的定位器类型 / Unsupported locator type: ${(locator as RangeLocator).type}`
          );
      }

      if (!range) {
        throw new Error(`无法找到指定的范围 / Cannot find specified range`);
      }

      // 加载范围的基本属性 / Load basic properties of range

      range.load("text,isEmpty,paragraphs,tables,contentControls");
      await context.sync();

      // 检查范围是否为空 / Check if range is empty
      if (range.isEmpty) {
        return {
          text: "",
          elements: [],
          metadata: {
            isEmpty: true,
            characterCount: 0,
            paragraphCount: 0,
            tableCount: 0,
            imageCount: 0,
            locatorType: locator.type,
          },
        };
      }

      // 加载段落集合 / Load paragraph collection
      const paragraphs = range.paragraphs;
      paragraphs.load("items");

      // 加载表格集合 / Load table collection
      if (includeTables) {
        const tables = range.tables;
        tables.load("items");
      }

      // 加载内容控件集合 / Load content control collection
      if (includeContentControls) {
        const contentControls = range.contentControls;
        contentControls.load("items");
      }

      await context.sync();

      // 批量加载段落详细信息 / Batch load paragraph details
      for (const paragraph of paragraphs.items) {
        paragraph.load(
          "text,style,alignment,firstLineIndent,leftIndent,rightIndent,lineSpacing,spaceAfter,spaceBefore,isListItem"
        );
        if (includeImages) {
          paragraph.inlinePictures.load("items");
        }
      }

      // 批量加载表格属性 / Batch load table properties
      if (includeTables) {
        for (const table of range.tables.items) {
          table.load("rowCount");
          table.columns.load("items");
          if (includeText && detailedMetadata) {
            table.rows.load("items");
          }
        }
      }

      // 批量加载内容控件属性 / Batch load content control properties
      if (includeContentControls) {
        for (const control of range.contentControls.items) {
          control.load("text,title,tag,type,cannotDelete,cannotEdit,placeholderText");
        }
      }

      await context.sync();

      // 批量加载图片和表格单元格的详细信息 / Batch load detailed image and table cell information
      if (includeImages) {
        for (const paragraph of paragraphs.items) {
          for (const picture of paragraph.inlinePictures.items) {
            picture.load("width,height,altTextTitle,altTextDescription,hyperlink");
          }
        }
      }

      // 加载表格单元格 / Load table cells
      if (includeTables && includeText && detailedMetadata) {
        for (const table of range.tables.items) {
          for (const row of table.rows.items) {
            row.cells.load("items");
          }
        }
      }

      await context.sync();

      // 批量加载单元格属性 / Batch load cell properties
      if (includeTables && includeText && detailedMetadata) {
        for (const table of range.tables.items) {
          for (const row of table.rows.items) {
            for (const cell of row.cells.items) {
              cell.load("value,width");
            }
          }
        }

        await context.sync();
      }

      // 构建范围内容信息 / Build range content information
      const elements: AnyContentElement[] = [];
      let elementIdCounter = 0;

      // 处理段落 / Process paragraphs
      for (const paragraph of paragraphs.items) {
        try {
          let paragraphText = paragraph.text;
          if (maxTextLength && paragraphText.length > maxTextLength) {
            paragraphText = paragraphText.substring(0, maxTextLength) + "...";
          }

          const paragraphElement: ParagraphElement = {
            id: `range-para-${elementIdCounter++}`,
            type: "Paragraph",
            text: includeText ? paragraphText : undefined,
            style: detailedMetadata ? paragraph.style : undefined,
            alignment: detailedMetadata ? paragraph.alignment : undefined,
            firstLineIndent: detailedMetadata ? paragraph.firstLineIndent : undefined,
            leftIndent: detailedMetadata ? paragraph.leftIndent : undefined,
            rightIndent: detailedMetadata ? paragraph.rightIndent : undefined,
            lineSpacing: detailedMetadata ? paragraph.lineSpacing : undefined,
            spaceAfter: detailedMetadata ? paragraph.spaceAfter : undefined,
            spaceBefore: detailedMetadata ? paragraph.spaceBefore : undefined,
            isListItem: detailedMetadata ? paragraph.isListItem : undefined,
          };

          elements.push(paragraphElement);

          // 检查段落中的内联图片 / Check inline pictures in paragraph
          if (includeImages) {
            try {
              const inlinePictures = paragraph.inlinePictures;

              for (const picture of inlinePictures.items) {
                const imageElement: InlinePictureElement = {
                  id: `range-img-${elementIdCounter++}`,
                  type: "InlinePicture",
                  width: picture.width,
                  height: picture.height,
                  altText: picture.altTextTitle || picture.altTextDescription,
                  hyperlink: picture.hyperlink,
                };

                elements.push(imageElement);
              }
            } catch (error) {
              console.warn("获取内联图片失败 / Failed to get inline pictures:", error);
            }
          }
        } catch (error) {
          console.warn("处理段落失败 / Failed to process paragraph:", error);
        }
      }

      // 获取范围中的表格 / Get tables in the range
      if (includeTables) {
        try {
          for (const table of range.tables.items) {
            const columns = table.columns;

            const tableElement: TableElement = {
              id: `range-table-${elementIdCounter++}`,
              type: "Table",
              rowCount: table.rowCount,
              columnCount: columns.items.length,
              cells: [],
            };

            // 获取表格单元格内容 / Get table cell content
            if (includeText && detailedMetadata) {
              try {
                const rows = table.rows;
                // 处理单元格数据 / Process cell data
                for (let rowIndex = 0; rowIndex < rows.items.length; rowIndex++) {
                  const row = rows.items[rowIndex];
                  const cells = row.cells;
                  const cellRow: TableCellInfo[] = [];

                  for (let colIndex = 0; colIndex < cells.items.length; colIndex++) {
                    const cell = cells.items[colIndex];

                    let cellText = cell.value;
                    if (maxTextLength && cellText.length > maxTextLength) {
                      cellText = cellText.substring(0, maxTextLength) + "...";
                    }

                    cellRow.push({
                      text: cellText,
                      rowIndex,
                      columnIndex: colIndex,
                      width: cell.width,
                    });
                  }

                  tableElement.cells!.push(cellRow);
                }
              } catch (error) {
                console.warn("获取表格单元格内容失败 / Failed to get table cell content:", error);
              }
            }

            elements.push(tableElement);
          }
        } catch (error) {
          console.warn("获取表格失败 / Failed to get tables:", error);
        }
      }

      // 获取范围中的内容控件 / Get content controls in the range
      if (includeContentControls) {
        try {
          for (const control of range.contentControls.items) {
            let controlText = control.text;
            if (maxTextLength && controlText.length > maxTextLength) {
              controlText = controlText.substring(0, maxTextLength) + "...";
            }

            const controlElement: ContentControlElement = {
              id: `range-ctrl-${elementIdCounter++}`,
              type: "ContentControl",
              text: includeText ? controlText : undefined,
              title: control.title,
              tag: control.tag,
              controlType: control.type,
              cannotDelete: detailedMetadata ? control.cannotDelete : undefined,
              cannotEdit: detailedMetadata ? control.cannotEdit : undefined,
              placeholderText: detailedMetadata ? control.placeholderText : undefined,
            };

            elements.push(controlElement);
          }
        } catch (error) {
          console.warn("获取内容控件失败 / Failed to get content controls:", error);
        }
      }

      // 计算统计信息 / Calculate statistics
      let paragraphCount = 0;
      let tableCount = 0;
      let imageCount = 0;

      for (const element of elements) {
        switch (element.type) {
          case "Paragraph":
            paragraphCount++;
            break;
          case "Table":
            tableCount++;
            break;
          case "Image":
          case "InlinePicture":
            imageCount++;
            break;
        }
      }

      // 处理文本长度限制 / Handle text length limit
      let rangeText = range.text;
      if (maxTextLength && rangeText.length > maxTextLength) {
        rangeText = rangeText.substring(0, maxTextLength) + "...";
      }

      return {
        text: includeText ? rangeText : "",
        elements,
        metadata: {
          isEmpty: false,
          characterCount: range.text.length,
          paragraphCount,
          tableCount,
          imageCount,
          locatorType: locator.type,
        },
      };
    });
  } catch (error) {
    console.error("获取范围内容失败 / Failed to get range content:", error);
    throw error;
  }
}

/**
 * 通过书签名称获取范围 / Get range by bookmark name
 */
async function getBookmarkRange(context: Word.RequestContext, name: string): Promise<Word.Range> {
  // Word API 不支持直接通过名称获取书签，需要遍历所有书签
  // Word API does not support getting bookmarks by name directly, need to iterate all bookmarks
  const body = context.document.body;
  const contentControls = body.contentControls;
  contentControls.load("items");
  await context.sync();

  // 批量加载内容控件属性 / Batch load content control properties
  for (const control of contentControls.items) {
    control.load("title,tag,type");
  }
  await context.sync();

  // 查找匹配的书签（通过内容控件的 tag 或 title）/ Find matching bookmark (via content control tag or title)
  for (const control of contentControls.items) {
    if (control.title === name || control.tag === name) {
      return control.getRange();
    }
  }

  // 如果没有找到匹配的内容控件，尝试使用搜索功能
  // If no matching content control found, try using search
  const searchResults = body.search(name, { matchCase: false, matchWholeWord: false });
  searchResults.load("items");
  await context.sync();

  if (searchResults.items.length > 0) {
    return searchResults.items[0];
  }

  throw new Error(`找不到书签: ${name} / Bookmark not found: ${name}`);
}

/**
 * 通过标题定位器获取范围 / Get range by heading locator
 */
async function getHeadingRange(
  context: Word.RequestContext,
  locator: { text?: string; level?: number; index?: number }
): Promise<Word.Range> {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("items");
  await context.sync();

  // 批量加载段落样式 / Batch load paragraph styles
  for (const paragraph of paragraphs.items) {
    paragraph.load("text,style,styleBuiltIn");
  }
  await context.sync();

  // 查找匹配的标题 / Find matching headings
  const matchingHeadings: Word.Paragraph[] = [];

  for (const paragraph of paragraphs.items) {
    // 检查是否为标题样式 / Check if it's a heading style
    const isHeading =
      paragraph.style.toLowerCase().startsWith("heading") ||
      (paragraph.styleBuiltIn >= Word.BuiltInStyleName.heading1 &&
        paragraph.styleBuiltIn <= Word.BuiltInStyleName.heading9);

    if (!isHeading) continue;

    // 检查标题级别 / Check heading level
    if (locator.level !== undefined) {
      const levelMatch = paragraph.style.match(/heading\s*(\d)/i);
      const paragraphLevel = levelMatch ? parseInt(levelMatch[1]) : 0;
      if (paragraphLevel !== locator.level) continue;
    }

    // 检查标题文本 / Check heading text
    if (locator.text !== undefined) {
      const paragraphText = paragraph.text.trim();
      if (!paragraphText.includes(locator.text)) continue;
    }

    matchingHeadings.push(paragraph);
  }

  if (matchingHeadings.length === 0) {
    throw new Error(`找不到匹配的标题 / No matching heading found`);
  }

  // 根据索引选择标题 / Select heading by index
  const index = locator.index ?? 0;
  if (index >= matchingHeadings.length) {
    throw new Error(
      `标题索引超出范围: ${index} (共 ${matchingHeadings.length} 个) / Heading index out of range: ${index} (total ${matchingHeadings.length})`
    );
  }

  const targetHeading = matchingHeadings[index];
  return targetHeading.getRange();
}

/**
 * 通过段落索引获取范围 / Get range by paragraph index
 */
async function getParagraphRange(
  context: Word.RequestContext,
  locator: { startIndex: number; endIndex?: number }
): Promise<Word.Range> {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("items");
  await context.sync();

  const { startIndex, endIndex = startIndex } = locator;

  if (startIndex < 0 || startIndex >= paragraphs.items.length) {
    throw new Error(
      `起始段落索引超出范围: ${startIndex} (共 ${paragraphs.items.length} 个段落) / Start paragraph index out of range: ${startIndex} (total ${paragraphs.items.length} paragraphs)`
    );
  }

  if (endIndex < startIndex || endIndex >= paragraphs.items.length) {
    throw new Error(
      `结束段落索引超出范围: ${endIndex} (共 ${paragraphs.items.length} 个段落) / End paragraph index out of range: ${endIndex} (total ${paragraphs.items.length} paragraphs)`
    );
  }

  const startParagraph = paragraphs.items[startIndex];
  const endParagraph = paragraphs.items[endIndex];

  const startRange = startParagraph.getRange(Word.RangeLocation.start);
  const endRange = endParagraph.getRange(Word.RangeLocation.end);

  return startRange.expandTo(endRange);
}

/**
 * 通过节索引获取范围 / Get range by section index
 */
async function getSectionRange(context: Word.RequestContext, index: number): Promise<Word.Range> {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  if (index < 0 || index >= sections.items.length) {
    throw new Error(
      `节索引超出范围: ${index} (共 ${sections.items.length} 个节) / Section index out of range: ${index} (total ${sections.items.length} sections)`
    );
  }

  const section = sections.items[index];
  return section.body.getRange();
}

/**
 * 通过内容控件定位器获取范围 / Get range by content control locator
 */
async function getContentControlRange(
  context: Word.RequestContext,
  locator: { title?: string; tag?: string; index?: number }
): Promise<Word.Range> {
  const body = context.document.body;
  const contentControls = body.contentControls;
  contentControls.load("items");
  await context.sync();

  // 批量加载内容控件属性 / Batch load content control properties
  for (const control of contentControls.items) {
    control.load("title,tag");
  }
  await context.sync();

  // 查找匹配的内容控件 / Find matching content controls
  const matchingControls: Word.ContentControl[] = [];

  for (const control of contentControls.items) {
    // 检查标题 / Check title
    if (locator.title !== undefined && control.title !== locator.title) {
      continue;
    }

    // 检查标签 / Check tag
    if (locator.tag !== undefined && control.tag !== locator.tag) {
      continue;
    }

    matchingControls.push(control);
  }

  if (matchingControls.length === 0) {
    throw new Error(`找不到匹配的内容控件 / No matching content control found`);
  }

  // 根据索引选择内容控件 / Select content control by index
  const index = locator.index ?? 0;
  if (index >= matchingControls.length) {
    throw new Error(
      `内容控件索引超出范围: ${index} (共 ${matchingControls.length} 个) / Content control index out of range: ${index} (total ${matchingControls.length})`
    );
  }

  const targetControl = matchingControls[index];
  return targetControl.getRange();
}
