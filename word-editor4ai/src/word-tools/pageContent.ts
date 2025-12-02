/**
 * 文件名: pageContent.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取指定页面内容的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

import type {
  PageInfo,
  ParagraphElement,
  TableElement,
  InlinePictureElement,
  ContentControlElement,
  TableCellInfo,
} from "./types";

/**
 * 获取页面内容的选项 / Get Page Content Options
 */
export interface GetPageContentOptions {
  /** 是否包含文本内容，默认为 true / Include text content, default true */
  includeText?: boolean;
  /** 是否包含图片信息，默认为 true / Include image info, default true */
  includeImages?: boolean;
  /** 是否包含表格信息，默认为 true / Include table info, default true */
  includeTables?: boolean;
  /** 是否包含内容控件，默认为 true / Include content controls, default true */
  includeContentControls?: boolean;
  /** 是否包含详细的元数据，默认为 false / Include detailed metadata, default false */
  detailedMetadata?: boolean;
  /** 文本内容的最大长度，默认不限制 / Max text length, default unlimited */
  maxTextLength?: number;
}

/**
 * 获取指定页面的内容
 * Get content of a specific page
 *
 * @param pageNumber - 页面编号（从1开始）/ Page number (1-based)
 * @param options - 获取选项 / Get options
 * @returns Promise<PageInfo> 页面内容信息，其中 index 是 0-based / Page content information, where index is 0-based
 *
 * @remarks
 * 注意：Word API 的 page.index 是 1-based（从1开始），但返回的 PageInfo.index 会转换为 0-based（从0开始）以保持内部一致性。
 * Note: Word API's page.index is 1-based (starts from 1), but the returned PageInfo.index is converted to 0-based (starts from 0) for internal consistency.
 *
 * @example
 * ```typescript
 * // 获取第1页的所有内容
 * const page = await getPageContent(1, {
 *   includeText: true,
 *   includeImages: true,
 *   includeTables: true,
 *   detailedMetadata: true
 * });
 *
 * console.log(`页面 ${page.index + 1} 包含 ${page.elements.length} 个元素`);
 * ```
 */
export async function getPageContent(
  pageNumber: number,
  options: GetPageContentOptions = {}
): Promise<PageInfo> {
  const {
    includeText = true,
    includeImages = true,
    includeTables = true,
    includeContentControls = true,
    detailedMetadata = false,
    maxTextLength,
  } = options;

  // 验证页面编号 / Validate page number
  if (pageNumber < 1) {
    throw new Error("页面编号必须大于等于1 / Page number must be greater than or equal to 1");
  }

  try {
    return await Word.run(async (context) => {
      // 获取文档的所有页面 / Get all pages in the document
      // 通过 body 的 range 来获取 pages / Get pages through body's range
      const bodyRange = context.document.body.getRange();
      const pages = bodyRange.pages;
      pages.load("items");
      await context.sync();

      // 检查页面是否存在 / Check if page exists
      const pageIndex = pageNumber - 1; // 转换为0基索引 / Convert to 0-based index
      if (pageIndex >= pages.items.length) {
        throw new Error(
          `页面 ${pageNumber} 不存在，文档共有 ${pages.items.length} 页 / Page ${pageNumber} does not exist, document has ${pages.items.length} pages`
        );
      }

      const page = pages.items[pageIndex];
      // 注意：Word API 的 page.index 是 1-based（从1开始），不是 0-based
      // Note: Word API's page.index is 1-based (starts from 1), not 0-based
      page.load("index");

      // 获取页面范围 / Get page range
      const pageRange = page.getRange();
      // eslint-disable-next-line office-addins/no-navigational-load -- 必须预加载导航属性以优化后续批量操作性能 / Must preload navigation properties to optimize subsequent batch operations
      pageRange.load("text,paragraphs,tables,contentControls");
      await context.sync();

      // 加载段落集合 / Load paragraph collection
      const paragraphs = pageRange.paragraphs;
      paragraphs.load("items");

      // 加载表格集合 / Load table collection
      if (includeTables) {
        const tables = pageRange.tables;
        tables.load("items");
      }

      // 加载内容控件集合 / Load content control collection
      if (includeContentControls) {
        const contentControls = pageRange.contentControls;
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
        for (const table of pageRange.tables.items) {
          table.load("rowCount");
          table.columns.load("items");
          if (includeText && detailedMetadata) {
            table.rows.load("items");
          }
        }
      }

      // 批量加载内容控件属性 / Batch load content control properties
      if (includeContentControls) {
        for (const control of pageRange.contentControls.items) {
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
        for (const table of pageRange.tables.items) {
          for (const row of table.rows.items) {
            row.cells.load("items");
          }
        }
      }

      await context.sync();

      // 批量加载单元格属性 / Batch load cell properties
      if (includeTables && includeText && detailedMetadata) {
        for (const table of pageRange.tables.items) {
          for (const row of table.rows.items) {
            for (const cell of row.cells.items) {
              cell.load("value,width");
            }
          }
        }

        await context.sync();
      }

      // 构建页面信息 / Build page information
      // 注意：page.index 是 1-based，我们需要转换为 0-based 以保持一致性
      // Note: page.index is 1-based, we need to convert to 0-based for consistency
      const pageInfo: PageInfo = {
        index: page.index - 1, // 将 1-based 转换为 0-based / Convert 1-based to 0-based
        elements: [],
        text: includeText ? pageRange.text : undefined,
      };

      // 处理段落 / Process paragraphs
      for (const paragraph of paragraphs.items) {
        try {
          let paragraphText = paragraph.text;
          if (maxTextLength && paragraphText.length > maxTextLength) {
            paragraphText = paragraphText.substring(0, maxTextLength) + "...";
          }

          const paragraphElement: ParagraphElement = {
            id: `para-${page.index}-${pageInfo.elements.length}`,
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

          pageInfo.elements.push(paragraphElement);

          // 检查段落中的内联图片 / Check inline pictures in paragraph
          if (includeImages) {
            try {
              const inlinePictures = paragraph.inlinePictures;

              for (const picture of inlinePictures.items) {
                const imageElement: InlinePictureElement = {
                  id: `img-${page.index}-${pageInfo.elements.length}`,
                  type: "InlinePicture",
                  width: picture.width,
                  height: picture.height,
                  altText: picture.altTextTitle || picture.altTextDescription,
                  hyperlink: picture.hyperlink,
                };

                pageInfo.elements.push(imageElement);
              }
            } catch (error) {
              console.warn("获取内联图片失败 / Failed to get inline pictures:", error);
            }
          }
        } catch (error) {
          console.warn("处理段落失败 / Failed to process paragraph:", error);
        }
      }

      // 获取页面中的表格 / Get tables in the page
      if (includeTables) {
        try {
          for (const table of pageRange.tables.items) {
            const columns = table.columns;

            const tableElement: TableElement = {
              id: `table-${page.index}-${pageInfo.elements.length}`,
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

            pageInfo.elements.push(tableElement);
          }
        } catch (error) {
          console.warn("获取表格失败 / Failed to get tables:", error);
        }
      }

      // 获取页面中的内容控件 / Get content controls in the page
      if (includeContentControls) {
        try {
          for (const control of pageRange.contentControls.items) {
            let controlText = control.text;
            if (maxTextLength && controlText.length > maxTextLength) {
              controlText = controlText.substring(0, maxTextLength) + "...";
            }

            const controlElement: ContentControlElement = {
              id: `ctrl-${page.index}-${pageInfo.elements.length}`,
              type: "ContentControl",
              text: includeText ? controlText : undefined,
              title: control.title,
              tag: control.tag,
              controlType: control.type,
              cannotDelete: detailedMetadata ? control.cannotDelete : undefined,
              cannotEdit: detailedMetadata ? control.cannotEdit : undefined,
              placeholderText: detailedMetadata ? control.placeholderText : undefined,
            };

            pageInfo.elements.push(controlElement);
          }
        } catch (error) {
          console.warn("获取内容控件失败 / Failed to get content controls:", error);
        }
      }

      return pageInfo;
    });
  } catch (error) {
    console.error("获取页面内容失败 / Failed to get page content:", error);
    throw error;
  }
}

/**
 * 获取指定页面的纯文本内容
 * Get plain text content of a specific page
 *
 * @param pageNumber - 页面编号（从1开始）/ Page number (1-based)
 * @returns Promise<string> 页面的文本内容 / Text content of the page
 *
 * @example
 * ```typescript
 * const text = await getPageText(1);
 * console.log(`第1页内容: ${text}`);
 * ```
 */
export async function getPageText(pageNumber: number): Promise<string> {
  try {
    const page = await getPageContent(pageNumber, {
      includeText: true,
      includeImages: false,
      includeTables: false,
      includeContentControls: false,
      detailedMetadata: false,
    });

    return page.text || "";
  } catch (error) {
    console.error("获取页面文本失败 / Failed to get page text:", error);
    throw error;
  }
}

/**
 * 获取指定页面的统计信息
 * Get statistics of a specific page
 *
 * @param pageNumber - 页面编号（从1开始）/ Page number (1-based)
 * @returns Promise<PageStats> 页面统计信息 / Page statistics
 *
 * @example
 * ```typescript
 * const stats = await getPageStats(1);
 * console.log(`第1页包含 ${stats.elementCount} 个元素，${stats.characterCount} 个字符`);
 * ```
 */
export async function getPageStats(pageNumber: number): Promise<{
  pageIndex: number;
  elementCount: number;
  characterCount: number;
  paragraphCount: number;
  tableCount: number;
  imageCount: number;
  contentControlCount: number;
}> {
  try {
    const page = await getPageContent(pageNumber, {
      includeText: true,
      includeImages: true,
      includeTables: true,
      includeContentControls: true,
      detailedMetadata: false,
    });

    let paragraphCount = 0;
    let tableCount = 0;
    let imageCount = 0;
    let contentControlCount = 0;

    for (const element of page.elements) {
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
        case "ContentControl":
          contentControlCount++;
          break;
      }
    }

    return {
      pageIndex: page.index,
      elementCount: page.elements.length,
      characterCount: page.text?.length || 0,
      paragraphCount,
      tableCount,
      imageCount,
      contentControlCount,
    };
  } catch (error) {
    console.error("获取页面统计信息失败 / Failed to get page statistics:", error);
    throw error;
  }
}
