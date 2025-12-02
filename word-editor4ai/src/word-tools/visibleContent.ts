/**
 * 文件名: visibleContent.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取用户可见范围内容的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

import type {
  PageInfo,
  AnyContentElement,
  ParagraphElement,
  TableElement,
  InlinePictureElement,
  ContentControlElement,
  TableCellInfo,
} from "./types";

/**
 * 获取可见内容的选项
 */
export interface GetVisibleContentOptions {
  includeText?: boolean; // 是否包含文本内容，默认为 true
  includeImages?: boolean; // 是否包含图片信息，默认为 true
  includeTables?: boolean; // 是否包含表格信息，默认为 true
  includeContentControls?: boolean; // 是否包含内容控件，默认为 true
  detailedMetadata?: boolean; // 是否包含详细的元数据，默认为 false
  maxTextLength?: number; // 文本内容的最大长度，默认不限制
}

/**
 * 获取用户当前可见范围的内容
 * @param options 获取选项
 * @returns Promise<PageInfo[]> 可见页面的内容列表
 */
export async function getVisibleContent(
  options: GetVisibleContentOptions = {}
): Promise<PageInfo[]> {
  const {
    includeText = true,
    includeImages = true,
    includeTables = true,
    includeContentControls = true,
    detailedMetadata = false,
    maxTextLength,
  } = options;

  try {
    const pageInfoList: PageInfo[] = [];

    await Word.run(async (context) => {
      // 获取活动窗口
      const activeWindow = context.document.activeWindow;

      // 获取活动窗格
      const activePane = activeWindow.activePane;

      // 获取视口中的页面集合
      const pages = activePane.pagesEnclosingViewport;
      pages.load("items");

      await context.sync();

      console.log(`检测到 ${pages.items.length} 个可见页面`);

      // 批量加载所有页面的基本信息 / Batch load basic info for all pages
      const pageRanges: Word.Range[] = [];
      for (let i = 0; i < pages.items.length; i++) {
        const page = pages.items[i];
        page.load("index");
        const pageRange = page.getRange();
        // eslint-disable-next-line office-addins/no-navigational-load -- 必须预加载导航属性以优化后续批量操作性能 / Must preload navigation properties to optimize subsequent batch operations
        pageRange.load("text,paragraphs,tables,contentControls");
        pageRanges.push(pageRange);
      }

      await context.sync();

      // 批量加载所有页面的段落、表格和内容控件 / Batch load paragraphs, tables and content controls for all pages
      for (const pageRange of pageRanges) {
        const paragraphs = pageRange.paragraphs;
        paragraphs.load("items");

        if (includeTables) {
          const tables = pageRange.tables;
          tables.load("items");
        }

        if (includeContentControls) {
          const contentControls = pageRange.contentControls;
          contentControls.load("items");
        }
      }

      await context.sync();

      // 批量加载段落和图片的详细信息 / Batch load detailed paragraph and image information
      for (const pageRange of pageRanges) {
        for (const paragraph of pageRange.paragraphs.items) {
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
      }

      await context.sync();

      // 批量加载图片和表格单元格的详细信息 / Batch load detailed image and table cell information
      for (const pageRange of pageRanges) {
        // 加载图片属性 / Load image properties
        if (includeImages) {
          for (const paragraph of pageRange.paragraphs.items) {
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
      }

      await context.sync();

      // 批量加载单元格属性 / Batch load cell properties
      if (includeTables && includeText && detailedMetadata) {
        for (const pageRange of pageRanges) {
          for (const table of pageRange.tables.items) {
            for (const row of table.rows.items) {
              for (const cell of row.cells.items) {
                cell.load("value,width");
              }
            }
          }
        }

        await context.sync();
      }

      // 遍历每个可见页面并构建数据结构 / Iterate through each visible page and build data structure
      for (let i = 0; i < pages.items.length; i++) {
        const page = pages.items[i];
        const pageRange = pageRanges[i];

        const pageInfo: PageInfo = {
          index: page.index,
          elements: [],
          text: includeText ? pageRange.text : undefined,
        };

        // 处理段落 / Process paragraphs
        for (const paragraph of pageRange.paragraphs.items) {
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
                console.warn("获取内联图片失败:", error);
              }
            }
          } catch (error) {
            console.warn("处理段落失败:", error);
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
                  console.warn("获取表格单元格内容失败:", error);
                }
              }

              pageInfo.elements.push(tableElement);
            }
          } catch (error) {
            console.warn("获取表格失败:", error);
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
            console.warn("获取内容控件失败:", error);
          }
        }

        pageInfoList.push(pageInfo);
      }
    });

    return pageInfoList;
  } catch (error) {
    console.error("获取可见内容失败:", error);
    throw error;
  }
}

/**
 * 获取当前可见内容的简化版本（仅文本）
 * @returns Promise<string> 可见内容的文本
 */
export async function getVisibleText(): Promise<string> {
  try {
    const pages = await getVisibleContent({
      includeText: true,
      includeImages: false,
      includeTables: false,
      includeContentControls: false,
      detailedMetadata: false,
    });

    return pages.map((page) => page.text || "").join("\n\n");
  } catch (error) {
    console.error("获取可见文本失败:", error);
    throw error;
  }
}

/**
 * 获取当前可见内容的统计信息
 * @returns Promise<{ pageCount: number; elementCount: number; characterCount: number }>
 */
export async function getVisibleContentStats(): Promise<{
  pageCount: number;
  elementCount: number;
  characterCount: number;
  paragraphCount: number;
  tableCount: number;
  imageCount: number;
  contentControlCount: number;
}> {
  try {
    const pages = await getVisibleContent({
      includeText: true,
      includeImages: true,
      includeTables: true,
      includeContentControls: true,
      detailedMetadata: false,
    });

    let elementCount = 0;
    let characterCount = 0;
    let paragraphCount = 0;
    let tableCount = 0;
    let imageCount = 0;
    let contentControlCount = 0;

    for (const page of pages) {
      elementCount += page.elements.length;
      characterCount += page.text?.length || 0;

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
    }

    return {
      pageCount: pages.length,
      elementCount,
      characterCount,
      paragraphCount,
      tableCount,
      imageCount,
      contentControlCount,
    };
  } catch (error) {
    console.error("获取可见内容统计信息失败:", error);
    throw error;
  }
}
