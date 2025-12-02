/**
 * 文件名: selectedContent.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取选中内容的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

import type {
  ParagraphElement,
  TableElement,
  InlinePictureElement,
  ContentControlElement,
  AnyContentElement,
  TableCellInfo,
} from "./types";

/**
 * 选中内容信息 / Selected Content Info
 */
export interface ContentInfo {
  /** 选中的文本内容 / Selected text content */
  text: string;
  /** 选中内容的元素列表 / List of elements in selection */
  elements: AnyContentElement[];
  /** 选中范围的元数据 / Selection range metadata */
  metadata?: {
    /** 是否为空选择 / Is empty selection */
    isEmpty: boolean;
    /** 选中内容的字符数 / Character count */
    characterCount: number;
    /** 选中内容的段落数 / Paragraph count */
    paragraphCount: number;
    /** 选中内容的表格数 / Table count */
    tableCount: number;
    /** 选中内容的图片数 / Image count */
    imageCount: number;
  };
}

/**
 * 获取选中内容的选项 / Get Selected Content Options
 */
export interface GetSelectedContentOptions {
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
 * 获取当前选中的内容
 * Get currently selected content
 *
 * @param options - 获取选项 / Get options
 * @returns Promise<ContentInfo> 选中内容信息 / Selected content information
 *
 * @remarks
 * 此函数获取用户当前在文档中选中的内容，包括文本、段落、表格、图片等元素。
 * This function gets the content currently selected by the user in the document, including text, paragraphs, tables, images, etc.
 *
 * @example
 * ```typescript
 * // 获取选中内容的所有信息
 * const selection = await getSelectedContent({
 *   includeText: true,
 *   includeImages: true,
 *   includeTables: true,
 *   detailedMetadata: true
 * });
 *
 * console.log(`选中了 ${selection.text.length} 个字符`);
 * console.log(`包含 ${selection.elements.length} 个元素`);
 * ```
 */
export async function getSelectedContent(
  options: GetSelectedContentOptions = {}
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
      // 获取当前选中的范围 / Get current selection range
      const selection = context.document.getSelection();
      
      // 加载选中范围的基本属性 / Load basic properties of selection
      // eslint-disable-next-line office-addins/no-navigational-load -- 必须预加载导航属性以优化后续批量操作性能 / Must preload navigation properties to optimize subsequent batch operations
      selection.load("text,isEmpty,paragraphs,tables,contentControls");
      await context.sync();

      // 检查是否有选中内容 / Check if there is a selection
      if (selection.isEmpty) {
        return {
          text: "",
          elements: [],
          metadata: {
            isEmpty: true,
            characterCount: 0,
            paragraphCount: 0,
            tableCount: 0,
            imageCount: 0,
          },
        };
      }

      // 加载段落集合 / Load paragraph collection
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");

      // 加载表格集合 / Load table collection
      if (includeTables) {
        const tables = selection.tables;
        tables.load("items");
      }

      // 加载内容控件集合 / Load content control collection
      if (includeContentControls) {
        const contentControls = selection.contentControls;
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
        for (const table of selection.tables.items) {
          table.load("rowCount");
          table.columns.load("items");
          if (includeText && detailedMetadata) {
            table.rows.load("items");
          }
        }
      }

      // 批量加载内容控件属性 / Batch load content control properties
      if (includeContentControls) {
        for (const control of selection.contentControls.items) {
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
        for (const table of selection.tables.items) {
          for (const row of table.rows.items) {
            row.cells.load("items");
          }
        }
      }

      await context.sync();

      // 批量加载单元格属性 / Batch load cell properties
      if (includeTables && includeText && detailedMetadata) {
        for (const table of selection.tables.items) {
          for (const row of table.rows.items) {
            for (const cell of row.cells.items) {
              cell.load("value,width");
            }
          }
        }

        await context.sync();
      }

      // 构建选中内容信息 / Build selected content information
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
            id: `sel-para-${elementIdCounter++}`,
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
                  id: `sel-img-${elementIdCounter++}`,
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

      // 获取选中范围中的表格 / Get tables in the selection
      if (includeTables) {
        try {
          for (const table of selection.tables.items) {
            const columns = table.columns;

            const tableElement: TableElement = {
              id: `sel-table-${elementIdCounter++}`,
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

      // 获取选中范围中的内容控件 / Get content controls in the selection
      if (includeContentControls) {
        try {
          for (const control of selection.contentControls.items) {
            let controlText = control.text;
            if (maxTextLength && controlText.length > maxTextLength) {
              controlText = controlText.substring(0, maxTextLength) + "...";
            }

            const controlElement: ContentControlElement = {
              id: `sel-ctrl-${elementIdCounter++}`,
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
      let selectionText = selection.text;
      if (maxTextLength && selectionText.length > maxTextLength) {
        selectionText = selectionText.substring(0, maxTextLength) + "...";
      }

      return {
        text: includeText ? selectionText : "",
        elements,
        metadata: {
          isEmpty: false,
          characterCount: selection.text.length,
          paragraphCount,
          tableCount,
          imageCount,
        },
      };
    });
  } catch (error) {
    console.error("获取选中内容失败 / Failed to get selected content:", error);
    throw error;
  }
}
