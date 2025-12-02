/**
 * 文件名: textBoxContent.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取文本框内容的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

import type {
  TextBoxInfo,
  GetTextBoxOptions,
  ParagraphElement,
} from "./types";

// 重新导出类型供外部使用 / Re-export types for external use
export type { TextBoxInfo, GetTextBoxOptions };

/**
 * 获取文本框内容
 * Get text box content
 *
 * @param options - 获取选项 / Get options
 * @returns Promise<TextBoxInfo[]> 文本框信息列表 / Text box information list
 *
 * @remarks
 * 此函数按以下优先级获取文本框：
 * 1. 如果用户有选择，优先返回选择范围内的文本框
 * 2. 如果没有选择，返回可见区域内的文本框
 * 3. 所有场景都可能返回多个文本框
 *
 * This function gets text boxes in the following priority:
 * 1. If user has a selection, return text boxes in the selection
 * 2. If no selection, return text boxes in the visible area
 * 3. All scenarios may return multiple text boxes
 *
 * @example
 * ```typescript
 * // 获取文本框内容
 * const textBoxes = await getTextBoxes({
 *   includeText: true,
 *   includeParagraphs: true,
 *   detailedMetadata: true
 * });
 *
 * console.log(`找到 ${textBoxes.length} 个文本框`);
 * textBoxes.forEach(box => {
 *   console.log(`文本框: ${box.name}, 内容: ${box.text}`);
 * });
 * ```
 */
export async function getTextBoxes(
  options: GetTextBoxOptions = {}
): Promise<TextBoxInfo[]> {
  const {
    includeText = true,
    includeParagraphs = false,
    detailedMetadata = false,
    maxTextLength,
  } = options;

  try {
    return await Word.run(async (context) => {
      // 获取当前选中的范围 / Get current selection range
      const selection = context.document.getSelection();
      selection.load("isEmpty");
      await context.sync();

      let shapes: Word.ShapeCollection;
      let rangeType: "selection" | "visible";

      // 判断是否有选择 / Check if there is a selection
      if (!selection.isEmpty) {
        // 有选择，从文档主体获取所有形状，稍后过滤选择范围内的 / Has selection, get all shapes from body, filter later
        shapes = context.document.body.shapes;
        rangeType = "selection";
      } else {
        // 没有选择，从文档主体获取所有形状 / No selection, get all shapes from body
        shapes = context.document.body.shapes;
        rangeType = "visible";
      }

      // 加载形状集合 / Load shapes collection
      shapes.load("items");
      await context.sync();

      // 过滤出文本框类型的形状 / Filter shapes to get text boxes
      const textBoxShapes: Word.Shape[] = [];
      for (const shape of shapes.items) {
        shape.load("type");
      }
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.type === Word.ShapeType.textBox) {
          textBoxShapes.push(shape);
        }
      }

      if (textBoxShapes.length === 0) {
        console.log(`在${rangeType === "selection" ? "选择范围" : "可见区域"}内未找到文本框 / No text boxes found in ${rangeType === "selection" ? "selection" : "visible area"}`);
        return [];
      }

      console.log(`在${rangeType === "selection" ? "选择范围" : "可见区域"}内找到 ${textBoxShapes.length} 个文本框 / Found ${textBoxShapes.length} text boxes in ${rangeType === "selection" ? "selection" : "visible area"}`);

      // 批量加载文本框的基本属性 / Batch load basic properties of text boxes
      for (const shape of textBoxShapes) {
        shape.load("id,name,width,height,left,top,rotation,visible,lockAspectRatio");
        if (includeText || includeParagraphs) {
          shape.load("body");
        }
      }
      await context.sync();

      // 批量加载文本框的文本内容 / Batch load text content of text boxes
      if (includeText || includeParagraphs) {
        for (const shape of textBoxShapes) {
          try {
            const body = shape.body;
            body.load("text");
            if (includeParagraphs) {
              body.load("paragraphs");
            }
          } catch (error) {
            console.warn(`加载文本框 ${shape.name} 的文本内容失败 / Failed to load text content of text box ${shape.name}:`, error);
          }
        }
        await context.sync();
      }

      // 批量加载段落详情 / Batch load paragraph details
      // 注意：文本框的段落访问可能受限，使用 try-catch 包裹每个操作
      // Note: Paragraph access in text boxes may be limited, wrap each operation with try-catch
      const shapesWithParagraphs: Word.Shape[] = [];
      if (includeParagraphs) {
        for (const shape of textBoxShapes) {
          try {
            const body = shape.body;
            const paragraphs = body.paragraphs;
            paragraphs.load("items");
            shapesWithParagraphs.push(shape);
          } catch (error) {
            console.warn(`加载文本框 ${shape.name} 段落失败，将跳过段落详情 / Failed to load paragraphs for text box ${shape.name}, will skip paragraph details:`, error);
          }
        }
        
        if (shapesWithParagraphs.length > 0) {
          try {
            await context.sync();
          } catch (error) {
            console.warn(`同步段落集合失败，将跳过所有段落详情 / Failed to sync paragraph collections, will skip all paragraph details:`, error);
            shapesWithParagraphs.length = 0; // 清空数组 / Clear array
          }
        }

        // 加载段落的详细属性 / Load detailed paragraph properties
        if (shapesWithParagraphs.length > 0) {
          for (const shape of shapesWithParagraphs) {
            try {
              const body = shape.body;
              const paragraphs = body.paragraphs;
              for (const paragraph of paragraphs.items) {
                paragraph.load(
                  "text,style,alignment,firstLineIndent,leftIndent,rightIndent,lineSpacing,spaceAfter,spaceBefore,isListItem"
                );
              }
            } catch (error) {
              console.warn(`加载文本框 ${shape.name} 段落详细属性失败 / Failed to load detailed paragraph properties for text box ${shape.name}:`, error);
            }
          }
          
          try {
            await context.sync();
          } catch (error) {
            console.warn(`同步段落详细属性失败 / Failed to sync paragraph details:`, error);
          }
        }
      }

      // 构建文本框信息列表 / Build text box information list
      const textBoxInfoList: TextBoxInfo[] = [];

      for (const shape of textBoxShapes) {
        try {
          const textBoxInfo: TextBoxInfo = {
            id: `textbox-${shape.id}`,
            name: detailedMetadata ? shape.name : undefined,
            width: detailedMetadata ? shape.width : undefined,
            height: detailedMetadata ? shape.height : undefined,
            left: detailedMetadata ? shape.left : undefined,
            top: detailedMetadata ? shape.top : undefined,
            rotation: detailedMetadata ? shape.rotation : undefined,
            visible: detailedMetadata ? shape.visible : undefined,
            lockAspectRatio: detailedMetadata ? shape.lockAspectRatio : undefined,
          };

          // 获取文本内容 / Get text content
          if (includeText || includeParagraphs) {
            try {
              const body = shape.body;
              let text = body.text;

              if (maxTextLength && text.length > maxTextLength) {
                text = text.substring(0, maxTextLength) + "...";
              }

              textBoxInfo.text = includeText ? text : undefined;

              // 获取段落详情 / Get paragraph details
              // 只处理成功加载段落的文本框 / Only process text boxes that successfully loaded paragraphs
              if (includeParagraphs && shapesWithParagraphs.includes(shape)) {
                try {
                  const paragraphs = body.paragraphs;
                  const paragraphElements: ParagraphElement[] = [];

                  for (let i = 0; i < paragraphs.items.length; i++) {
                    const paragraph = paragraphs.items[i];
                    let paragraphText = paragraph.text;

                    if (maxTextLength && paragraphText.length > maxTextLength) {
                      paragraphText = paragraphText.substring(0, maxTextLength) + "...";
                    }

                    const paragraphElement: ParagraphElement = {
                      id: `${textBoxInfo.id}-para-${i}`,
                      type: "Paragraph",
                      text: paragraphText,
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

                    paragraphElements.push(paragraphElement);
                  }

                  textBoxInfo.paragraphs = paragraphElements;
                } catch (error) {
                  console.warn(`获取文本框 ${shape.name} 的段落详情失败 / Failed to get paragraph details for text box ${shape.name}:`, error);
                }
              }
            } catch (error) {
              console.warn(`获取文本框 ${shape.name} 的文本内容失败 / Failed to get text content of text box ${shape.name}:`, error);
            }
          }

          textBoxInfoList.push(textBoxInfo);
        } catch (error) {
          console.warn(`处理文本框失败 / Failed to process text box:`, error);
        }
      }

      return textBoxInfoList;
    });
  } catch (error) {
    console.error("获取文本框内容失败 / Failed to get text box content:", error);
    throw error;
  }
}
