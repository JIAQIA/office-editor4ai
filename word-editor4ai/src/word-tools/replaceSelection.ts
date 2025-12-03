/**
 * 文件名: replaceSelection.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 替换选中内容（文本和图片）的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

import type { ReplaceSelectionOptions, TextFormat, ImageData } from "./types";

// 重新导出类型供外部使用 / Re-export types for external use
export type { ReplaceSelectionOptions, TextFormat, ImageData };

/**
 * 应用文本格式到范围 / Apply text format to range
 *
 * @param range - Word 范围对象 / Word range object
 * @param format - 文本格式 / Text format
 *
 * @remarks
 * 此函数将指定的文本格式应用到给定的范围
 * This function applies the specified text format to the given range
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
 * 获取选中范围的原始格式 / Get original format of selection
 *
 * @param range - Word 范围对象 / Word range object
 * @returns Promise<TextFormat> 文本格式 / Text format
 *
 * @remarks
 * 此函数获取选中范围的原始文本格式
 * This function gets the original text format of the selection
 */
async function getOriginalFormat(range: Word.Range): Promise<TextFormat> {
  const font = range.font;

  font.load([
    "name",
    "size",
    "bold",
    "italic",
    "underline",
    "color",
    "highlightColor",
    "strikeThrough",
    "superscript",
    "subscript",
  ]);

  return {
    fontName: font.name,
    fontSize: font.size,
    bold: font.bold,
    italic: font.italic,
    underline: font.underline,
    color: font.color,
    highlightColor: font.highlightColor,
    strikeThrough: font.strikeThrough,
    superscript: font.superscript,
    subscript: font.subscript,
  };
}

/**
 * 替换选中内容（文本和图片）
 * Replace selected content (text and images)
 *
 * @param options - 替换选项 / Replace options
 * @returns Promise<void>
 *
 * @remarks
 * 此函数替换当前选中的内容为新的文本和图片。
 * 如果指定了 replaceSelection 为 true（默认），则会替换选中的内容（包括文本和图片）。
 * 如果未指定格式，则使用原选中范围的格式。
 * 图片会按顺序插入到文本之后。
 *
 * This function replaces the currently selected content with new text and images.
 * If replaceSelection is true (default), it will replace the selected content (including text and images).
 * If format is not specified, it will use the original selection format.
 * Images will be inserted in order after the text.
 *
 * @example
 * ```typescript
 * // 替换选中内容为新文本，保持原格式
 * await replaceSelection({
 *   text: "新文本内容"
 * });
 *
 * // 替换选中内容为新文本，并应用新格式
 * await replaceSelection({
 *   text: "新文本内容",
 *   format: {
 *     fontName: "Arial",
 *     fontSize: 14,
 *     bold: true,
 *     color: "#FF0000"
 *   }
 * });
 *
 * // 替换选中内容为文本和图片
 * await replaceSelection({
 *   text: "这是一段文本",
 *   images: [
 *     {
 *       base64: "data:image/png;base64,iVBORw0KGgoAAAANS...",
 *       width: 200,
 *       height: 150,
 *       altText: "示例图片"
 *     }
 *   ]
 * });
 *
 * // 在选中位置后插入内容，不替换
 * await replaceSelection({
 *   text: "插入的文本",
 *   replaceSelection: false
 * });
 * ```
 */
export async function replaceSelection(options: ReplaceSelectionOptions): Promise<void> {
  const { text, format, images, replaceSelection = true } = options;

  // 验证参数 / Validate parameters
  if (!text && (!images || images.length === 0)) {
    throw new Error("必须提供文本或图片 / Must provide text or images");
  }

  try {
    await Word.run(async (context) => {
      // 获取当前选中的范围 / Get current selection range
      const selection = context.document.getSelection();

      // 加载选中范围的基本属性 / Load basic properties of selection
      // eslint-disable-next-line office-addins/no-navigational-load
      selection.load(["text", "isEmpty", "font"]);
      await context.sync();

      // 保存原始格式（如果需要）/ Save original format (if needed)
      let originalFormat: TextFormat | undefined;
      if (!format && !selection.isEmpty) {
        originalFormat = await getOriginalFormat(selection);
        await context.sync();
      }

      // 确定要使用的格式 / Determine format to use
      const formatToApply = format || originalFormat;

      // 处理替换或插入 / Handle replace or insert
      let targetRange: Word.Range;

      if (replaceSelection) {
        // 替换模式：清空选中内容 / Replace mode: clear selection
        if (!selection.isEmpty) {
          selection.clear();
          await context.sync();
        }
        targetRange = selection;
      } else {
        // 插入模式：在选中位置后插入 / Insert mode: insert after selection
        targetRange = selection.getRange("End");
      }

      // 插入文本 / Insert text
      let insertedRange: Word.Range | undefined;
      if (text) {
        insertedRange = targetRange.insertText(text, "Replace");

        // 应用格式 / Apply format
        if (formatToApply) {
          applyTextFormat(insertedRange, formatToApply);
        }

        await context.sync();
      }

      // 插入图片 / Insert images
      if (images && images.length > 0) {
        // 确定插入位置 / Determine insert position
        let imageInsertRange: Word.Range;
        if (insertedRange) {
          // 如果插入了文本，在文本后插入图片 / If text was inserted, insert images after text
          imageInsertRange = insertedRange.getRange("End");
        } else {
          // 如果没有插入文本，直接在目标位置插入 / If no text was inserted, insert at target position
          imageInsertRange = targetRange;
        }

        // 按顺序插入图片 / Insert images in order
        const insertedPictures: Word.InlinePicture[] = [];
        for (const imageData of images) {
          try {
            // 移除 base64 前缀（如果有）/ Remove base64 prefix (if exists)
            let base64Data = imageData.base64;
            if (base64Data.includes(",")) {
              base64Data = base64Data.split(",")[1];
            }

            // 插入图片 / Insert image
            const inlinePicture = imageInsertRange.insertInlinePictureFromBase64(base64Data, "End");

            // 设置图片属性 / Set image properties
            if (imageData.width !== undefined) {
              inlinePicture.width = imageData.width;
            }
            if (imageData.height !== undefined) {
              inlinePicture.height = imageData.height;
            }
            if (imageData.altText !== undefined) {
              inlinePicture.altTextTitle = imageData.altText;
            }

            insertedPictures.push(inlinePicture);

            // 更新插入位置到当前图片之后 / Update insert position to after current image
            imageInsertRange = inlinePicture.getRange("End");
          } catch (error) {
            console.warn(`插入图片失败 / Failed to insert image:`, error);
            // 继续插入下一张图片 / Continue to insert next image
          }
        }

        // 批量同步所有图片插入操作 / Batch sync all image insert operations
        if (insertedPictures.length > 0) {
          await context.sync();
        }
      }

      console.log(
        `成功${replaceSelection ? "替换" : "插入"}内容 / Successfully ${replaceSelection ? "replaced" : "inserted"} content`
      );
    });
  } catch (error) {
    console.error(
      `${replaceSelection ? "替换" : "插入"}内容失败 / Failed to ${replaceSelection ? "replace" : "insert"} content:`,
      error
    );
    throw error;
  }
}

/**
 * 替换选中文本（简化版本）
 * Replace selected text (simplified version)
 *
 * @param text - 文本内容 / Text content
 * @param format - 文本格式（可选）/ Text format (optional)
 * @returns Promise<void>
 *
 * @remarks
 * 这是 replaceSelection 的简化版本，只替换文本
 * This is a simplified version of replaceSelection that only replaces text
 *
 * @example
 * ```typescript
 * // 替换选中内容为新文本
 * await replaceTextAtSelection("新文本内容");
 *
 * // 替换选中内容为新文本，并应用格式
 * await replaceTextAtSelection("新文本内容", {
 *   fontName: "Arial",
 *   fontSize: 14,
 *   bold: true
 * });
 * ```
 */
export async function replaceTextAtSelection(text: string, format?: TextFormat): Promise<void> {
  return replaceSelection({
    text,
    format,
    replaceSelection: true,
  });
}
