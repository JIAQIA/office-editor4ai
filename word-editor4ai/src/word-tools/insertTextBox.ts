/**
 * 文件名: insertTextBox.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 插入文本框工具核心逻辑
 */

/* global Word, console */

import type { InsertLocation, TextFormat } from "./types";

// 重新导出以保持向后兼容性 / Re-export for backward compatibility
export type { InsertLocation };

/**
 * 文本框选项 / Text Box Options
 */
export interface TextBoxOptions {
  /** 文本框宽度（磅），默认为 150 / Text box width in points, default 150 */
  width?: number;
  /** 文本框高度（磅），默认为 100 / Text box height in points, default 100 */
  height?: number;
  /** 文本框名称 / Text box name */
  name?: string;
  /** 文本格式 / Text format */
  format?: TextFormat;
  /** 是否锁定纵横比，默认为 false / Lock aspect ratio, default false */
  lockAspectRatio?: boolean;
  /** 是否可见，默认为 true / Visible, default true */
  visible?: boolean;
  /** 左边距（磅）/ Left position in points */
  left?: number;
  /** 上边距（磅）/ Top position in points */
  top?: number;
  /** 旋转角度（度）/ Rotation in degrees */
  rotation?: number;
}

/**
 * 插入文本框结果 / Insert Text Box Result
 */
export interface InsertTextBoxResult {
  /** 是否成功 / Success */
  success: boolean;
  /** 文本框标识符 / Text box identifier */
  textBoxId?: string;
  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;
}

/**
 * 在文档中插入文本框
 * Insert text box in document
 *
 * @param text - 文本框内容 / Text box content
 * @param location - 插入位置 / Insert location
 * @param options - 文本框选项 / Text box options
 * @returns Promise<InsertTextBoxResult> 插入结果 / Insert result
 *
 * @remarks
 * 注意：Word JavaScript API 对文本框的支持有限
 * - 文本框通过 Shape 对象创建
 * - 插入位置基于当前选择或文档范围
 * - 某些高级属性（如精确定位）可能需要 OOXML
 *
 * Note: Word JavaScript API has limited support for text boxes
 * - Text boxes are created through Shape objects
 * - Insert location is based on current selection or document range
 * - Some advanced properties (like precise positioning) may require OOXML
 *
 * @example
 * ```typescript
 * // 插入简单文本框
 * await insertTextBox("Hello World", "End");
 *
 * // 插入带格式的文本框
 * await insertTextBox("Formatted Text", "End", {
 *   width: 200,
 *   height: 150,
 *   name: "MyTextBox",
 *   format: {
 *     fontName: "Arial",
 *     fontSize: 14,
 *     bold: true,
 *     color: "#FF0000"
 *   }
 * });
 * ```
 */
export async function insertTextBox(
  text: string,
  location: InsertLocation = "End",
  options: TextBoxOptions = {}
): Promise<InsertTextBoxResult> {
  const {
    width = 150,
    height = 100,
    name,
    format,
    lockAspectRatio = false,
    visible = true,
    left,
    top,
    rotation,
  } = options;

  // 验证参数 / Validate parameters
  if (!text) {
    return {
      success: false,
      error: "必须提供文本内容 / Text content is required",
    };
  }

  try {
    let textBoxId: string | undefined;

    await Word.run(async (context) => {
      // 获取插入范围 / Get insert range
      let insertRange: Word.Range;
      const selection = context.document.getSelection();

      switch (location) {
        case "Start":
          insertRange = context.document.body.getRange("Start");
          break;
        case "End":
          insertRange = context.document.body.getRange("End");
          break;
        case "Before":
          insertRange = selection;
          break;
        case "After":
          insertRange = selection;
          break;
        case "Replace":
          insertRange = selection;
          break;
        default:
          insertRange = context.document.body.getRange("End");
      }

      // 在范围处插入文本框
      // Insert text box at range
      // 注意：Word JavaScript API 通过 insertTextBox 创建文本框，返回 Word.Shape 对象
      // Note: Word JavaScript API creates text boxes through insertTextBox, returns Word.Shape object
      const insertShapeOptions: Word.InsertShapeOptions = {
        width,
        height,
      };

      // 添加位置参数（如果提供）/ Add position parameters (if provided)
      if (left !== undefined) {
        insertShapeOptions.left = left;
      }
      if (top !== undefined) {
        insertShapeOptions.top = top;
      }

      const textBox = insertRange.insertTextBox(text, insertShapeOptions);

      // 设置文本框属性 / Set text box properties
      if (name) {
        textBox.name = name;
      }
      if (lockAspectRatio !== undefined) {
        textBox.lockAspectRatio = lockAspectRatio;
      }
      if (visible !== undefined) {
        textBox.visible = visible;
      }
      if (rotation !== undefined) {
        textBox.rotation = rotation;
      }

      // 应用文本格式 / Apply text format
      if (format) {
        try {
          // 获取文本框的文本范围 / Get text box text range
          // Word.Shape 对象的 body 属性包含文本内容（仅适用于文本框和几何形状）
          // Word.Shape object's body property contains text content (only applies to text boxes and geometric shapes)
          const textBoxBody = textBox.body;
          const textRange = textBoxBody.getRange("Whole");
          const font = textRange.font;

          if (format.fontName) {
            font.name = format.fontName;
          }
          if (format.fontSize) {
            font.size = format.fontSize;
          }
          if (format.bold !== undefined) {
            font.bold = format.bold;
          }
          if (format.italic !== undefined) {
            font.italic = format.italic;
          }
          if (format.underline) {
            font.underline = format.underline as Word.UnderlineType;
          }
          if (format.color) {
            font.color = format.color;
          }
          if (format.highlightColor) {
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
        } catch (error) {
          console.warn("应用文本格式时出错 / Error applying text format:", error);
        }
      }

      // 加载文本框 ID / Load text box ID
      textBox.load("id");
      await context.sync();

      textBoxId = `textbox-${textBox.id}`;
    });

    return {
      success: true,
      textBoxId,
    };
  } catch (error) {
    console.error("插入文本框失败 / Insert text box failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 批量插入文本框 / Batch Insert Text Boxes
 *
 * @param textBoxes - 文本框列表 / Text box list
 * @returns Promise<InsertTextBoxResult[]> 插入结果列表 / Insert result list
 */
export async function insertTextBoxes(
  textBoxes: Array<{
    text: string;
    location: InsertLocation;
    options?: TextBoxOptions;
  }>
): Promise<InsertTextBoxResult[]> {
  const results: InsertTextBoxResult[] = [];

  for (const textBoxData of textBoxes) {
    const result = await insertTextBox(textBoxData.text, textBoxData.location, textBoxData.options);
    results.push(result);
  }

  return results;
}
