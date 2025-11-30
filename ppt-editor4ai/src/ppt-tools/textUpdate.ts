/**
 * 文件名: textUpdate.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 文本框更新工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint, console */

export interface TextUpdateOptions {
  elementId: string; // 要更新的元素ID
  text?: string; // 更新文本内容
  fontSize?: number; // 字号
  fontName?: string; // 字体名称
  fontColor?: string; // 字体颜色（十六进制，如 "#FF0000"）
  bold?: boolean; // 是否加粗
  italic?: boolean; // 是否斜体
  underline?: boolean; // 是否下划线
  horizontalAlignment?: "Left" | "Center" | "Right" | "Justify" | "Distributed"; // 水平对齐
  verticalAlignment?: "Top" | "Middle" | "Bottom"; // 垂直对齐
  backgroundColor?: string; // 背景颜色（十六进制）
  left?: number; // X坐标
  top?: number; // Y坐标
  width?: number; // 宽度
  height?: number; // 高度
}

export interface TextUpdateResult {
  success: boolean;
  message: string;
  elementId?: string;
}

/**
 * 更新文本框
 * @param options 更新选项
 * @returns Promise<TextUpdateResult>
 */
export async function updateTextBox(options: TextUpdateOptions): Promise<TextUpdateResult> {
  const { elementId } = options;

  if (!elementId) {
    return {
      success: false,
      message: "元素ID不能为空",
    };
  }

  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // eslint-disable-next-line office-addins/no-navigational-load
      slide.load("shapes");
      await context.sync();

      // 直接通过 shapes 添加/操作，不读取 items
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      // 查找目标元素 - 合并为一次循环
      let targetShape: PowerPoint.Shape | null = null;
      for (const shape of shapes.items) {
        shape.load("id,type");
      }
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.id === elementId) {
          targetShape = shape;
          break;
        }
      }

      if (!targetShape) {
        throw new Error(`未找到ID为 ${elementId} 的元素`);
      }

      // 验证元素类型是否支持文本
      const supportedTypes = ["TextBox", "Placeholder", "GeometricShape"];
      if (!supportedTypes.includes(targetShape.type)) {
        throw new Error(`元素类型 ${targetShape.type} 不支持文本编辑`);
      }

      // 更新位置和尺寸
      if (options.left !== undefined) {
        targetShape.left = options.left;
      }
      if (options.top !== undefined) {
        targetShape.top = options.top;
      }
      if (options.width !== undefined) {
        targetShape.width = options.width;
      }
      if (options.height !== undefined) {
        targetShape.height = options.height;
      }

      // 更新背景颜色
      if (options.backgroundColor !== undefined) {
        targetShape.fill.setSolidColor(options.backgroundColor);
      }

      // 获取文本框
      const textFrame = targetShape.textFrame;
      textFrame.load("textRange");
      await context.sync();

      const textRange = textFrame.textRange;

      // 更新文本内容
      if (options.text !== undefined) {
        textRange.text = options.text;
      }

      // 加载字体对象
      const font = textRange.font;
      font.load("name,size,color,bold,italic,underline");
      await context.sync();

      // 更新字体属性
      if (options.fontSize !== undefined) {
        font.size = options.fontSize;
      }
      if (options.fontName !== undefined) {
        font.name = options.fontName;
      }
      if (options.fontColor !== undefined) {
        font.color = options.fontColor;
      }
      if (options.bold !== undefined) {
        font.bold = options.bold;
      }
      if (options.italic !== undefined) {
        font.italic = options.italic;
      }
      if (options.underline !== undefined) {
        font.underline = options.underline ? "Single" : "None";
      }

      // 更新段落对齐
      if (options.horizontalAlignment !== undefined) {
        const paragraphFormat = textRange.paragraphFormat;
        paragraphFormat.load("horizontalAlignment");
        await context.sync();
        paragraphFormat.horizontalAlignment = options.horizontalAlignment;
      }

      // 更新垂直对齐
      if (options.verticalAlignment !== undefined) {
        textFrame.verticalAlignment = options.verticalAlignment;
      }

      await context.sync();
    });

    return {
      success: true,
      message: "文本框更新成功",
      elementId,
    };
  } catch (error) {
    console.error("更新文本框失败:", error);
    return {
      success: false,
      message: error instanceof Error ? error.message : "未知错误",
      elementId,
    };
  }
}

/**
 * 批量更新多个文本框
 * @param updates 更新选项数组
 * @returns Promise<TextUpdateResult[]>
 */
export async function updateTextBoxes(updates: TextUpdateOptions[]): Promise<TextUpdateResult[]> {
  const results: TextUpdateResult[] = [];

  for (const update of updates) {
    const result = await updateTextBox(update);
    results.push(result);
  }

  return results;
}

/**
 * 获取文本框的当前样式信息
 * @param elementId 元素ID
 * @returns Promise<TextUpdateOptions | null>
 */
export async function getTextBoxStyle(elementId: string): Promise<TextUpdateOptions | null> {
  if (!elementId) {
    return null;
  }

  try {
    let styleInfo: TextUpdateOptions | null = null;

    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // eslint-disable-next-line office-addins/no-navigational-load
      slide.load("shapes");
      await context.sync();

      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      // 查找目标元素
      let targetShape: PowerPoint.Shape | null = null;
      for (const shape of shapes.items) {
        shape.load("id,type,left,top,width,height");
      }
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.id === elementId) {
          targetShape = shape;
          break;
        }
      }

      if (!targetShape) {
        throw new Error(`未找到ID为 ${elementId} 的元素`);
      }

      // 获取文本框
      const textFrame = targetShape.textFrame;
      textFrame.load("textRange,verticalAlignment");
      await context.sync();

      const textRange = textFrame.textRange;
      textRange.load("text");

      // 加载字体对象
      const font = textRange.font;
      font.load("name,size,color,bold,italic,underline");

      // 加载段落对齐
      const paragraphFormat = textRange.paragraphFormat;
      paragraphFormat.load("horizontalAlignment");
      await context.sync();

      const horizontalAlignment = paragraphFormat.horizontalAlignment as
        | "Left"
        | "Center"
        | "Right"
        | "Justify"
        | "Distributed";

      // 加载背景颜色
      targetShape.fill.load("type,foregroundColor");
      await context.sync();

      let backgroundColor: string | undefined;
      if (targetShape.fill.type === "Solid") {
        backgroundColor = targetShape.fill.foregroundColor;
      }

      styleInfo = {
        elementId,
        text: textRange.text,
        fontSize: font.size,
        fontName: font.name,
        fontColor: font.color,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline !== "None",
        horizontalAlignment,
        verticalAlignment: textFrame.verticalAlignment as "Top" | "Middle" | "Bottom",
        backgroundColor,
        left: targetShape.left,
        top: targetShape.top,
        width: targetShape.width,
        height: targetShape.height,
      };
    });

    return styleInfo;
  } catch (error) {
    console.error("获取文本框样式失败:", error);
    return null;
  }
}
