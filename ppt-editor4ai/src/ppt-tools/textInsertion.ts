/**
 * 文件名: textInsertion.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 文本插入工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint */

export interface TextInsertionOptions {
  text: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  fillColor?: string;
  lineColor?: string;
  lineWeight?: number;
}

/**
 * 插入文本框到幻灯片
 * @param options 文本插入选项
 * @returns Promise<void>
 */
export async function insertTextToSlide(options: TextInsertionOptions): Promise<void> {
  const {
    text,
    left,
    top,
    width = 300,
    height = 100,
    fillColor = "white",
    lineColor = "black",
    lineWeight = 1,
  } = options;

  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);

      // 如果指定了位置参数，使用指定位置；否则使用默认位置
      let textBox;
      if (left !== undefined && top !== undefined) {
        textBox = slide.shapes.addTextBox(text, {
          left,
          top,
          width,
          height,
        });
      } else {
        textBox = slide.shapes.addTextBox(text);
      }

      // 设置样式
      textBox.fill.setSolidColor(fillColor);
      textBox.lineFormat.color = lineColor;
      textBox.lineFormat.weight = lineWeight;
      textBox.lineFormat.dashStyle = "Solid";
      
      await context.sync();
    });
  } catch (error) {
    console.error("插入文本失败:", error);
    throw error;
  }
}

/**
 * 简化版本：插入文本框（兼容旧接口）
 * @param text 文本内容
 * @param left X坐标（可选）
 * @param top Y坐标（可选）
 */
export async function insertText(text: string, left?: number, top?: number): Promise<void> {
  return insertTextToSlide({ text, left, top });
}
