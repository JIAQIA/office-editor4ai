/**
 * 文件名: elementsList.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 元素列表获取工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint */

export interface SlideElement {
  id: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  name?: string;
  text?: string;
}

export interface GetElementsOptions {
  slideIndex?: number; // 幻灯片索引，默认为当前选中的第一张
  includeText?: boolean; // 是否包含文本内容，默认为 true
}

/**
 * 获取指定幻灯片的所有元素
 * @param options 获取选项
 * @returns Promise<SlideElement[]> 元素列表
 */
export async function getSlideElements(options: GetElementsOptions = {}): Promise<SlideElement[]> {
  const { slideIndex = 0, includeText = true } = options;

  try {
    const elementsList: SlideElement[] = [];

    await PowerPoint.run(async (context) => {
      // 获取幻灯片
      const slides = context.presentation.slides;
      const selectedSlide = slides.getItemAt(slideIndex);

      // 加载幻灯片的形状集合
      const shapes = selectedSlide.shapes;
      shapes.load("items");

      await context.sync();

      // 收集所有元素信息
      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];

        // 加载形状的基本属性
        shape.load("id,type,left,top,width,height,name");
        await context.sync();

        // 尝试获取文本内容
        let textContent: string | undefined;
        if (includeText) {
          try {
            const textFrame = shape.textFrame;
            textFrame.load("textRange");
            await context.sync();

            const textRange = textFrame.textRange;
            textRange.load("text");
            await context.sync();

            textContent = textRange.text?.trim() || undefined;
          } catch (e) {
            // 如果形状没有文本框，忽略错误
            textContent = undefined;
          }
        }

        elementsList.push({
          id: shape.id,
          type: shape.type,
          left: Math.round(shape.left * 100) / 100,
          top: Math.round(shape.top * 100) / 100,
          width: Math.round(shape.width * 100) / 100,
          height: Math.round(shape.height * 100) / 100,
          name: shape.name || undefined,
          text: textContent,
        });
      }
    });

    return elementsList;
  } catch (error) {
    console.error("获取元素列表失败:", error);
    throw error;
  }
}

/**
 * 获取当前选中幻灯片的所有元素（简化版本）
 * @returns Promise<SlideElement[]> 元素列表
 */
export async function getCurrentSlideElements(): Promise<SlideElement[]> {
  return getSlideElements({ slideIndex: 0, includeText: true });
}
