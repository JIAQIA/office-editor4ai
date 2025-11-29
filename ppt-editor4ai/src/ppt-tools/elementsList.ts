/**
 * 文件名: elementsList.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 元素列表获取工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint, console */

export interface SlideElement {
  id: string;
  type: string; // Shape 的主类型，如 "Image", "TextBox", "Placeholder", "GeometricShape" 等
  left: number;
  top: number;
  width: number;
  height: number;
  name?: string;
  text?: string;
  placeholderType?: string; // 当 type === "Placeholder" 时，表示占位符的具体类型，如 "Title", "Body", "Picture" 等
  placeholderContainedType?: string; // 当 type === "Placeholder" 时，表示占位符内包含的内容类型
}

export interface GetElementsOptions {
  slideNumber?: number; // 幻灯片页码（从1开始），不填则使用当前页
  includeText?: boolean; // 是否包含文本内容，默认为 true
}

/**
 * 获取指定幻灯片的所有元素
 * @param options 获取选项
 * @returns Promise<SlideElement[]> 元素列表，如果页码不存在则返回空数组
 */
export async function getSlideElements(options: GetElementsOptions = {}): Promise<SlideElement[]> {
  const { slideNumber, includeText = true } = options;

  try {
    const elementsList: SlideElement[] = [];

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // 确定要获取的幻灯片
      let targetSlide: PowerPoint.Slide;

      if (slideNumber !== undefined) {
        // 使用指定的页码（从1开始）
        const slideIndex = slideNumber - 1;

        // 验证页码是否存在
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          console.warn(`页码 ${slideNumber} 不存在，总共有 ${slides.items.length} 页`);
          return; // 返回空数组
        }

        targetSlide = slides.items[slideIndex];
      } else {
        // 使用当前选中的幻灯片
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length === 0) {
          console.warn("没有选中的幻灯片");
          return; // 返回空数组
        }

        targetSlide = selectedSlides.items[0];
      }

      // 加载幻灯片的形状集合
      const shapes = targetSlide.shapes;
      shapes.load("items");

      await context.sync();

      // 批量加载所有 shape 的基本属性
      for (const shape of shapes.items) {
        shape.load("id,type,left,top,width,height,name");
      }

      try {
        await context.sync();
      } catch (error) {
        // 如果批量加载失败，可能是因为某些 shape 不支持这些属性
        console.warn("批量加载失败，尝试逐个加载:", error);

        // 清空并重新逐个加载
        for (const shape of shapes.items) {
          shape.load("id,type,left,top,width,height,name");
        }
        // 批量同步所有加载请求
        try {
          await context.sync();
        } catch (e) {
          console.warn("部分元素无法加载基本属性:", e);
        }
      }

      // 第一步：批量加载所有需要的额外属性
      for (const shape of shapes.items) {
        const supportsText = shape.type !== "Media" && shape.type !== "Line";
        if (includeText && supportsText) {
          try {
            shape.textFrame.load("hasText,textRange");
          } catch {
            // 忽略不支持文本框的形状
          }
        }
        if (shape.type === "Placeholder") {
          try {
            shape.load("placeholderFormat");
          } catch {
            // 忽略加载失败
          }
        }
      }

      // 统一同步所有额外属性
      try {
        await context.sync();
      } catch {
        // 忽略部分属性加载失败
      }

      // 第二步：收集所有元素信息（不再需要 sync）
      for (const shape of shapes.items) {
        try {
          // 获取文本内容
          let textContent: string | undefined;
          const supportsText = shape.type !== "Media" && shape.type !== "Line";

          if (includeText && supportsText) {
            try {
              const textFrame = shape.textFrame;
              if (textFrame.hasText && textFrame.textRange) {
                textContent = textFrame.textRange.text?.trim() || undefined;
              }
            } catch {
              // 如果形状不支持文本框，忽略错误
              textContent = undefined;
            }
          }

          // 获取 Placeholder 的详细类型信息
          let placeholderType: string | undefined;
          let placeholderContainedType: string | undefined;

          if (shape.type === "Placeholder") {
            try {
              placeholderType = shape.placeholderFormat.type;
              placeholderContainedType = shape.placeholderFormat.containedType || undefined;
            } catch {
              // 如果加载 placeholderFormat 失败，忽略错误
              placeholderType = undefined;
              placeholderContainedType = undefined;
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
            placeholderType,
            placeholderContainedType,
          });
        } catch (error) {
          // 如果某个 shape 处理失败，记录错误但继续处理其他 shape
          console.warn(`跳过无法处理的元素:`, error);
        }
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
  return getSlideElements({ includeText: true });
}

/**
 * 获取指定页码幻灯片的所有元素
 * @param slideNumber 页码（从1开始）
 * @param includeText 是否包含文本内容
 * @returns Promise<SlideElement[]> 元素列表，如果页码不存在则返回空数组
 */
export async function getSlideElementsByPageNumber(
  slideNumber: number,
  includeText: boolean = true
): Promise<SlideElement[]> {
  return getSlideElements({ slideNumber, includeText });
}
