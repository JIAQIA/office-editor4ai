/**
 * 文件名: slideMove.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片移动工具核心逻辑，与 Office API 交互，支持修改幻灯片页码/排序
 */

/* global PowerPoint, console */

export interface SlideMoveOptions {
  fromIndex: number; // 源位置索引（从1开始）
  toIndex: number; // 目标位置索引（从1开始）
}

export interface SlideMoveResult {
  success: boolean;
  message: string;
  fromIndex?: number;
  toIndex?: number;
  totalSlides?: number;
}

export interface SlideInfo {
  index: number; // 幻灯片索引（从1开始）
  id: string; // 幻灯片ID
  title?: string; // 幻灯片标题（如果有）
}

/**
 * 移动幻灯片到新位置
 * @param options 移动选项
 * @returns Promise<SlideMoveResult>
 */
export async function moveSlide(options: SlideMoveOptions): Promise<SlideMoveResult> {
  const { fromIndex, toIndex } = options;

  // 验证输入
  if (!Number.isInteger(fromIndex) || fromIndex < 1) {
    return {
      success: false,
      message: "源位置索引必须是大于0的整数",
    };
  }

  if (!Number.isInteger(toIndex) || toIndex < 1) {
    return {
      success: false,
      message: "目标位置索引必须是大于0的整数",
    };
  }

  if (fromIndex === toIndex) {
    return {
      success: false,
      message: "源位置和目标位置相同，无需移动",
    };
  }

  try {
    let totalSlides = 0;

    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      slides.load("items");
      await context.sync();

      totalSlides = slides.items.length;

      // 验证索引范围
      if (fromIndex > totalSlides) {
        throw new Error(`源位置索引 ${fromIndex} 超出范围，当前共有 ${totalSlides} 张幻灯片`);
      }

      if (toIndex > totalSlides) {
        throw new Error(`目标位置索引 ${toIndex} 超出范围，当前共有 ${totalSlides} 张幻灯片`);
      }

      // 获取要移动的幻灯片（索引从0开始，所以减1）
      const slideToMove = slides.items[fromIndex - 1];

      // PowerPoint API 的 moveTo 方法：将幻灯片移动到指定位置
      // 注意：moveTo 的索引参数是从0开始的
      slideToMove.moveTo(toIndex - 1);

      await context.sync();
    });

    return {
      success: true,
      message: `成功将幻灯片从位置 ${fromIndex} 移动到位置 ${toIndex}`,
      fromIndex,
      toIndex,
      totalSlides,
    };
  } catch (error) {
    console.error("移动幻灯片失败:", error);
    return {
      success: false,
      message: error instanceof Error ? error.message : "未知错误",
      fromIndex,
      toIndex,
    };
  }
}

/**
 * 将当前选中的幻灯片移动到指定位置
 * @param toIndex 目标位置索引（从1开始）
 * @returns Promise<SlideMoveResult>
 */
export async function moveCurrentSlide(toIndex: number): Promise<SlideMoveResult> {
  if (!Number.isInteger(toIndex) || toIndex < 1) {
    return {
      success: false,
      message: "目标位置索引必须是大于0的整数",
    };
  }

  try {
    let fromIndex = 0;
    let totalSlides = 0;

    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      slides.load("items");

      // 获取当前选中的幻灯片
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");
      await context.sync();

      totalSlides = slides.items.length;

      if (selectedSlides.items.length === 0) {
        throw new Error("未选中任何幻灯片");
      }

      if (selectedSlides.items.length > 1) {
        throw new Error("请只选中一张幻灯片");
      }

      const slideToMove = selectedSlides.items[0];
      slideToMove.load("id");
      await context.sync();

      // 查找当前幻灯片的索引
      for (let i = 0; i < slides.items.length; i++) {
        slides.items[i].load("id");
      }
      await context.sync();

      for (let i = 0; i < slides.items.length; i++) {
        if (slides.items[i].id === slideToMove.id) {
          fromIndex = i + 1; // 转换为从1开始的索引
          break;
        }
      }

      if (fromIndex === 0) {
        throw new Error("无法找到选中幻灯片的位置");
      }

      if (fromIndex === toIndex) {
        throw new Error("幻灯片已在目标位置，无需移动");
      }

      if (toIndex > totalSlides) {
        throw new Error(`目标位置索引 ${toIndex} 超出范围，当前共有 ${totalSlides} 张幻灯片`);
      }

      // 移动幻灯片
      slideToMove.moveTo(toIndex - 1);
      await context.sync();
    });

    return {
      success: true,
      message: `成功将当前幻灯片从位置 ${fromIndex} 移动到位置 ${toIndex}`,
      fromIndex,
      toIndex,
      totalSlides,
    };
  } catch (error) {
    console.error("移动当前幻灯片失败:", error);
    return {
      success: false,
      message: error instanceof Error ? error.message : "未知错误",
      toIndex,
    };
  }
}

/**
 * 批量移动多张幻灯片
 * @param moves 移动操作数组
 * @returns Promise<SlideMoveResult[]>
 */
export async function moveSlides(moves: SlideMoveOptions[]): Promise<SlideMoveResult[]> {
  const results: SlideMoveResult[] = [];

  // 注意：批量移动时需要考虑顺序，因为每次移动都会改变其他幻灯片的索引
  // 这里采用逐个移动的策略，每次移动后重新计算索引
  for (const move of moves) {
    const result = await moveSlide(move);
    results.push(result);

    // 如果某次移动失败，停止后续操作
    if (!result.success) {
      break;
    }
  }

  return results;
}

/**
 * 获取所有幻灯片的基本信息
 * @returns Promise<SlideInfo[]>
 */
export async function getAllSlidesInfo(): Promise<SlideInfo[]> {
  try {
    const slidesInfo: SlideInfo[] = [];

    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      slides.load("items");
      await context.sync();

      // 批量加载所有幻灯片的基本信息
      for (let i = 0; i < slides.items.length; i++) {
        const slide = slides.items[i];
        slide.load("id,shapes");
        const shapes = slide.shapes;
        shapes.load("items");
      }
      await context.sync();

      // 批量加载所有 shape 的属性
      for (let i = 0; i < slides.items.length; i++) {
        const shapes = slides.items[i].shapes;
        for (const shape of shapes.items) {
          shape.load("type,name,textFrame");
        }
      }
      await context.sync();

      // 批量加载所有文本内容
      for (let i = 0; i < slides.items.length; i++) {
        const shapes = slides.items[i].shapes;
        for (const shape of shapes.items) {
          // 查找标题占位符或第一个文本框
          if (shape.type === "Placeholder" || shape.type === "TextBox") {
            try {
              const textFrame = shape.textFrame;
              textFrame.load("textRange");
              const textRange = textFrame.textRange;
              textRange.load("text");
              // eslint-disable-next-line @typescript-eslint/no-unused-vars
            } catch (_e) {
              // 忽略无法读取文本的形状
              continue;
            }
          }
        }
      }
      await context.sync();

      // 处理数据，提取标题
      for (let i = 0; i < slides.items.length; i++) {
        const slide = slides.items[i];
        let title: string | undefined;
        const shapes = slide.shapes;

        for (const shape of shapes.items) {
          if (shape.type === "Placeholder" || shape.type === "TextBox") {
            try {
              const textRange = shape.textFrame.textRange;
              if (textRange.text && textRange.text.trim()) {
                title = textRange.text.trim();
                break;
              }
              // eslint-disable-next-line @typescript-eslint/no-unused-vars
            } catch (_e) {
              // 忽略无法读取文本的形状
              continue;
            }
          }
        }

        slidesInfo.push({
          index: i + 1,
          id: slide.id,
          title,
        });
      }
    });

    return slidesInfo;
  } catch (error) {
    console.error("获取幻灯片信息失败:", error);
    return [];
  }
}

/**
 * 交换两张幻灯片的位置
 * @param index1 第一张幻灯片的索引（从1开始）
 * @param index2 第二张幻灯片的索引（从1开始）
 * @returns Promise<SlideMoveResult>
 */
export async function swapSlides(index1: number, index2: number): Promise<SlideMoveResult> {
  if (!Number.isInteger(index1) || index1 < 1) {
    return {
      success: false,
      message: "第一张幻灯片索引必须是大于0的整数",
    };
  }

  if (!Number.isInteger(index2) || index2 < 1) {
    return {
      success: false,
      message: "第二张幻灯片索引必须是大于0的整数",
    };
  }

  if (index1 === index2) {
    return {
      success: false,
      message: "两张幻灯片索引相同，无需交换",
    };
  }

  try {
    // 交换策略：先将第一张移到末尾，再将第二张移到第一张的位置，最后将第一张移到第二张的位置
    let totalSlides = 0;

    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const slides = presentation.slides;
      slides.load("items");
      await context.sync();

      totalSlides = slides.items.length;

      if (index1 > totalSlides || index2 > totalSlides) {
        throw new Error(`索引超出范围，当前共有 ${totalSlides} 张幻灯片`);
      }

      // 确定较小和较大的索引
      const smallerIndex = Math.min(index1, index2);
      const largerIndex = Math.max(index1, index2);

      // 先移动较大索引的幻灯片到较小索引位置
      const slide2 = slides.items[largerIndex - 1];
      slide2.moveTo(smallerIndex - 1);
      await context.sync();

      // 再移动原来较小索引的幻灯片到较大索引位置
      // 注意：由于第一次移动，原来较小索引的幻灯片现在在 smallerIndex 位置
      const slide1 = slides.items[smallerIndex];
      slide1.moveTo(largerIndex - 1);
      await context.sync();
    });

    return {
      success: true,
      message: `成功交换位置 ${index1} 和位置 ${index2} 的幻灯片`,
      fromIndex: index1,
      toIndex: index2,
      totalSlides,
    };
  } catch (error) {
    console.error("交换幻灯片失败:", error);
    return {
      success: false,
      message: error instanceof Error ? error.message : "未知错误",
      fromIndex: index1,
      toIndex: index2,
    };
  }
}
