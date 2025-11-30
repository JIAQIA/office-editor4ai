/**
 * 文件名: elementDeletion.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 元素删除工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint, console */

export interface DeleteElementOptions {
  slideNumber?: number; // 幻灯片页码（从1开始），不填则使用当前页
  elementId?: string; // 元素ID
  elementName?: string; // 元素名称（如果没有ID）
  elementIndex?: number; // 元素索引（从0开始，如果没有ID和名称）
}

export interface DeleteElementResult {
  success: boolean;
  deletedCount: number;
  message?: string;
}

/**
 * 通过ID删除元素
 * @param elementId 元素ID
 * @param slideNumber 幻灯片页码（从1开始），不填则使用当前页
 * @returns Promise<DeleteElementResult> 删除结果
 */
export async function deleteElementById(
  elementId: string,
  slideNumber?: number
): Promise<DeleteElementResult> {
  try {
    let deletedCount = 0;

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // 确定要操作的幻灯片
      let targetSlide: PowerPoint.Slide;

      if (slideNumber !== undefined) {
        const slideIndex = slideNumber - 1;
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`页码 ${slideNumber} 不存在，总共有 ${slides.items.length} 页`);
        }
        targetSlide = slides.items[slideIndex];
      } else {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length === 0) {
          throw new Error("没有选中的幻灯片");
        }
        targetSlide = selectedSlides.items[0];
      }

      // 加载所有形状
      const shapes = targetSlide.shapes;
      shapes.load("items");
      await context.sync();

      // 加载所有形状的ID
      for (const shape of shapes.items) {
        shape.load("id");
      }
      await context.sync();

      // 查找并删除匹配的形状
      for (const shape of shapes.items) {
        if (shape.id === elementId) {
          shape.delete();
          deletedCount++;
          break; // ID应该是唯一的，找到后就退出
        }
      }

      await context.sync();
    });

    if (deletedCount === 0) {
      return {
        success: false,
        deletedCount: 0,
        message: `未找到ID为 ${elementId} 的元素`,
      };
    }

    return {
      success: true,
      deletedCount,
      message: `成功删除 ${deletedCount} 个元素`,
    };
  } catch (error) {
    console.error("删除元素失败:", error);
    return {
      success: false,
      deletedCount: 0,
      message: error instanceof Error ? error.message : "删除失败",
    };
  }
}

/**
 * 通过名称删除元素（可能删除多个同名元素）
 * @param elementName 元素名称
 * @param slideNumber 幻灯片页码（从1开始），不填则使用当前页
 * @param deleteAll 是否删除所有同名元素，默认为false（只删除第一个）
 * @returns Promise<DeleteElementResult> 删除结果
 */
export async function deleteElementByName(
  elementName: string,
  slideNumber?: number,
  deleteAll: boolean = false
): Promise<DeleteElementResult> {
  try {
    let deletedCount = 0;

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // 确定要操作的幻灯片
      let targetSlide: PowerPoint.Slide;

      if (slideNumber !== undefined) {
        const slideIndex = slideNumber - 1;
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`页码 ${slideNumber} 不存在，总共有 ${slides.items.length} 页`);
        }
        targetSlide = slides.items[slideIndex];
      } else {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length === 0) {
          throw new Error("没有选中的幻灯片");
        }
        targetSlide = selectedSlides.items[0];
      }

      // 加载所有形状
      const shapes = targetSlide.shapes;
      shapes.load("items");
      await context.sync();

      // 加载所有形状的名称
      for (const shape of shapes.items) {
        shape.load("name");
      }
      await context.sync();

      // 查找并删除匹配的形状
      for (const shape of shapes.items) {
        if (shape.name === elementName) {
          shape.delete();
          deletedCount++;
          if (!deleteAll) {
            break; // 只删除第一个
          }
        }
      }

      await context.sync();
    });

    if (deletedCount === 0) {
      return {
        success: false,
        deletedCount: 0,
        message: `未找到名称为 ${elementName} 的元素`,
      };
    }

    return {
      success: true,
      deletedCount,
      message: `成功删除 ${deletedCount} 个元素`,
    };
  } catch (error) {
    console.error("删除元素失败:", error);
    return {
      success: false,
      deletedCount: 0,
      message: error instanceof Error ? error.message : "删除失败",
    };
  }
}

/**
 * 通过索引删除元素
 * @param elementIndex 元素索引（从0开始）
 * @param slideNumber 幻灯片页码（从1开始），不填则使用当前页
 * @returns Promise<DeleteElementResult> 删除结果
 */
export async function deleteElementByIndex(
  elementIndex: number,
  slideNumber?: number
): Promise<DeleteElementResult> {
  try {
    let deletedCount = 0;

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // 确定要操作的幻灯片
      let targetSlide: PowerPoint.Slide;

      if (slideNumber !== undefined) {
        const slideIndex = slideNumber - 1;
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`页码 ${slideNumber} 不存在，总共有 ${slides.items.length} 页`);
        }
        targetSlide = slides.items[slideIndex];
      } else {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length === 0) {
          throw new Error("没有选中的幻灯片");
        }
        targetSlide = selectedSlides.items[0];
      }

      // 加载所有形状
      const shapes = targetSlide.shapes;
      shapes.load("items");
      await context.sync();

      // 检查索引是否有效
      if (elementIndex < 0 || elementIndex >= shapes.items.length) {
        throw new Error(`索引 ${elementIndex} 超出范围，总共有 ${shapes.items.length} 个元素`);
      }

      // 删除指定索引的形状
      shapes.items[elementIndex].delete();
      deletedCount = 1;

      await context.sync();
    });

    return {
      success: true,
      deletedCount,
      message: `成功删除索引为 ${elementIndex} 的元素`,
    };
  } catch (error) {
    console.error("删除元素失败:", error);
    return {
      success: false,
      deletedCount: 0,
      message: error instanceof Error ? error.message : "删除失败",
    };
  }
}

/**
 * 通用删除元素方法（支持多种选择方式）
 * @param options 删除选项
 * @returns Promise<DeleteElementResult> 删除结果
 */
export async function deleteElement(options: DeleteElementOptions): Promise<DeleteElementResult> {
  const { slideNumber, elementId, elementName, elementIndex } = options;

  // 优先级：ID > 名称 > 索引
  if (elementId) {
    return deleteElementById(elementId, slideNumber);
  } else if (elementName) {
    return deleteElementByName(elementName, slideNumber);
  } else if (elementIndex !== undefined) {
    return deleteElementByIndex(elementIndex, slideNumber);
  } else {
    return {
      success: false,
      deletedCount: 0,
      message: "必须提供 elementId、elementName 或 elementIndex 中的至少一个",
    };
  }
}

/**
 * 批量删除元素
 * @param elementIds 元素ID数组
 * @param slideNumber 幻灯片页码（从1开始），不填则使用当前页
 * @returns Promise<DeleteElementResult> 删除结果
 */
export async function deleteElementsByIds(
  elementIds: string[],
  slideNumber?: number
): Promise<DeleteElementResult> {
  try {
    let deletedCount = 0;

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // 确定要操作的幻灯片
      let targetSlide: PowerPoint.Slide;

      if (slideNumber !== undefined) {
        const slideIndex = slideNumber - 1;
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`页码 ${slideNumber} 不存在，总共有 ${slides.items.length} 页`);
        }
        targetSlide = slides.items[slideIndex];
      } else {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length === 0) {
          throw new Error("没有选中的幻灯片");
        }
        targetSlide = selectedSlides.items[0];
      }

      // 加载所有形状
      const shapes = targetSlide.shapes;
      shapes.load("items");
      await context.sync();

      // 加载所有形状的ID
      for (const shape of shapes.items) {
        shape.load("id");
      }
      await context.sync();

      // 创建ID集合以便快速查找
      const idSet = new Set(elementIds);

      // 查找并删除匹配的形状
      for (const shape of shapes.items) {
        if (idSet.has(shape.id)) {
          shape.delete();
          deletedCount++;
        }
      }

      await context.sync();
    });

    if (deletedCount === 0) {
      return {
        success: false,
        deletedCount: 0,
        message: `未找到任何匹配的元素`,
      };
    }

    return {
      success: true,
      deletedCount,
      message: `成功删除 ${deletedCount} 个元素`,
    };
  } catch (error) {
    console.error("批量删除元素失败:", error);
    return {
      success: false,
      deletedCount: 0,
      message: error instanceof Error ? error.message : "删除失败",
    };
  }
}
