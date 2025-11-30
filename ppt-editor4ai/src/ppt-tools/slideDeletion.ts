/**
 * 文件名: slideDeletion.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片删除工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint, console */

export interface DeleteSlideOptions {
  slideNumbers?: number[]; // 要删除的幻灯片页码数组（从1开始）
  deleteCurrentSlide?: boolean; // 是否删除当前选中的幻灯片（默认true）
}

export interface DeleteSlideResult {
  success: boolean;
  deletedCount: number;
  failedCount: number;
  message: string;
  details?: {
    deleted: number[];
    notFound: number[];
    errors: Array<{ slideNumber: number; error: string }>;
  };
}

/**
 * 删除指定页码的幻灯片
 * @param slideNumbers 要删除的幻灯片页码数组（从1开始）
 * @returns Promise<DeleteSlideResult> 删除结果
 */
export async function deleteSlidesByNumbers(slideNumbers: number[]): Promise<DeleteSlideResult> {
  try {
    const deletedSlides: number[] = [];
    const notFoundSlides: number[] = [];
    const errors: Array<{ slideNumber: number; error: string }> = [];

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      const totalSlides = slides.items.length;

      // 按照从大到小的顺序排序，避免删除后索引变化
      const sortedNumbers = Array.from(new Set(slideNumbers)).sort((a, b) => b - a);

      for (const slideNumber of sortedNumbers) {
        try {
          const slideIndex = slideNumber - 1;

          // 检查页码是否存在
          if (slideIndex < 0 || slideIndex >= totalSlides) {
            console.warn(`页码 ${slideNumber} 不存在，总共有 ${totalSlides} 页，跳过删除`);
            notFoundSlides.push(slideNumber);
            continue;
          }

          // 删除幻灯片
          slides.items[slideIndex].delete();
          deletedSlides.push(slideNumber);
          console.log(`已标记删除第 ${slideNumber} 页`);
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : "未知错误";
          console.error(`删除第 ${slideNumber} 页失败:`, errorMessage);
          errors.push({ slideNumber, error: errorMessage });
        }
      }

      // 执行删除操作
      await context.sync();
    });

    const deletedCount = deletedSlides.length;
    const failedCount = notFoundSlides.length + errors.length;

    // 构建详细消息
    let message = `删除操作完成: 成功 ${deletedCount} 页`;
    if (notFoundSlides.length > 0) {
      message += `, 页码不存在 ${notFoundSlides.length} 页 (${notFoundSlides.join(", ")})`;
    }
    if (errors.length > 0) {
      message += `, 删除失败 ${errors.length} 页`;
    }

    return {
      success: deletedCount > 0,
      deletedCount,
      failedCount,
      message,
      details: {
        deleted: deletedSlides,
        notFound: notFoundSlides,
        errors,
      },
    };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "未知错误";
    console.error("删除幻灯片失败:", error);
    return {
      success: false,
      deletedCount: 0,
      failedCount: slideNumbers.length,
      message: `删除幻灯片失败: ${errorMessage}`,
    };
  }
}

/**
 * 删除当前选中的幻灯片
 * @returns Promise<DeleteSlideResult> 删除结果
 */
export async function deleteCurrentSlide(): Promise<DeleteSlideResult> {
  try {
    let deletedSlideNumber: number | null = null;

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");

      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");

      await context.sync();

      if (selectedSlides.items.length === 0) {
        throw new Error("没有选中的幻灯片");
      }

      // 获取当前选中的第一个幻灯片
      const currentSlide = selectedSlides.items[0];
      currentSlide.load("id");
      await context.sync();

      // 找到该幻灯片的索引
      for (let i = 0; i < slides.items.length; i++) {
        slides.items[i].load("id");
      }
      await context.sync();

      for (let i = 0; i < slides.items.length; i++) {
        if (slides.items[i].id === currentSlide.id) {
          deletedSlideNumber = i + 1;
          slides.items[i].delete();
          break;
        }
      }

      await context.sync();
    });

    if (deletedSlideNumber === null) {
      return {
        success: false,
        deletedCount: 0,
        failedCount: 1,
        message: "未找到当前选中的幻灯片",
      };
    }

    return {
      success: true,
      deletedCount: 1,
      failedCount: 0,
      message: `成功删除第 ${deletedSlideNumber} 页`,
      details: {
        deleted: [deletedSlideNumber],
        notFound: [],
        errors: [],
      },
    };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "未知错误";
    console.error("删除当前幻灯片失败:", error);
    return {
      success: false,
      deletedCount: 0,
      failedCount: 1,
      message: `删除当前幻灯片失败: ${errorMessage}`,
    };
  }
}

/**
 * 删除幻灯片（通用接口）
 * @param options 删除选项
 * @returns Promise<DeleteSlideResult> 删除结果
 */
export async function deleteSlides(options: DeleteSlideOptions = {}): Promise<DeleteSlideResult> {
  const { slideNumbers, deleteCurrentSlide: shouldDeleteCurrent = true } = options;

  // 如果指定了页码，按页码删除
  if (slideNumbers && slideNumbers.length > 0) {
    return deleteSlidesByNumbers(slideNumbers);
  }

  // 否则删除当前选中的幻灯片
  if (shouldDeleteCurrent) {
    return deleteCurrentSlide();
  }

  return {
    success: false,
    deletedCount: 0,
    failedCount: 0,
    message: "未指定要删除的幻灯片",
  };
}
