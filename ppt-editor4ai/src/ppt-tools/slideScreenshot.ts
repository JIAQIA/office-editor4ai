/**
 * 文件名: slideScreenshot.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片截图工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint, console */

/**
 * 截图选项
 */
export interface SlideScreenshotOptions {
  /** 幻灯片索引（从 0 开始），不指定则使用当前选中的幻灯片 */
  slideIndex?: number;
  /** 图片宽度（像素），可选 */
  width?: number;
  /** 图片高度（像素），可选 */
  height?: number;
}

/**
 * 截图结果
 */
export interface SlideScreenshotResult {
  /** Base64 编码的 PNG 图片数据（不包含 data URL 前缀） */
  imageBase64: string;
  /** 幻灯片索引 */
  slideIndex: number;
  /** 幻灯片 ID */
  slideId: string;
  /** 图片宽度（像素） */
  width?: number;
  /** 图片高度（像素） */
  height?: number;
}

/**
 * 获取指定幻灯片的截图
 *
 * 使用 PowerPoint.Slide.getImageAsBase64 API
 * 返回 PNG 格式的 Base64 编码图片
 *
 * @param options 截图选项
 * @returns Promise<SlideScreenshotResult> 截图结果
 *
 * @example
 * ```typescript
 * // 获取当前幻灯片的截图
 * const result = await getSlideScreenshot({});
 *
 * // 获取第一张幻灯片的截图，指定尺寸
 * const result = await getSlideScreenshot({
 *   slideIndex: 0,
 *   width: 800,
 *   height: 600
 * });
 *
 * // 使用返回的 Base64 数据
 * const dataUrl = `data:image/png;base64,${result.imageBase64}`;
 * ```
 */
export async function getSlideScreenshot(
  options: SlideScreenshotOptions = {}
): Promise<SlideScreenshotResult> {
  const { slideIndex, width, height } = options;

  console.log("[getSlideScreenshot] 开始获取幻灯片截图", options);

  try {
    return await PowerPoint.run(async (context) => {
      let slide: PowerPoint.Slide;

      // 获取指定幻灯片或当前选中的幻灯片
      if (slideIndex !== undefined) {
        // 使用指定索引
        const slides = context.presentation.slides;
        slide = slides.getItemAt(slideIndex);
        console.log(`[getSlideScreenshot] 获取索引为 ${slideIndex} 的幻灯片`);
      } else {
        // 获取当前选中的幻灯片
        const selectedSlides = context.presentation.getSelectedSlides();
        slide = selectedSlides.getItemAt(0);
        console.log("[getSlideScreenshot] 获取当前选中的幻灯片");
      }

      // 加载幻灯片的 ID 和索引
      slide.load("id,index");

      // 构建截图选项
      const imageOptions: PowerPoint.SlideGetImageOptions = {};
      if (width !== undefined) {
        imageOptions.width = width;
      }
      if (height !== undefined) {
        imageOptions.height = height;
      }

      // 获取截图
      const imageResult = slide.getImageAsBase64(imageOptions);

      // 同步以获取结果
      await context.sync();

      console.log("[getSlideScreenshot] 截图获取成功");
      console.log(`[getSlideScreenshot] 幻灯片索引: ${slide.index}, ID: ${slide.id}`);

      return {
        imageBase64: imageResult.value,
        slideIndex: slide.index,
        slideId: slide.id,
        width,
        height,
      };
    });
  } catch (error) {
    console.error("[getSlideScreenshot] 获取截图失败");
    console.error("[getSlideScreenshot] 错误名称:", (error as Error).name);
    console.error("[getSlideScreenshot] 错误消息:", (error as Error).message);
    console.error("[getSlideScreenshot] 错误堆栈:", (error as Error).stack);

    throw error;
  }
}

/**
 * 获取当前选中幻灯片的截图（简化版本）
 *
 * @param width 图片宽度（像素），可选
 * @param height 图片高度（像素），可选
 * @returns Promise<SlideScreenshotResult> 截图结果
 */
export async function getCurrentSlideScreenshot(
  width?: number,
  height?: number
): Promise<SlideScreenshotResult> {
  return getSlideScreenshot({ width, height });
}

/**
 * 获取指定页码的幻灯片截图
 *
 * @param pageNumber 页码（从 1 开始）
 * @param width 图片宽度（像素），可选
 * @param height 图片高度（像素），可选
 * @returns Promise<SlideScreenshotResult> 截图结果
 */
export async function getSlideScreenshotByPageNumber(
  pageNumber: number,
  width?: number,
  height?: number
): Promise<SlideScreenshotResult> {
  if (pageNumber < 1) {
    throw new Error("页码必须从 1 开始");
  }
  return getSlideScreenshot({ slideIndex: pageNumber - 1, width, height });
}

/**
 * 获取所有幻灯片的截图
 *
 * @param width 图片宽度（像素），可选
 * @param height 图片高度（像素），可选
 * @returns Promise<SlideScreenshotResult[]> 所有幻灯片的截图结果数组
 */
export async function getAllSlidesScreenshots(
  width?: number,
  height?: number
): Promise<SlideScreenshotResult[]> {
  console.log("[getAllSlidesScreenshots] 开始获取所有幻灯片截图");

  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      const slideCount = slides.items.length;
      console.log(`[getAllSlidesScreenshots] 共有 ${slideCount} 张幻灯片`);

      const results: SlideScreenshotResult[] = [];

      // 逐个获取每张幻灯片的截图
      for (let i = 0; i < slideCount; i++) {
        const result = await getSlideScreenshot({ slideIndex: i, width, height });
        results.push(result);
      }

      console.log("[getAllSlidesScreenshots] 所有截图获取完成");
      return results;
    });
  } catch (error) {
    console.error("[getAllSlidesScreenshots] 获取截图失败");
    console.error("[getAllSlidesScreenshots] 错误:", error);
    throw error;
  }
}
