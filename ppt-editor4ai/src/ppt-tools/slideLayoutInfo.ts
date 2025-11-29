/**
 * 文件名: slideLayoutInfo.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 页面布局信息获取工具，提供完整的页面尺寸、布局类型和元素详细信息
 */

/* global PowerPoint, console */

/**
 * 页面尺寸信息
 */
export interface SlideDimensions {
  width: number; // 宽度（points）
  height: number; // 高度（points）
  aspectRatio: string; // 宽高比，如 "16:9", "4:3"
  isFromAPI: boolean; // 是否通过 Office API 获取（true: API获取, false: 默认值）
}

/**
 * 相对位置信息（百分比）
 */
export interface RelativePosition {
  leftPercent: number; // 相对于页面宽度的百分比
  topPercent: number; // 相对于页面高度的百分比
  widthPercent: number; // 相对于页面宽度的百分比
  heightPercent: number; // 相对于页面高度的百分比
}

/**
 * 文本信息
 */
export interface TextInfo {
  content: string;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  alignment?: string;
}

/**
 * 图片信息
 */
export interface ImageInfo {
  format: string; // 图片格式
  data?: string; // Base64 编码数据
  url?: string; // 外部链接（如果有）
}

/**
 * 填充信息
 */
export interface FillInfo {
  type: "solid" | "gradient" | "image" | "none" | "unknown";
  color?: string;
}

/**
 * 增强的元素信息
 */
export interface EnhancedElement {
  // 基础信息
  id: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  name?: string;

  // 相对位置（百分比）
  relativePosition: RelativePosition;

  // 旋转角度
  rotation?: number;

  // Z轴顺序
  zOrder?: number;

  // 文本内容（如果有）
  text?: TextInfo;

  // 图片信息（如果是图片类型）
  image?: ImageInfo;

  // 填充信息
  fill?: FillInfo;
}

/**
 * 布局信息
 */
export interface LayoutInfo {
  name: string; // 布局名称
  type: string; // 布局类型
}

/**
 * 背景信息
 */
export interface BackgroundInfo {
  type: "solid" | "gradient" | "image" | "pattern" | "none" | "unknown";
  color?: string;
  imageData?: string; // Base64 或 URL
}

/**
 * 完整的页面布局信息
 */
export interface SlideLayoutInfo {
  // 页面基础信息
  slideNumber: number;
  slideId: string;

  // 尺寸信息
  dimensions: SlideDimensions;

  // 布局信息
  layout: LayoutInfo;

  // 元素列表
  elements: EnhancedElement[];

  // 背景信息
  background?: BackgroundInfo;
}

/**
 * 获取布局信息的选项
 */
export interface GetLayoutInfoOptions {
  slideNumber?: number; // 幻灯片页码（从1开始），不填则使用当前页
  includeImages?: boolean; // 是否包含图片的 Base64 数据，默认为 false
  includeBackground?: boolean; // 是否包含背景信息，默认为 false
  includeTextDetails?: boolean; // 是否包含文本详细信息，默认为 false
}

/**
 * 计算宽高比
 */
function calculateAspectRatio(width: number, height: number): string {
  const gcd = (a: number, b: number): number => (b === 0 ? a : gcd(b, a % b));
  const divisor = gcd(Math.round(width), Math.round(height));
  const w = Math.round(width / divisor);
  const h = Math.round(height / divisor);

  // 识别常见比例
  if (w === 16 && h === 9) return "16:9";
  if (w === 4 && h === 3) return "4:3";
  if (w === 16 && h === 10) return "16:10";
  if (w === 1 && h === 1) return "1:1";

  return `${w}:${h}`;
}

/**
 * 计算相对位置（百分比）
 */
function calculateRelativePosition(
  left: number,
  top: number,
  width: number,
  height: number,
  slideWidth: number,
  slideHeight: number
): RelativePosition {
  return {
    leftPercent: Math.round((left / slideWidth) * 10000) / 100,
    topPercent: Math.round((top / slideHeight) * 10000) / 100,
    widthPercent: Math.round((width / slideWidth) * 10000) / 100,
    heightPercent: Math.round((height / slideHeight) * 10000) / 100,
  };
}

/**
 * 获取演示文稿的全局尺寸
 * 动态检测是否支持 pageSetup API（未来可能进入正式版）
 * 如果不支持则使用标准 16:9 宽屏尺寸作为降级方案
 */
export async function getPresentationDimensions(): Promise<SlideDimensions> {
  try {
    let dimensions: SlideDimensions = {
      width: 0,
      height: 0,
      aspectRatio: "",
      isFromAPI: false,
    };

    await PowerPoint.run(async (context) => {
      try {
        // 动态检测 pageSetup API 是否存在
        const presentation = context.presentation as any;
        console.log("[DEBUG] 检测 pageSetup API...");
        if (presentation.pageSetup) {
          console.log("[DEBUG] pageSetup API 可用，正在获取尺寸...");
          const pageSetup = presentation.pageSetup;
          pageSetup.load("slideWidth,slideHeight");
          await context.sync();

          dimensions = {
            width: pageSetup.slideWidth,
            height: pageSetup.slideHeight,
            aspectRatio: calculateAspectRatio(pageSetup.slideWidth, pageSetup.slideHeight),
            isFromAPI: true, // 通过 API 成功获取
          };
          console.log("[DEBUG] 成功获取尺寸:", dimensions);
        } else {
          console.log("[DEBUG] pageSetup API 不存在");
          throw new Error("pageSetup API not available");
        }
      } catch (error) {
        // 降级方案：使用标准 16:9 宽屏尺寸
        console.warn("[降级方案] pageSetup API 不可用，使用标准 16:9 尺寸:", error);
        dimensions = {
          width: 720,
          height: 405,
          aspectRatio: "16:9",
          isFromAPI: false, // 使用默认值
        };
        console.log("[降级方案] 使用默认尺寸:", dimensions);
      }
    });

    return dimensions;
  } catch (error) {
    console.error("获取演示文稿尺寸失败:", error);
    throw error;
  }
}

/**
 * 获取指定幻灯片的尺寸信息（轻量级）
 */
export async function getSlideDimensions(): Promise<SlideDimensions> {
  // PowerPoint 中所有幻灯片共享相同的尺寸，直接返回全局尺寸
  return getPresentationDimensions();
}

/**
 * 获取完整的页面布局信息
 */
export async function getSlideLayoutInfo(
  options: GetLayoutInfoOptions = {}
): Promise<SlideLayoutInfo> {
  const {
    slideNumber,
    includeImages = false,
    includeBackground = false,
    includeTextDetails = false,
  } = options;

  try {
    let layoutInfo: SlideLayoutInfo = {
      slideNumber: 0,
      slideId: "",
      dimensions: { width: 0, height: 0, aspectRatio: "", isFromAPI: false },
      layout: { name: "", type: "" },
      elements: [],
    };

    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // 确定要获取的幻灯片
      let targetSlide: PowerPoint.Slide;
      let actualSlideNumber: number;

      if (slideNumber !== undefined) {
        const slideIndex = slideNumber - 1;
        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          throw new Error(`页码 ${slideNumber} 不存在，总共有 ${slides.items.length} 页`);
        }
        targetSlide = slides.items[slideIndex];
        actualSlideNumber = slideNumber;
      } else {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length === 0) {
          throw new Error("没有选中的幻灯片");
        }

        targetSlide = selectedSlides.items[0];
        // 找到当前幻灯片的索引
        targetSlide.load("id");
        await context.sync();
        const targetId = targetSlide.id;
        actualSlideNumber = slides.items.findIndex((s) => s.id === targetId) + 1;
      }

      // 加载幻灯片基本信息
      targetSlide.load("id");

      // 获取布局信息
      const layout = targetSlide.layout;
      layout.load("name,type");

      // 获取尺寸信息
      // 动态检测 pageSetup API 是否存在（未来可能进入正式版）
      let slideWidth = 720;
      let slideHeight = 405;
      let isFromAPI = false;

      try {
        const presentation = context.presentation as any;
        if (presentation.pageSetup) {
          const pageSetup = presentation.pageSetup;
          pageSetup.load("slideWidth,slideHeight");
          await context.sync();

          slideWidth = pageSetup.slideWidth;
          slideHeight = pageSetup.slideHeight;
          isFromAPI = true; // 通过 API 成功获取
        } else {
          console.warn("pageSetup API 不可用，使用标准 16:9 尺寸");
          await context.sync();
        }
      } catch (error) {
        console.warn("pageSetup API 调用失败，使用标准 16:9 尺寸:", error);
        await context.sync();
      }

      layoutInfo.slideNumber = actualSlideNumber;
      layoutInfo.slideId = targetSlide.id;
      layoutInfo.dimensions = {
        width: slideWidth,
        height: slideHeight,
        aspectRatio: calculateAspectRatio(slideWidth, slideHeight),
        isFromAPI, // 标识数据来源
      };
      layoutInfo.layout = {
        name: layout.name || "Unknown",
        type: layout.type || "Unknown",
      };

      // 获取背景信息（如果需要）
      // 注意：PowerPoint JavaScript API 目前不支持直接访问背景信息
      if (includeBackground) {
        // 背景信息功能暂不可用
        layoutInfo.background = {
          type: "unknown",
        };
      }

      // 获取形状集合
      const shapes = targetSlide.shapes;
      shapes.load("items");
      await context.sync();

      // 批量加载所有形状的基本属性
      for (const shape of shapes.items) {
        shape.load("id,type,left,top,width,height,name");

        // 加载文本框
        try {
          shape.textFrame.load("textRange");
        } catch {
          // 形状没有文本框
        }

        // 如果需要图片数据
        if (includeImages && shape.type === PowerPoint.ShapeType.geometricShape) {
          try {
            shape.fill.load("type");
          } catch {
            // 无法加载填充信息
          }
        }
      }
      await context.sync();

      // 加载文本内容
      for (const shape of shapes.items) {
        try {
          shape.textFrame.textRange.load("text");
          if (includeTextDetails) {
            shape.textFrame.textRange.font.load("name,size,color");
          }
        } catch {
          // 形状没有文本框
        }
      }
      await context.sync();

      // 收集所有元素信息
      const elements: EnhancedElement[] = [];

      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];

        const element: EnhancedElement = {
          id: shape.id,
          type: shape.type,
          left: Math.round(shape.left * 100) / 100,
          top: Math.round(shape.top * 100) / 100,
          width: Math.round(shape.width * 100) / 100,
          height: Math.round(shape.height * 100) / 100,
          name: shape.name || undefined,
          // rotation 属性在 PowerPoint JavaScript API 中不可用
          zOrder: i, // 使用索引作为 Z 轴顺序
          relativePosition: calculateRelativePosition(
            shape.left,
            shape.top,
            shape.width,
            shape.height,
            slideWidth,
            slideHeight
          ),
        };

        // 获取文本内容
        try {
          const textContent = shape.textFrame.textRange.text?.trim();
          if (textContent) {
            element.text = {
              content: textContent,
            };

            if (includeTextDetails) {
              try {
                const font = shape.textFrame.textRange.font;
                element.text.fontSize = font.size;
                element.text.fontFamily = font.name;
                element.text.color = font.color;
              } catch {
                // 无法获取字体详细信息
              }
            }
          }
        } catch {
          // 形状没有文本框
        }

        // 获取填充信息
        // 注意：PowerPoint JavaScript API 对填充信息的支持有限
        // 目前只能获取填充类型，无法获取具体颜色值
        try {
          shape.fill.load("type");
          await context.sync();

          const fillType = shape.fill.type as string;
          // PowerPoint.FillType 枚举值："solid", "gradient", "pattern", "pictureAndTexture", "noFill"
          if (fillType === "solid") {
            element.fill = { type: "solid" };
            // 颜色信息在当前 API 版本中不可用
          } else if (fillType === "gradient") {
            element.fill = { type: "gradient" };
          } else if (fillType === "pattern" || fillType === "pictureAndTexture") {
            element.fill = { type: "image" };
          } else if (fillType === "noFill") {
            element.fill = { type: "none" };
          } else {
            element.fill = { type: "unknown" };
          }
        } catch {
          // 无法获取填充信息
        }

        // 如果是图片类型且需要包含图片数据
        if (includeImages && shape.type === PowerPoint.ShapeType.geometricShape) {
          try {
            // 尝试获取图片数据（这里需要根据实际 API 调整）
            // PowerPoint JS API 对图片的支持有限，可能需要其他方法
            element.image = {
              format: "unknown",
            };
          } catch {
            // 无法获取图片数据
          }
        }

        elements.push(element);
      }

      layoutInfo.elements = elements;
    });

    return layoutInfo;
  } catch (error) {
    console.error("获取页面布局信息失败:", error);
    throw error;
  }
}

/**
 * 获取当前选中幻灯片的布局信息（简化版本）
 */
export async function getCurrentSlideLayoutInfo(): Promise<SlideLayoutInfo> {
  return getSlideLayoutInfo({
    includeImages: false,
    includeBackground: false,
    includeTextDetails: false,
  });
}
