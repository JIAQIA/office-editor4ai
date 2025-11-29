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

  console.log("[getSlideLayoutInfo] 开始获取布局信息，选项:", options);

  try {
    let layoutInfo: SlideLayoutInfo = {
      slideNumber: 0,
      slideId: "",
      dimensions: { width: 0, height: 0, aspectRatio: "", isFromAPI: false },
      layout: { name: "", type: "" },
      elements: [],
    };

    await PowerPoint.run(async (context) => {
      console.log("[PowerPoint.run] 进入上下文");

      const slides = context.presentation.slides;
      slides.load("items");
      console.log("[PowerPoint.run] 加载 slides.items");
      await context.sync();
      console.log("[PowerPoint.run] 同步完成，幻灯片数量:", slides.items.length);

      // 确定要获取的幻灯片
      let targetSlide: PowerPoint.Slide;
      let actualSlideNumber: number;

      if (slideNumber !== undefined) {
        const slideIndex = slideNumber - 1;
        console.log("[PowerPoint.run] 使用指定页码:", slideNumber, "索引:", slideIndex);

        if (slideIndex < 0 || slideIndex >= slides.items.length) {
          const errorMsg = `页码 ${slideNumber} 不存在，总共有 ${slides.items.length} 页`;
          console.error("[PowerPoint.run] 错误:", errorMsg);
          throw new Error(errorMsg);
        }
        targetSlide = slides.items[slideIndex];
        actualSlideNumber = slideNumber;
      } else {
        console.log("[PowerPoint.run] 使用当前选中的幻灯片");
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();
        console.log("[PowerPoint.run] 选中的幻灯片数量:", selectedSlides.items.length);

        if (selectedSlides.items.length === 0) {
          const errorMsg = "没有选中的幻灯片";
          console.error("[PowerPoint.run] 错误:", errorMsg);
          throw new Error(errorMsg);
        }

        targetSlide = selectedSlides.items[0];
        // 找到当前幻灯片的索引
        targetSlide.load("id");
        await context.sync();
        const targetId = targetSlide.id;
        actualSlideNumber = slides.items.findIndex((s) => s.id === targetId) + 1;
      }

      // 加载幻灯片基本信息
      console.log("[PowerPoint.run] 加载幻灯片基本信息");
      targetSlide.load("id");

      // 获取布局信息
      const layout = targetSlide.layout;
      layout.load("name,type");
      console.log("[PowerPoint.run] 加载布局信息");

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
      console.log("[PowerPoint.run] 获取形状集合");
      const shapes = targetSlide.shapes;
      shapes.load("items");
      await context.sync();
      console.log("[PowerPoint.run] 形状数量:", shapes.items.length);

      // 尝试获取所有图片（如果 API 支持）
      if (includeImages) {
        try {
          const slideAny = targetSlide as any;
          if (slideAny.getPictures) {
            console.log("[PowerPoint.run] 尝试使用 getPictures API");
            const pictures = slideAny.getPictures();
            pictures.load("items");
            await context.sync();
            console.log(`[PowerPoint.run] 通过 getPictures 找到 ${pictures.items.length} 张图片`);
          }
        } catch (error: any) {
          console.log("[PowerPoint.run] getPictures API 不可用:", error.message);
        }
      }

      // 批量加载所有形状的基本属性
      console.log("[PowerPoint.run] 开始加载形状属性");

      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        // 先加载基本属性和类型
        shape.load("id,type,left,top,width,height,name");
      }

      console.log("[PowerPoint.run] 同步形状基本属性");
      await context.sync();
      console.log("[PowerPoint.run] 形状基本属性同步完成");

      // 对于 Placeholder 类型的形状，加载 placeholderFormat
      console.log("[PowerPoint.run] 开始加载 Placeholder 格式信息");
      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        if (shape.type === "Placeholder") {
          console.log(`[PowerPoint.run] 形状 ${i + 1} 是 Placeholder，加载 placeholderFormat`);
          shape.load("placeholderFormat");
          shape.placeholderFormat.load("type,containedType");
        }
      }
      
      try {
        await context.sync();
        console.log("[PowerPoint.run] Placeholder 格式信息加载完成");
      } catch (error: any) {
        console.log("[PowerPoint.run] Placeholder 格式信息加载失败:", error.message);
        console.log("[PowerPoint.run] 将继续处理，但无法获取 Placeholder 的 containedType 信息");
      }

      // 注意：PowerPoint JavaScript API 不支持 placeholderType 属性
      // 我们将通过其他特征来识别图片（不支持 textFrame 和 fill 的 Placeholder）

      // 逐个处理形状，避免批量操作时错误传播
      // 存储每个形状是否成功加载了文本框
      const shapeTextSupport: boolean[] = new Array(shapes.items.length).fill(false);

      console.log("[PowerPoint.run] 开始逐个加载文本框");
      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        const shapeType = shape.type as string;

        console.log(
          `[PowerPoint.run] 形状 ${i + 1}/${shapes.items.length} 类型: ${shapeType}, ID: ${shape.id}`
        );

        // 尝试加载 textFrame，每个形状单独 sync
        try {
          const textFrame = shape.textFrame;
          textFrame.load("textRange");
          await context.sync();

          shapeTextSupport[i] = true;
          console.log(`[PowerPoint.run] ✓ 形状 ${i + 1} textFrame 加载成功`);
        } catch (error: any) {
          console.log(
            `[PowerPoint.run] ✗ 形状 ${i + 1} 不支持 textFrame (${shapeType}):`,
            error.message
          );
          // 不支持 textFrame 的形状，继续处理下一个
        }
      }

      // 加载文本内容
      console.log("[PowerPoint.run] 开始加载文本内容");
      for (let i = 0; i < shapes.items.length; i++) {
        if (shapeTextSupport[i]) {
          const shape = shapes.items[i];
          try {
            const textFrame = shape.textFrame;
            textFrame.textRange.load("text");
            if (includeTextDetails) {
              textFrame.textRange.font.load("name,size,color");
            }
            await context.sync();
            console.log(`[PowerPoint.run] ✓ 形状 ${i + 1} 文本内容加载成功`);
          } catch (error: any) {
            console.log(`[PowerPoint.run] ✗ 形状 ${i + 1} 文本内容加载失败:`, error.message);
          }
        }
      }

      // 注意：填充信息将在后续逐个形状处理时加载，避免批量加载时的错误传播

      console.log("[PowerPoint.run] 所有形状属性加载完成");

      // 收集所有元素信息
      console.log("[PowerPoint.run] 开始收集元素信息");
      const elements: EnhancedElement[] = [];

      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        const shapeType = shape.type as string;

        console.log(
          `[PowerPoint.run] 处理形状 ${i + 1}/${shapes.items.length}, 类型: ${shapeType}, ID: ${shape.id}, 名称: ${shape.name}`
        );

        const element: EnhancedElement = {
          id: shape.id,
          type: shapeType,
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
        // Placeholder 类型的形状可能不支持 fill 属性
        try {
          // 先加载 fill 对象
          shape.fill.load("type");
          await context.sync();

          // 检查 fill 是否为 null
          if (!shape.fill.isNullObject) {
            const fillType = shape.fill.type as string;
            console.log(`[PowerPoint.run] 形状 ${i + 1} 填充类型: ${fillType}`);

            // PowerPoint.FillType 枚举值："solid", "gradient", "pattern", "pictureAndTexture", "noFill"
            if (fillType === "solid") {
              element.fill = { type: "solid" };
              // 颜色信息在当前 API 版本中不可用
            } else if (fillType === "gradient") {
              element.fill = { type: "gradient" };
            } else if (fillType === "pattern" || fillType === "pictureAndTexture") {
              element.fill = { type: "image" };
              console.log(`[PowerPoint.run] 形状 ${i + 1} 检测到图片填充 (${fillType})`);
            } else if (fillType === "noFill") {
              element.fill = { type: "none" };
            } else {
              element.fill = { type: "unknown" };
            }
          }
        } catch (error: any) {
          // Placeholder 等某些类型的形状不支持 fill 属性
          console.log(
            `[PowerPoint.run] 形状 ${i + 1} (${shapeType}) 不支持填充信息: ${error.message}`
          );
        }

        // 检查是否包含图片（无论形状类型）
        if (includeImages) {
          try {
            let isImageShape = false;

            // 1. 检查形状类型是否为图片
            if (shapeType === "Picture" || shapeType === "picture" || shapeType === "Image") {
              console.log(`[PowerPoint.run] ✓ 形状 ${i + 1} 是图片类型 (${shapeType})`);
              isImageShape = true;
              element.image = {
                format: "picture",
              };
            }

            // 2. 检查 Placeholder 的 containedType
            if (shapeType === "Placeholder") {
              console.log(`[PowerPoint.run] 形状 ${i + 1} 开始检查 Placeholder 内容类型`);
              try {
                const placeholderFormat = shape.placeholderFormat;
                
                // 检查 placeholderFormat 是否为 null
                if (!placeholderFormat || placeholderFormat.isNullObject) {
                  console.log(`[PowerPoint.run] 形状 ${i + 1} placeholderFormat 为 null 或未加载`);
                } else {
                  const containedType = placeholderFormat.containedType as string;
                  const placeholderType = placeholderFormat.type as string;

                  console.log(
                    `[PowerPoint.run] 形状 ${i + 1} Placeholder 类型: ${placeholderType}, 包含内容: ${containedType}`
                  );

                  if (containedType === "Image" || containedType === "Picture") {
                    console.log(`[PowerPoint.run] ✓✓ 形状 ${i + 1} Placeholder 包含图片`);
                    isImageShape = true;
                    element.image = {
                      format: "picture-placeholder",
                    };
                  }
                }
              } catch (error: any) {
                console.log(
                  `[PowerPoint.run] 形状 ${i + 1} 读取 placeholderFormat 失败:`,
                  error.message
                );
              }
            }

            // 3. 检查填充是否为图片
            if (element.fill && element.fill.type === "image") {
              console.log(`[PowerPoint.run] ✓ 形状 ${i + 1} 的填充是图片类型`);
              if (!element.image) {
                isImageShape = true;
                element.image = {
                  format: "picture-fill",
                };
              }
            }

            // 4. 如果检测到图片，尝试获取 base64 数据
            if (isImageShape && element.image) {
              // 动态检测 getImageAsBase64 API 是否可用（Beta API）
              if ("getImageAsBase64" in shape && typeof shape?.getImageAsBase64 === "function") {
                try {
                  console.log(`[PowerPoint.run] 尝试获取形状 ${i + 1} 的图片 base64 数据`);
                  const imageBase64Result = shape.getImageAsBase64();
                  await context.sync();
                  const imageData = imageBase64Result.value;

                  if (imageData) {
                    element.image.data = imageData;
                    console.log(
                      `[PowerPoint.run] ✓✓✓ 形状 ${i + 1} 图片 base64 数据获取成功，长度: ${imageData.length}`
                    );
                  }
                } catch (error: any) {
                  console.log(
                    `[PowerPoint.run] 形状 ${i + 1} 获取图片 base64 失败:`,
                    error.message
                  );
                  element.image.data = `图片获取失败: ${error.message}`;
                }
              } else {
                // API 不可用，提供中文描述
                console.log(
                  `[PowerPoint.run] ⚠️ 形状 ${i + 1} getImageAsBase64 API 不可用（需要 Office 版本支持 Beta API）`
                );
                element.image.data =
                  "当前 Office 版本不支持图片数据获取（需要 PowerPointApi Beta 版本）";
              }
            }
          } catch (error: any) {
            console.log(`[PowerPoint.run] 形状 ${i + 1} 检查图片信息时出错: ${error.message}`);
          }
        }

        elements.push(element);
      }

      layoutInfo.elements = elements;
      console.log("[PowerPoint.run] 元素信息收集完成，共", elements.length, "个元素");
    });

    console.log("[getSlideLayoutInfo] 布局信息获取成功");

    return layoutInfo;
  } catch (error) {
    console.error("[getSlideLayoutInfo] 获取页面布局信息失败");
    console.error("[getSlideLayoutInfo] 错误名称:", error.name);
    console.error("[getSlideLayoutInfo] 错误消息:", error.message);
    console.error("[getSlideLayoutInfo] 错误堆栈:", error.stack);

    // 打印 Office.js 特定的调试信息
    if (error.debugInfo) {
      console.error(
        "[getSlideLayoutInfo] Office.js 调试信息:",
        JSON.stringify(error.debugInfo, null, 2)
      );
    }

    // 打印完整的错误对象
    console.error(
      "[getSlideLayoutInfo] 完整错误对象:",
      JSON.stringify(error, Object.getOwnPropertyNames(error), 2)
    );

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
