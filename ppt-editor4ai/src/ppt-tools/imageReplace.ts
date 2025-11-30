/**
 * 文件名: imageReplace.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 图片替换工具核心逻辑，与 Office API 交互
 *       支持替换普通图片（Picture/Image）和所有类型的占位符（Placeholder）
 *       - Placeholder-Picture: 使用 fill.setImage() 方法
 *       - Placeholder-Content: 使用删除+插入方式
 */

/* global PowerPoint, Office, console */

/**
 * 图片替换选项
 */
export interface ImageReplaceOptions {
  /** 要替换的元素ID（可选，如果不提供则使用当前选中的元素） */
  elementId?: string;
  /** 新图片的 Base64 数据（支持带或不带 data URL 前缀） */
  imageSource: string;
  /** 是否保持原图片的位置和尺寸（默认 true） */
  keepDimensions?: boolean;
  /** 新图片的宽度（可选，仅在 keepDimensions 为 false 时生效） */
  width?: number;
  /** 新图片的高度（可选，仅在 keepDimensions 为 false 时生效） */
  height?: number;
}

/**
 * 图片替换结果
 */
export interface ImageReplaceResult {
  success: boolean;
  message: string;
  elementId?: string;
  elementType?: string;
  originalDimensions?: {
    left: number;
    top: number;
    width: number;
    height: number;
  };
}

/**
 * 替换图片
 * 支持普通图片（Picture/Image）和所有类型的占位符（Placeholder）
 *
 * 处理方式：
 * - Picture/Image 类型：删除原图片，在相同位置插入新图片
 * - Placeholder-Picture 类型：使用 fill.setImage() 直接替换填充
 * - Placeholder-Content 等其他类型：删除占位符，在相同位置插入新图片
 *
 * @param options 替换选项
 * @returns Promise<ImageReplaceResult>
 *
 * @example
 * ```typescript
 * // 替换选中的图片
 * const result = await replaceImage({
 *   imageSource: "data:image/png;base64,iVBORw0KGgoAAAANS...",
 *   keepDimensions: true
 * });
 *
 * // 替换指定ID的图片
 * const result = await replaceImage({
 *   elementId: "shape123",
 *   imageSource: "iVBORw0KGgoAAAANS...",
 *   keepDimensions: false,
 *   width: 300,
 *   height: 200
 * });
 * ```
 */
export async function replaceImage(options: ImageReplaceOptions): Promise<ImageReplaceResult> {
  const { elementId, imageSource, keepDimensions = true, width, height } = options;

  if (!imageSource) {
    return {
      success: false,
      message: "图片数据不能为空",
    };
  }

  try {
    let result: ImageReplaceResult = {
      success: false,
      message: "",
    };

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

      if (elementId) {
        // 通过ID查找
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
      } else {
        // 使用当前选中的元素
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (selectedShapes.items.length === 0) {
          throw new Error("请先选中一个图片元素");
        }

        if (selectedShapes.items.length > 1) {
          throw new Error("请只选中一个图片元素");
        }

        targetShape = selectedShapes.items[0];
        targetShape.load("id,type");
        await context.sync();
      }

      // 验证元素类型
      const supportedTypes = ["Picture", "Image", "Placeholder"];
      if (!supportedTypes.includes(targetShape.type)) {
        throw new Error(
          `选中的元素类型为 ${targetShape.type}，不支持图片替换。请选择图片或占位符。`
        );
      }

      // 加载元素的位置和尺寸信息
      targetShape.load("left,top,width,height,name");
      await context.sync();

      const originalDimensions = {
        left: targetShape.left,
        top: targetShape.top,
        width: targetShape.width,
        height: targetShape.height,
      };

      // 处理 Base64 数据
      let imageData = imageSource;
      if (imageData.includes(",")) {
        imageData = imageData.split(",")[1];
      }

      // 计算新的尺寸
      const newWidth = keepDimensions
        ? originalDimensions.width
        : width || originalDimensions.width;
      const newHeight = keepDimensions
        ? originalDimensions.height
        : height || originalDimensions.height;

      // 对于 Placeholder 类型，需要根据占位符类型采用不同的方法
      if (targetShape.type === "Placeholder") {
        // 加载占位符类型信息
        const placeholderFormat = targetShape.placeholderFormat;
        placeholderFormat.load("type");
        await context.sync();

        console.log(`[replaceImage] Placeholder 类型: ${placeholderFormat.type}`);

        // Picture 类型的占位符可以直接使用 fill.setImage()
        if (placeholderFormat.type === "Picture") {
          targetShape.fill.setImage(imageData);
          await context.sync();

          // 调整尺寸
          targetShape.width = newWidth;
          targetShape.height = newHeight;
          await context.sync();

          result = {
            success: true,
            message: "图片替换成功（Placeholder-Picture）",
            elementId: targetShape.id,
            elementType: targetShape.type,
            originalDimensions,
          };
        } else {
          // Content 等其他类型的占位符，需要在其位置插入图片
          // 先选中目标形状，以便在其位置插入新图片
          slide.setSelectedShapes([targetShape.id]);
          await context.sync();

          // 使用 setSelectedDataAsync 在当前位置插入图片
          await new Promise<void>((resolve, reject) => {
            Office.context.document.setSelectedDataAsync(
              imageData,
              {
                coercionType: Office.CoercionType.Image,
              },
              (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  reject(new Error(asyncResult.error.message));
                } else {
                  resolve();
                }
              }
            );
          });

          // 等待图片插入完成
          await context.sync();

          // 获取新插入的图片（应该是最后一个形状）
          shapes.load("items");
          await context.sync();

          const newImage = shapes.items[shapes.items.length - 1];
          newImage.load("id,type");
          await context.sync();

          // 设置位置和尺寸
          newImage.left = originalDimensions.left;
          newImage.top = originalDimensions.top;
          newImage.width = newWidth;
          newImage.height = newHeight;

          await context.sync();

          // 删除旧占位符
          targetShape.delete();
          await context.sync();

          // 选中新插入的图片，方便后续操作
          slide.setSelectedShapes([newImage.id]);
          await context.sync();

          result = {
            success: true,
            message: `图片替换成功（Placeholder-${placeholderFormat.type}）`,
            elementId: newImage.id,
            elementType: newImage.type,
            originalDimensions,
          };
        }
      } else {
        // 对于普通 Picture/Image 类型，使用删除+插入的方式
        // 先选中目标形状，以便在其位置插入新图片
        slide.setSelectedShapes([targetShape.id]);
        await context.sync();

        // 使用 setSelectedDataAsync 在当前位置插入图片
        // 这是 PowerPoint JavaScript API 推荐的插入图片方法
        await new Promise<void>((resolve, reject) => {
          Office.context.document.setSelectedDataAsync(
            imageData,
            {
              coercionType: Office.CoercionType.Image,
            },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(new Error(asyncResult.error.message));
              } else {
                resolve();
              }
            }
          );
        });

        // 等待图片插入完成
        await context.sync();

        // 获取新插入的图片（应该是最后一个形状）
        shapes.load("items");
        await context.sync();

        const newImage = shapes.items[shapes.items.length - 1];
        newImage.load("id,type");
        await context.sync();

        // 设置位置和尺寸
        newImage.left = originalDimensions.left;
        newImage.top = originalDimensions.top;
        newImage.width = newWidth;
        newImage.height = newHeight;

        await context.sync();

        // 删除旧图片
        targetShape.delete();
        await context.sync();

        // 选中新插入的图片，方便后续操作
        slide.setSelectedShapes([newImage.id]);
        await context.sync();

        result = {
          success: true,
          message: "图片替换成功",
          elementId: newImage.id,
          elementType: newImage.type,
          originalDimensions,
        };
      }
    });

    return result;
  } catch (error) {
    console.error("[replaceImage] 替换图片失败:", error);
    return {
      success: false,
      message: error instanceof Error ? error.message : "未知错误",
      elementId,
    };
  }
}

/**
 * 替换当前选中的图片
 *
 * @param imageSource 新图片的 Base64 数据
 * @param keepDimensions 是否保持原图片的位置和尺寸
 * @returns Promise<ImageReplaceResult>
 */
export async function replaceSelectedImage(
  imageSource: string,
  keepDimensions = true
): Promise<ImageReplaceResult> {
  return replaceImage({ imageSource, keepDimensions });
}

/**
 * 批量替换多个图片
 *
 * @param replacements 替换选项数组
 * @returns Promise<ImageReplaceResult[]>
 */
export async function replaceImages(
  replacements: ImageReplaceOptions[]
): Promise<ImageReplaceResult[]> {
  const results: ImageReplaceResult[] = [];

  for (const replacement of replacements) {
    const result = await replaceImage(replacement);
    results.push(result);
  }

  return results;
}

/**
 * 获取图片元素的信息
 *
 * @param elementId 元素ID（可选，如果不提供则使用当前选中的元素）
 * @returns Promise<ImageElementInfo | null>
 */
export interface ImageElementInfo {
  elementId: string;
  elementType: string;
  name: string;
  left: number;
  top: number;
  width: number;
  height: number;
  isPlaceholder: boolean;
  placeholderType?: string;
}

export async function getImageInfo(elementId?: string): Promise<ImageElementInfo | null> {
  try {
    let imageInfo: ImageElementInfo | null = null;

    await PowerPoint.run(async (context) => {
      let targetShape: PowerPoint.Shape | null = null;

      if (elementId) {
        // 通过ID查找
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        // eslint-disable-next-line office-addins/no-navigational-load
        slide.load("shapes");
        await context.sync();

        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        for (const shape of shapes.items) {
          shape.load("id");
        }
        await context.sync();

        for (const shape of shapes.items) {
          if (shape.id === elementId) {
            targetShape = shape;
            break;
          }
        }
      } else {
        // 使用当前选中的元素
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (selectedShapes.items.length > 0) {
          targetShape = selectedShapes.items[0];
        }
      }

      if (!targetShape) {
        // 没有选中元素时返回 null，而不是抛出错误
        return;
      }

      // 加载元素信息
      targetShape.load("id,type,name,left,top,width,height");
      await context.sync();

      const isPlaceholder = targetShape.type === "Placeholder";
      let placeholderType: string | undefined;

      if (isPlaceholder) {
        const placeholderFormat = targetShape.placeholderFormat;
        placeholderFormat.load("type");
        await context.sync();
        placeholderType = placeholderFormat.type;
      }

      imageInfo = {
        elementId: targetShape.id,
        elementType: targetShape.type,
        name: targetShape.name,
        left: targetShape.left,
        top: targetShape.top,
        width: targetShape.width,
        height: targetShape.height,
        isPlaceholder,
        placeholderType,
      };
    });

    return imageInfo;
  } catch (error) {
    console.error("[getImageInfo] 获取图片信息失败:", error);
    return null;
  }
}
