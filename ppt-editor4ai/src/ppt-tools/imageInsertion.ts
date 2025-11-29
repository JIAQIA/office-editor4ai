/**
 * 文件名: imageInsertion.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 图片插入工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint, console */

/**
 * 图片插入选项
 */
export interface ImageInsertionOptions {
  /** 图片来源：Base64 编码的数据或 URL */
  imageSource: string;
  /** 图片来源类型：'base64' 或 'url' */
  sourceType: "base64" | "url";
  /** X 坐标（可选，单位：磅） */
  left?: number;
  /** Y 坐标（可选，单位：磅） */
  top?: number;
  /** 宽度（可选，单位：磅） */
  width?: number;
  /** 高度（可选，单位：磅） */
  height?: number;
}

/**
 * 插入图片结果
 */
export interface ImageInsertionResult {
  /** 插入的图片形状 ID */
  shapeId: string;
  /** 图片实际宽度 */
  width: number;
  /** 图片实际高度 */
  height: number;
}

/**
 * 插入图片到幻灯片
 * 
 * 使用 Office.context.document.setSelectedDataAsync 和 Office.CoercionType.Image
 * 这是官方推荐的方式，插入的元素类型为 Picture 而不是 Rectangle
 * 
 * @param options 图片插入选项
 * @returns Promise<ImageInsertionResult> 插入结果
 * 
 * @example
 * ```typescript
 * // 使用 Base64 插入图片
 * const result = await insertImageToSlide({
 *   imageSource: "data:image/png;base64,iVBORw0KGgoAAAANS...",
 *   sourceType: "base64",
 *   left: 100,
 *   top: 100,
 *   width: 200,
 *   height: 150
 * });
 * 
 * // 使用 URL 插入图片（建议先用 fetchImageAsBase64 转换）
 * const base64Data = await fetchImageAsBase64("https://example.com/image.png");
 * const result = await insertImageToSlide({
 *   imageSource: base64Data,
 *   sourceType: "base64",
 *   left: 100,
 *   top: 100
 * });
 * ```
 */
export async function insertImageToSlide(
  options: ImageInsertionOptions
): Promise<ImageInsertionResult> {
  const { imageSource, sourceType, left, top, width, height } = options;

  console.log("[insertImageToSlide] 开始插入图片，类型:", sourceType);

  try {
    // 处理 Base64 数据
    let imageData = imageSource;
    
    // 如果包含 data URL 前缀，提取纯 Base64 部分
    if (imageData.includes(",")) {
      imageData = imageData.split(",")[1];
    }
    
    console.log("[insertImageToSlide] 插入图片，来源类型:", sourceType);

    // 使用 Office Common API 插入图片
    // 这会创建一个真正的 Picture 类型元素，而不是 Rectangle
    return new Promise((resolve, reject) => {
      const imageOptions: Office.SetSelectedDataOptions = {
        coercionType: Office.CoercionType.Image,
      };

      // 设置位置和尺寸（如果提供）
      if (left !== undefined) {
        imageOptions.imageLeft = left;
      }
      if (top !== undefined) {
        imageOptions.imageTop = top;
      }
      if (width !== undefined) {
        imageOptions.imageWidth = width;
      }
      if (height !== undefined) {
        imageOptions.imageHeight = height;
      }

      Office.context.document.setSelectedDataAsync(
        imageData,
        imageOptions,
        (asyncResult: Office.AsyncResult<void>) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("[insertImageToSlide] 插入失败:", asyncResult.error?.message);
            reject(new Error(asyncResult.error?.message || "插入图片失败"));
          } else {
            console.log("[insertImageToSlide] 图片插入成功");
            
            // 注意：setSelectedDataAsync 不返回形状信息
            // 我们返回用户指定的尺寸，如果没有指定则返回默认值
            const result: ImageInsertionResult = {
              shapeId: "", // Common API 不提供 ID
              width: width || 200,
              height: height || 150,
            };
            
            resolve(result);
          }
        }
      );
    });
  } catch (error) {
    console.error("[insertImageToSlide] 插入图片失败");
    console.error("[insertImageToSlide] 错误名称:", (error as Error).name);
    console.error("[insertImageToSlide] 错误消息:", (error as Error).message);
    console.error("[insertImageToSlide] 错误堆栈:", (error as Error).stack);

    throw error;
  }
}

/**
 * 从文件读取器读取图片并转换为 Base64
 * 
 * @param file 图片文件
 * @returns Promise<string> Base64 编码的图片数据（包含 data URL 前缀）
 */
export function readImageAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      if (typeof reader.result === "string") {
        resolve(reader.result);
      } else {
        reject(new Error("读取文件失败：结果不是字符串"));
      }
    };

    reader.onerror = () => {
      reject(new Error("读取文件失败"));
    };

    reader.readAsDataURL(file);
  });
}

/**
 * 从 URL 加载图片并转换为 Base64
 * 
 * @param url 图片 URL
 * @returns Promise<string> Base64 编码的图片数据（包含 data URL 前缀）
 */
export async function fetchImageAsBase64(url: string): Promise<string> {
  try {
    // 使用 fetch 获取图片
    const response = await fetch(url);
    
    if (!response.ok) {
      throw new Error(`获取图片失败: ${response.status} ${response.statusText}`);
    }

    // 转换为 Blob
    const blob = await response.blob();

    // 转换为 Base64
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = () => {
        if (typeof reader.result === "string") {
          resolve(reader.result);
        } else {
          reject(new Error("转换失败：结果不是字符串"));
        }
      };

      reader.onerror = () => {
        reject(new Error("转换失败"));
      };

      reader.readAsDataURL(blob);
    });
  } catch (error) {
    console.error("[fetchImageAsBase64] 获取图片失败:", error);
    throw new Error(`无法从 URL 加载图片: ${(error as Error).message}`);
  }
}

/**
 * 简化版本：插入图片（兼容旧接口）
 * 
 * @param imageSource 图片来源（Base64 或 URL）
 * @param sourceType 来源类型
 * @param left X 坐标（可选）
 * @param top Y 坐标（可选）
 * @returns Promise<ImageInsertionResult> 插入结果
 */
export async function insertImage(
  imageSource: string,
  sourceType: "base64" | "url" = "base64",
  left?: number,
  top?: number
): Promise<ImageInsertionResult> {
  return insertImageToSlide({ imageSource, sourceType, left, top });
}
