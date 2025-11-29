/**
 * 文件名: videoInsertion.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 视频插入工具核心逻辑，与 Office API 交互
 *
 * ⚠️ 功能状态：不可用
 *
 * PowerPoint JavaScript API 目前不支持通过 ShapeCollection 插入媒体元素（视频/音频）。
 * 这是一个已知的功能限制，Microsoft 正在评估该功能请求。
 *
 * 详细信息：
 * - 官方功能请求：https://techcommunity.microsoft.com/idea/microsoft365developerplatform/support-for-inserting-media-elements-via-powerpoint-shapecollection-addmediaelem/4404793
 * - 当前 API 状态：不支持视频/音频插入
 * - 预计支持时间：待定
 *
 * 替代方案：
 * 1. 使用 PowerPoint 桌面版手动插入视频
 * 2. 使用在线视频嵌入（YouTube, Microsoft Stream）
 * 3. 等待 Microsoft 官方 API 支持
 */

/* global console, File, FileReader, fetch */

/**
 * 视频插入选项
 */
export interface VideoInsertionOptions {
  /** 视频来源：Base64 编码的数据（支持带或不带 data URL 前缀） */
  videoSource: string;
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
 * 插入视频结果
 */
export interface VideoInsertionResult {
  /** 插入的视频 ID（Common API 不提供，返回空字符串） */
  videoId: string;
  /** 视频实际宽度 */
  width: number;
  /** 视频实际高度 */
  height: number;
}

/**
 * 插入视频到幻灯片
 *
 * ⚠️ 此功能当前不可用
 *
 * PowerPoint JavaScript API 不支持通过 ShapeCollection 插入媒体元素（视频/音频）。
 * 这是 Microsoft Office JavaScript API 的已知限制。
 *
 * 技术原因：
 * - Office.CoercionType 在 PowerPoint 中仅支持：Text, Matrix, Table, Image, SlideRange
 * - Ooxml 类型仅支持 Word，不支持 PowerPoint
 * - ShapeCollection 没有 addVideo 或 addMedia 方法
 * - PowerPoint.Shapes.addMediaFromBase64() 等方法不存在
 *
 * 官方功能请求：
 * https://techcommunity.microsoft.com/idea/microsoft365developerplatform/support-for-inserting-media-elements-via-powerpoint-shapecollection-addmediaelem/4404793
 *
 * 相关 Issue：
 * - https://github.com/OfficeDev/office-js/issues/5653
 *
 * 替代方案：
 * 1. 使用 PowerPoint 桌面版手动插入视频
 * 2. 使用在线视频嵌入（YouTube, Microsoft Stream）
 * 3. 等待 Microsoft 官方 API 支持
 *
 * @returns Promise<VideoInsertionResult> 插入结果（当前会抛出错误）
 * @throws Error 功能不可用错误
 *
 * @example
 * ```typescript
 * // 此功能当前不可用
 * try {
 *   const result = await insertVideoToSlide({
 *     videoSource: "data:video/mp4;base64,AAAAIGZ0eXBpc29t...",
 *     left: 100,
 *     top: 100,
 *     width: 400,
 *     height: 300
 *   });
 * } catch (error) {
 *   console.error("功能不可用:", error.message);
 * }
 * ```
 * @param _options
 */
export async function insertVideoToSlide(
  _options: VideoInsertionOptions
): Promise<VideoInsertionResult> {
  console.error("[insertVideoToSlide] 功能不可用");
  console.error("[insertVideoToSlide] PowerPoint JavaScript API 不支持插入视频/音频");
  console.error(
    "[insertVideoToSlide] 详情: https://techcommunity.microsoft.com/idea/microsoft365developerplatform/support-for-inserting-media-elements-via-powerpoint-shapecollection-addmediaelem/4404793"
  );

  throw new Error(
    "视频插入功能当前不可用。\n\n" +
      "PowerPoint JavaScript API 不支持通过 ShapeCollection 插入媒体元素（视频/音频）。\n\n" +
      "这是 Microsoft Office JavaScript API 的已知限制。\n\n" +
      "详细信息：\n" +
      "https://techcommunity.microsoft.com/idea/microsoft365developerplatform/support-for-inserting-media-elements-via-powerpoint-shapecollection-addmediaelem/4404793\n\n" +
      "替代方案：\n" +
      "1. 使用 PowerPoint 桌面版手动插入视频\n" +
      "2. 使用在线视频嵌入（YouTube, Microsoft Stream）\n" +
      "3. 等待 Microsoft 官方 API 支持"
  );
}

/**
 * 从文件读取器读取视频并转换为 Base64
 *
 * @param file 视频文件
 * @returns Promise<string> Base64 编码的视频数据（包含 data URL 前缀）
 */
export function readVideoAsBase64(file: File): Promise<string> {
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
 * 从 URL 加载视频并转换为 Base64
 *
 * @param url 视频 URL
 * @returns Promise<string> Base64 编码的视频数据（包含 data URL 前缀）
 */
export async function fetchVideoAsBase64(url: string): Promise<string> {
  try {
    // 使用 fetch 获取视频
    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`获取视频失败: ${response.status} ${response.statusText}`);
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
    console.error("[fetchVideoAsBase64] 获取视频失败:", error);
    throw new Error(`无法从 URL 加载视频: ${(error as Error).message}`);
  }
}

/**
 * 生成视频占位符图片
 *
 * 创建一个带有播放按钮图标的黑色背景图片，作为视频预览
 *
 * @param _width 图片宽度（磅）
 * @param _height 图片高度（磅）
 * @returns Base64 编码的 PNG 图片（不含 data URL 前缀）
 */
function _generateVideoPlaceholder(_width: number, _height: number): string {
  // 此函数暂时未使用，保留供将来可能的占位符功能使用
  throw new Error("generateVideoPlaceholder is not implemented in Office Add-in context");
}

/**
 * 简化版本：插入视频（兼容旧接口）
 *
 * @param videoSource 视频来源（Base64 编码）
 * @param left X 坐标（可选）
 * @param top Y 坐标（可选）
 * @returns Promise<VideoInsertionResult> 插入结果
 */
export async function insertVideo(
  videoSource: string,
  left?: number,
  top?: number
): Promise<VideoInsertionResult> {
  return insertVideoToSlide({ videoSource, left, top });
}
