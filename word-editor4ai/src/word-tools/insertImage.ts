/**
 * 文件名: insertImage.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 插入图片工具核心逻辑（支持内联和浮动图片）
 */

/* global Word, console */

import type { InsertLocation, WrapType } from "./types";

// 重新导出以保持向后兼容性 / Re-export for backward compatibility
export type { InsertLocation, WrapType };

/**
 * 图片布局类型 / Image Layout Type
 */
export type ImageLayoutType = "inline" | "floating";

/**
 * 图片定位选项 / Image Position Options
 */
export interface ImagePositionOptions {
  /** 水平位置（磅）/ Horizontal position in points */
  left?: number;
  /** 垂直位置（磅）/ Vertical position in points */
  top?: number;
  /** 水平对齐方式 / Horizontal alignment */
  horizontalAlignment?: "Left" | "Center" | "Right" | "Inside" | "Outside";
  /** 垂直对齐方式 / Vertical alignment */
  verticalAlignment?: "Top" | "Center" | "Bottom" | "Inside" | "Outside";
  /** 相对于页面/段落/列等 / Relative to page/paragraph/column etc */
  relativeTo?: "Page" | "Margin" | "Column" | "Paragraph";
}

/**
 * 浮动图片选项 / Floating Image Options
 */
export interface FloatingImageOptions {
  /** 文本环绕方式 / Text wrapping type */
  wrapType?: WrapType;
  /** 图片定位选项 / Image position options */
  position?: ImagePositionOptions;
  /** 是否锁定锚点 / Lock anchor */
  lockAnchor?: boolean;
  /** 是否允许与文字重叠 / Allow overlap with text */
  allowOverlap?: boolean;
}

/**
 * 插入图片选项 / Insert Image Options
 */
export interface InsertImageOptions {
  /** Base64 编码的图片数据（必需）/ Base64 encoded image data (required) */
  base64: string;
  /** 图片宽度（磅），可选 / Image width in points, optional */
  width?: number;
  /** 图片高度（磅），可选 / Image height in points, optional */
  height?: number;
  /** 替代文本，可选 / Alt text, optional */
  altText?: string;
  /** 图片描述 / Image description */
  description?: string;
  /** 插入位置，默认为 "End" / Insert location, default "End" */
  insertLocation?: InsertLocation;
  /** 图片布局类型，默认为 "inline" / Image layout type, default "inline" */
  layoutType?: ImageLayoutType;
  /** 浮动图片选项（仅当 layoutType 为 "floating" 时有效）/ Floating image options (only valid when layoutType is "floating") */
  floatingOptions?: FloatingImageOptions;
  /** 是否保持纵横比，默认为 true / Keep aspect ratio, default true */
  keepAspectRatio?: boolean;
  /** 超链接 URL / Hyperlink URL */
  hyperlink?: string;
}

/**
 * 插入图片结果 / Insert Image Result
 */
export interface InsertImageResult {
  /** 是否成功 / Success */
  success: boolean;
  /** 图片标识符（内联图片使用 altText，浮动图片使用 shape ID）/ Image identifier (inline pictures use altText, floating pictures use shape ID) */
  imageId?: string;
  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;
}

/**
 * 在文档中插入图片（支持内联和浮动，支持文字环绕）
 * Insert image in document (supports inline and floating with text wrapping)
 *
 * @remarks
 * - 内联图片（layoutType: "inline"）：使用 insertInlinePictureFromBase64，不支持文字环绕
 * - 浮动图片（layoutType: "floating"）：使用 insertPictureFromBase64，返回 Word.Shape 对象，支持文字环绕
 * - 文字环绕类型通过 floatingOptions.wrapType 设置
 * - Inline pictures (layoutType: "inline"): Uses insertInlinePictureFromBase64, does not support text wrapping
 * - Floating pictures (layoutType: "floating"): Uses insertPictureFromBase64, returns Word.Shape object, supports text wrapping
 * - Text wrapping type is set via floatingOptions.wrapType
 *
 * @example
 * ```typescript
 * // 插入浮动图片，四周型环绕 / Insert floating picture with square wrapping
 * await insertImage({
 *   base64: "...",
 *   layoutType: "floating",
 *   floatingOptions: {
 *     wrapType: "Square"
 *   }
 * });
 * ```
 */
export async function insertImage(options: InsertImageOptions): Promise<InsertImageResult> {
  const {
    base64,
    width,
    height,
    altText,
    description,
    insertLocation = "End",
    layoutType = "inline",
    floatingOptions,
    keepAspectRatio = true,
    hyperlink,
  } = options;

  // 验证参数 / Validate parameters
  if (!base64) {
    return {
      success: false,
      error: "必须提供 base64 图片数据 / Base64 image data is required",
    };
  }

  try {
    let imageId: string | undefined;

    await Word.run(async (context) => {
      // 清理 base64 数据 / Clean base64 data
      let cleanBase64 = base64;
      if (cleanBase64.includes(",")) {
        cleanBase64 = cleanBase64.split(",")[1];
      }

      // 获取插入范围 / Get insert range
      let insertRange: Word.Range;
      const selection = context.document.getSelection();

      switch (insertLocation) {
        case "Start":
          insertRange = context.document.body.getRange("Start");
          break;
        case "End":
          insertRange = context.document.body.getRange("End");
          break;
        case "Before":
          insertRange = selection;
          break;
        case "After":
          insertRange = selection;
          break;
        case "Replace":
          insertRange = selection;
          break;
        default:
          insertRange = context.document.body.getRange("End");
      }

      // 根据布局类型插入图片 / Insert image based on layout type
      if (layoutType === "inline") {
        // 插入内联图片 / Insert inline picture
        const inlinePicture = insertRange.insertInlinePictureFromBase64(
          cleanBase64,
          insertLocation
        );

        // 设置图片属性 / Set image properties
        if (width !== undefined) {
          inlinePicture.width = width;
        }
        if (height !== undefined) {
          inlinePicture.height = height;
        }
        if (altText) {
          inlinePicture.altTextTitle = altText;
        }
        if (description) {
          inlinePicture.altTextDescription = description;
        }
        if (keepAspectRatio) {
          inlinePicture.lockAspectRatio = true;
        }
        if (hyperlink) {
          inlinePicture.hyperlink = hyperlink;
        }

        await context.sync();

        // Word API 的 InlinePicture 没有 id 属性
        // InlinePicture in Word API does not have id property
        // 使用 altTextTitle 作为标识符（如果提供）
        // Use altTextTitle as identifier (if provided)
        imageId = altText || undefined;
      } else if (layoutType === "floating") {
        // 插入浮动图片（返回 Word.Shape 对象）
        // Insert floating picture (returns Word.Shape object)
        // 注意：insertPictureFromBase64 需要 WordApiDesktop 1.2
        // Note: insertPictureFromBase64 requires WordApiDesktop 1.2
        const insertShapeOptions: Word.InsertShapeOptions = {};

        // 设置位置和尺寸 / Set position and size
        if (width !== undefined) {
          insertShapeOptions.width = width;
        }
        if (height !== undefined) {
          insertShapeOptions.height = height;
        }
        if (floatingOptions?.position?.left !== undefined) {
          insertShapeOptions.left = floatingOptions.position.left;
        }
        if (floatingOptions?.position?.top !== undefined) {
          insertShapeOptions.top = floatingOptions.position.top;
        }

        // 插入浮动图片 / Insert floating picture
        const pictureShape = insertRange.insertPictureFromBase64(
          cleanBase64,
          insertShapeOptions
        );

        // 设置形状属性 / Set shape properties
        // 注意：Word.Shape 只有 altTextDescription 属性，没有 altTextTitle
        // Note: Word.Shape only has altTextDescription property, not altTextTitle
        if (description) {
          pictureShape.altTextDescription = description;
        } else if (altText) {
          // 如果没有提供 description，使用 altText 作为 description
          // If description is not provided, use altText as description
          pictureShape.altTextDescription = altText;
        }
        if (keepAspectRatio) {
          pictureShape.lockAspectRatio = true;
        }

        // 设置文字环绕 / Set text wrapping
        if (floatingOptions?.wrapType) {
          try {
            const textWrap = pictureShape.textWrap;
            // 将我们的 WrapType 映射到 Word.ShapeTextWrapType
            // Map our WrapType to Word.ShapeTextWrapType
            const wrapTypeMap: Record<WrapType, Word.ShapeTextWrapType> = {
              Inline: Word.ShapeTextWrapType.inline,
              Square: Word.ShapeTextWrapType.square,
              Tight: Word.ShapeTextWrapType.tight,
              Through: Word.ShapeTextWrapType.through,
              TopBottom: Word.ShapeTextWrapType.topBottom,
              Behind: Word.ShapeTextWrapType.behind,
              Front: Word.ShapeTextWrapType.front,
            };
            textWrap.type = wrapTypeMap[floatingOptions.wrapType];
          } catch (error) {
            console.warn("设置文字环绕时出错 / Error setting text wrapping:", error);
          }
        }

        // 设置其他浮动选项 / Set other floating options
        if (floatingOptions) {
          try {
            if (floatingOptions.allowOverlap !== undefined) {
              pictureShape.allowOverlap = floatingOptions.allowOverlap;
            }
            // 注意：lockAnchor 属性在当前 API 中可能不可用
            // Note: lockAnchor property may not be available in current API
          } catch (error) {
            console.warn("设置浮动选项时出错 / Error setting floating options:", error);
          }
        }

        // 加载形状 ID / Load shape ID
        pictureShape.load("id");
        await context.sync();

        // 使用形状 ID 作为标识符 / Use shape ID as identifier
        imageId = `shape-${pictureShape.id}`;
        
        // 注意：浮动图片作为 Shape 对象，不支持 hyperlink 属性
        // Note: Floating pictures as Shape objects do not support hyperlink property
        if (hyperlink) {
          console.warn(
            "浮动图片不支持超链接属性。如需超链接，请使用内联图片。/ " +
            "Floating pictures do not support hyperlink property. Use inline pictures for hyperlinks."
          );
        }
      }

      await context.sync();
    });

    return {
      success: true,
      imageId,
    };
  } catch (error) {
    console.error("插入图片失败 / Insert image failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 批量插入图片 / Batch Insert Images
 */
export async function insertImages(images: InsertImageOptions[]): Promise<InsertImageResult[]> {
  const results: InsertImageResult[] = [];

  for (const imageOptions of images) {
    const result = await insertImage(imageOptions);
    results.push(result);
  }

  return results;
}
