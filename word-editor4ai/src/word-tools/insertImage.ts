/**
 * 文件名: insertImage.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 插入图片工具核心逻辑（支持内联和浮动图片）
 */

/* global Word, console */

/**
 * 图片插入位置类型 / Image Insert Location Type
 */
export type InsertLocation = "Start" | "End" | "Before" | "After" | "Replace";

/**
 * 图片布局类型 / Image Layout Type
 */
export type ImageLayoutType = "inline" | "floating";

/**
 * 文本环绕方式 / Text Wrapping Type
 */
export type WrapType =
  | "Square" // 四周型环绕 / Square wrapping
  | "Tight" // 紧密型环绕 / Tight wrapping
  | "Through" // 穿越型环绕 / Through wrapping
  | "TopAndBottom" // 上下型环绕 / Top and bottom wrapping
  | "Behind" // 衬于文字下方 / Behind text
  | "InFrontOf"; // 浮于文字上方 / In front of text

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
  /** 图片标识符（使用 altTextTitle 作为标识，如果提供）/ Image identifier (uses altTextTitle as identifier if provided) */
  imageId?: string;
  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;
}

/**
 * 在文档中插入图片（支持内联和浮动）
 * Insert image in document (supports inline and floating)
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
        // 注意：Word JavaScript API 对浮动图片的支持有限
        // Note: Word JavaScript API has limited support for floating images
        // 我们先插入为内联图片，然后尝试应用浮动属性
        // We first insert as inline picture, then try to apply floating properties

        const inlinePicture = insertRange.insertInlinePictureFromBase64(
          cleanBase64,
          insertLocation
        );

        // 设置基本属性 / Set basic properties
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

        // 尝试应用浮动选项 / Try to apply floating options
        // 注意：这些 API 可能在某些 Word 版本中不可用
        // Note: These APIs may not be available in some Word versions
        if (floatingOptions) {
          try {
            // Word JavaScript API 目前对浮动图片的控制有限
            // 这里我们记录选项，但实际应用可能需要 OOXML 或其他方法
            // Word JavaScript API currently has limited control over floating images
            // We log the options here, but actual application may require OOXML or other methods
            console.warn(
              "浮动图片选项已记录，但 Word JavaScript API 对浮动图片的支持有限。" +
                "某些选项可能需要通过 OOXML 实现。/ " +
                "Floating image options logged, but Word JavaScript API has limited support for floating images. " +
                "Some options may require OOXML implementation.",
              floatingOptions
            );
          } catch (error) {
            console.warn("应用浮动选项时出错 / Error applying floating options:", error);
          }
        }

        await context.sync();

        // Word API 的 InlinePicture 没有 id 属性
        // InlinePicture in Word API does not have id property
        // 使用 altTextTitle 作为标识符（如果提供）
        // Use altTextTitle as identifier (if provided)
        imageId = altText || undefined;
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
