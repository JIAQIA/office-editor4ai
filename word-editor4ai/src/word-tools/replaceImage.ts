/**
 * 文件名: replaceImage.ts
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 统一的图片替换工具，支持选区、搜索、索引、范围四种定位方式
 */

/* global Word, console */

import type { RangeLocator } from "./types";

/**
 * 图片属性 / Image Properties
 */
export interface ImageProperties {
  /** 宽度（磅）/ Width in points */
  width?: number;
  /** 高度（磅）/ Height in points */
  height?: number;
  /** 替代文本 / Alt text */
  altText?: string;
  /** 超链接 / Hyperlink */
  hyperlink?: string;
  /** 是否锁定纵横比 / Lock aspect ratio */
  lockAspectRatio?: boolean;
}

/**
 * 图片搜索选项 / Image Search Options
 */
export interface ImageSearchOptions {
  /** 按替代文本搜索 / Search by alt text */
  altText?: string;
  /** 最小宽度（磅）/ Minimum width in points */
  minWidth?: number;
  /** 最大宽度（磅）/ Maximum width in points */
  maxWidth?: number;
  /** 最小高度（磅）/ Minimum height in points */
  minHeight?: number;
  /** 最大高度（磅）/ Maximum height in points */
  maxHeight?: number;
}

/**
 * 替换图片定位器 / Replace Image Locator
 */
export type ReplaceImageLocator =
  | {
      /** 定位类型：当前选区 / Locator type: current selection */
      type: "selection";
    }
  | {
      /** 定位类型：按索引 / Locator type: by index */
      type: "index";
      /** 图片索引（从0开始）/ Image index (0-based) */
      index: number;
    }
  | {
      /** 定位类型：搜索匹配 / Locator type: search match */
      type: "search";
      /** 搜索选项 / Search options */
      searchOptions: ImageSearchOptions;
    }
  | {
      /** 定位类型：指定范围 / Locator type: specific range */
      type: "range";
      /** 范围定位器 / Range locator */
      rangeLocator: RangeLocator;
    };

/**
 * 替换图片选项 / Replace Image Options
 */
export interface ReplaceImageOptions {
  /** 定位方式 / Locator */
  locator: ReplaceImageLocator;
  /** 新图片数据（Base64），可选 / New image data (Base64), optional */
  newImageData?: string;
  /** 图片属性（可选）/ Image properties (optional) */
  properties?: ImageProperties;
  /** 是否替换所有匹配项（仅 search 和 range 模式）/ Replace all matches (search and range mode only) */
  replaceAll?: boolean;
}

/**
 * 替换结果 / Replace Result
 */
export interface ReplaceImageResult {
  /** 替换的数量 / Number of replacements */
  count: number;
  /** 是否成功 / Success */
  success: boolean;
  /** 错误信息（如果有）/ Error message (if any) */
  error?: string;
}

/**
 * 应用图片属性 / Apply image properties
 */
function applyImageProperties(image: Word.InlinePicture, properties: ImageProperties): void {
  if (properties.width !== undefined) {
    image.width = properties.width;
  }
  if (properties.height !== undefined) {
    image.height = properties.height;
  }
  if (properties.altText !== undefined) {
    image.altTextDescription = properties.altText;
  }
  if (properties.hyperlink !== undefined) {
    image.hyperlink = properties.hyperlink;
  }
  if (properties.lockAspectRatio !== undefined) {
    image.lockAspectRatio = properties.lockAspectRatio;
  }
}

/**
 * 检查图片是否匹配搜索条件 / Check if image matches search criteria
 */
async function matchesSearchCriteria(
  context: Word.RequestContext,
  image: Word.InlinePicture,
  options: ImageSearchOptions
): Promise<boolean> {
  image.load(["width", "height", "altTextDescription"]);
  await context.sync();

  if (options.altText !== undefined) {
    if (!image.altTextDescription || !image.altTextDescription.includes(options.altText)) {
      return false;
    }
  }

  if (options.minWidth !== undefined && image.width < options.minWidth) {
    return false;
  }

  if (options.maxWidth !== undefined && image.width > options.maxWidth) {
    return false;
  }

  if (options.minHeight !== undefined && image.height < options.minHeight) {
    return false;
  }

  if (options.maxHeight !== undefined && image.height > options.maxHeight) {
    return false;
  }

  return true;
}

/**
 * 根据范围定位器获取范围 / Get range by range locator
 */
async function getRangeByLocator(
  context: Word.RequestContext,
  locator: RangeLocator
): Promise<Word.Range> {
  switch (locator.type) {
    case "bookmark": {
      const bookmark = context.document.getBookmarkRangeOrNullObject(locator.name);
      bookmark.load("text");
      await context.sync();
      if (bookmark.isNullObject) {
        throw new Error(
          `书签 "${locator.name}" 不存在 / Bookmark "${locator.name}" does not exist`
        );
      }
      return bookmark;
    }

    case "heading": {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const matchedParagraphs: Word.Paragraph[] = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load(["style", "text"]);
      }
      await context.sync();

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const style = para.style.toLowerCase();

        if (style.includes("heading") || style.includes("标题")) {
          if (locator.level) {
            const levelMatch = style.match(/\d+/);
            if (levelMatch && parseInt(levelMatch[0]) === locator.level) {
              if (!locator.text || para.text.includes(locator.text)) {
                matchedParagraphs.push(para);
              }
            }
          } else if (!locator.text || para.text.includes(locator.text)) {
            matchedParagraphs.push(para);
          }
        }
      }

      if (matchedParagraphs.length === 0) {
        throw new Error("未找到匹配的标题 / No matching heading found");
      }

      const index = locator.index || 0;
      if (index >= matchedParagraphs.length) {
        throw new Error(`标题索引 ${index} 超出范围 / Heading index ${index} out of range`);
      }

      return matchedParagraphs[index].getRange();
    }

    case "paragraph": {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      if (locator.startIndex >= paragraphs.items.length) {
        throw new Error(
          `段落索引 ${locator.startIndex} 超出范围 / Paragraph index ${locator.startIndex} out of range`
        );
      }

      const startPara = paragraphs.items[locator.startIndex];
      if (locator.endIndex !== undefined) {
        if (locator.endIndex >= paragraphs.items.length) {
          throw new Error(
            `段落索引 ${locator.endIndex} 超出范围 / Paragraph index ${locator.endIndex} out of range`
          );
        }
        const endPara = paragraphs.items[locator.endIndex];
        const startRange = startPara.getRange("Start");
        const endRange = endPara.getRange("End");
        return startRange.expandTo(endRange);
      } else {
        return startPara.getRange();
      }
    }

    case "section": {
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      if (locator.index >= sections.items.length) {
        throw new Error(
          `节索引 ${locator.index} 超出范围 / Section index ${locator.index} out of range`
        );
      }

      const section = sections.items[locator.index];
      return section.body.getRange();
    }

    case "contentControl": {
      const contentControls = context.document.contentControls;
      contentControls.load("items");
      await context.sync();

      const matchedControls: Word.ContentControl[] = [];

      for (let i = 0; i < contentControls.items.length; i++) {
        const control = contentControls.items[i];
        control.load(["title", "tag"]);
      }
      await context.sync();

      for (let i = 0; i < contentControls.items.length; i++) {
        const control = contentControls.items[i];
        if (locator.title && control.title === locator.title) {
          matchedControls.push(control);
        } else if (locator.tag && control.tag === locator.tag) {
          matchedControls.push(control);
        }
      }

      if (matchedControls.length === 0) {
        throw new Error("未找到匹配的内容控件 / No matching content control found");
      }

      const index = locator.index || 0;
      if (index >= matchedControls.length) {
        throw new Error(
          `内容控件索引 ${index} 超出范围 / Content control index ${index} out of range`
        );
      }

      return matchedControls[index].getRange();
    }

    default:
      throw new Error(`不支持的定位器类型 / Unsupported locator type`);
  }
}

/**
 * 替换单个图片 / Replace single image
 */
async function replaceSingleImage(
  _context: Word.RequestContext,
  image: Word.InlinePicture,
  newImageData?: string,
  properties?: ImageProperties
): Promise<void> {
  if (newImageData) {
    const range = image.getRange();
    const newImage = range.insertInlinePictureFromBase64(newImageData, "Replace");

    if (properties) {
      applyImageProperties(newImage, properties);
    }
  } else if (properties) {
    applyImageProperties(image, properties);
  }
}

/**
 * 替换图片 / Replace Image
 *
 * @param options - 替换选项 / Replace options
 * @returns Promise<ReplaceImageResult> 替换结果 / Replace result
 *
 * @remarks
 * 此函数提供统一的图片替换能力，支持四种定位方式：
 * 1. selection - 替换当前选中的图片
 * 2. index - 按索引替换图片
 * 3. search - 查找并替换匹配的图片
 * 4. range - 替换指定范围内的图片
 *
 * This function provides unified image replacement capability with four locator types:
 * 1. selection - Replace currently selected image
 * 2. index - Replace image by index
 * 3. search - Find and replace matching images
 * 4. range - Replace images in specific range
 *
 * @example
 * ```typescript
 * // 替换选中图片
 * await replaceImage({
 *   locator: { type: "selection" },
 *   newImageData: "base64...",
 *   properties: { width: 200, height: 150 }
 * });
 *
 * // 按索引替换
 * await replaceImage({
 *   locator: { type: "index", index: 0 },
 *   properties: { altText: "新的替代文本" }
 * });
 *
 * // 搜索并替换
 * await replaceImage({
 *   locator: {
 *     type: "search",
 *     searchOptions: { altText: "旧图片" }
 *   },
 *   newImageData: "base64...",
 *   replaceAll: true
 * });
 *
 * // 替换指定范围
 * await replaceImage({
 *   locator: {
 *     type: "range",
 *     rangeLocator: { type: "paragraph", startIndex: 0 }
 *   },
 *   properties: { width: 300 }
 * });
 * ```
 */
export async function replaceImage(options: ReplaceImageOptions): Promise<ReplaceImageResult> {
  const { locator, newImageData, properties, replaceAll = false } = options;

  if (!newImageData && !properties) {
    return {
      count: 0,
      success: false,
      error:
        "必须提供 newImageData 或 properties 中的至少一个 / Must provide at least one of newImageData or properties",
    };
  }

  try {
    return await Word.run(async (context) => {
      let count = 0;

      switch (locator.type) {
        case "selection": {
          const selection = context.document.getSelection();

          if (newImageData) {
            const newImage = selection.insertInlinePictureFromBase64(newImageData, "Replace");
            await context.sync();

            if (properties) {
              applyImageProperties(newImage, properties);
              await context.sync();
            }
            count = 1;
          } else if (properties) {
            // eslint-disable-next-line office-addins/no-navigational-load
            selection.load("inlinePictures");
            await context.sync();

            const images = selection.inlinePictures;
            images.load("items");
            await context.sync();

            if (images.items.length === 0) {
              return {
                count: 0,
                success: false,
                error: "选区中没有图片 / No images in selection",
              };
            }

            applyImageProperties(images.items[0], properties);
            await context.sync();
            count = 1;
          }

          break;
        }

        case "index": {
          const allImages = context.document.body.inlinePictures;
          allImages.load("items");
          await context.sync();

          if (locator.index >= allImages.items.length) {
            return {
              count: 0,
              success: false,
              error: `图片索引 ${locator.index} 超出范围 / Image index ${locator.index} out of range`,
            };
          }

          await replaceSingleImage(
            context,
            allImages.items[locator.index],
            newImageData,
            properties
          );
          await context.sync();

          count = 1;
          break;
        }

        case "search": {
          const allImages = context.document.body.inlinePictures;
          allImages.load("items");
          await context.sync();

          const matchedImages: Word.InlinePicture[] = [];

          for (const image of allImages.items) {
            if (await matchesSearchCriteria(context, image, locator.searchOptions)) {
              matchedImages.push(image);
            }
          }

          if (matchedImages.length === 0) {
            return {
              count: 0,
              success: false,
              error: "未找到匹配的图片 / No matching images found",
            };
          }

          const imagesToReplace = replaceAll ? matchedImages : [matchedImages[0]];

          for (let i = imagesToReplace.length - 1; i >= 0; i--) {
            await replaceSingleImage(context, imagesToReplace[i], newImageData, properties);
            count++;
          }
          await context.sync();

          break;
        }

        case "range": {
          const range = await getRangeByLocator(context, locator.rangeLocator);
          const images = range.inlinePictures;
          images.load("items");
          await context.sync();

          if (images.items.length === 0) {
            return {
              count: 0,
              success: false,
              error: "指定范围内没有图片 / No images in specified range",
            };
          }

          const imagesToReplace = replaceAll ? images.items : [images.items[0]];

          for (let i = imagesToReplace.length - 1; i >= 0; i--) {
            await replaceSingleImage(context, imagesToReplace[i], newImageData, properties);
            count++;
          }
          await context.sync();

          break;
        }

        default:
          return {
            count: 0,
            success: false,
            error: `不支持的定位器类型 / Unsupported locator type`,
          };
      }

      console.log(`成功替换 ${count} 张图片 / Successfully replaced ${count} image(s)`);
      return {
        count,
        success: true,
      };
    });
  } catch (error) {
    console.error("替换图片失败 / Failed to replace image:", error);
    return {
      count: 0,
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
