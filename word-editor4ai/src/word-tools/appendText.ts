/**
 * 文件名: appendText.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 在文档末尾追加文本的工具核心逻辑
 */

/* global Word, console */

import type { TextFormat, ImageData } from "./types";

export interface AppendOptions {
  text?: string;
  format?: TextFormat;
  images?: ImageData[];
  file?: File; // 新增file属性
}

/**
 * 在文档末尾追加内容
 * Append content to the end of document
 */
export async function appendText(options: AppendOptions): Promise<void> {
  const { text, format, images, file } = options;

  // 验证参数
  if (!text && (!images || images.length === 0) && !file) {
    throw new Error("必须提供文本、图片或文件");
  }

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const endRange = body.getRange("End");
      
      let insertedRange: Word.Range | undefined;
      if (text) {
        insertedRange = endRange.insertText(text, "End");
        
        if (format) {
          const font = insertedRange.font;
          if (format.fontName) font.name = format.fontName;
          if (format.fontSize) font.size = format.fontSize;
          if (format.bold) font.bold = format.bold;
          if (format.italic) font.italic = format.italic;
          if (format.color) font.color = format.color;
        }
      }
      
      if (images && images.length > 0) {
        let insertPosition = insertedRange ? insertedRange.getRange("End") : endRange;
        
        for (const imageData of images) {
          try {
            let base64Data = imageData.base64;
            if (base64Data.includes(",")) {
              base64Data = base64Data.split(",")[1];
            }
            
            const inlinePicture = insertPosition.insertInlinePictureFromBase64(base64Data, "End");
            
            if (imageData.width) inlinePicture.width = imageData.width;
            if (imageData.height) inlinePicture.height = imageData.height;
            if (imageData.altText) inlinePicture.altTextTitle = imageData.altText;
            
            insertPosition = inlinePicture.getRange("End");
          } catch (error) {
            console.warn("插入图片失败:", error);
          }
        }
      }
      
      if (file) {
        // 上传文件逻辑
        // Upload file logic
        // ...
      }
      
      await context.sync();
    });
  } catch (error) {
    console.error("追加内容失败:", error);
    throw error;
  }
}
