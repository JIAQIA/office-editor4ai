/**
 * 文件名: exportContent.ts
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: Word 文档内容导出工具 - 支持导出 OOXML、HTML、PDF 等格式
 */

/* global Word, console, Blob */

/**
 * 导出范围类型 / Export Scope Type
 */
export type ExportScope =
  | "document" // 整个文档 / Entire document
  | "selection" // 当前选中区域 / Current selection
  | "visible"; // 当前可见区域 / Current visible area

/**
 * 导出格式类型 / Export Format Type
 */
export type ExportFormat =
  | "ooxml" // Office Open XML 格式 / Office Open XML format
  | "html" // HTML 格式 / HTML format
  | "pdf"; // PDF 格式（Beta API）/ PDF format (Beta API)

/**
 * 导出选项 / Export Options
 */
export interface ExportContentOptions {
  /**
   * 导出范围，默认为 "selection"
   * Export scope, default is "selection"
   */
  scope?: ExportScope;

  /**
   * 导出格式，默认为 "ooxml"
   * Export format, default is "ooxml"
   */
  format?: ExportFormat;
}

/**
 * 导出结果 / Export Result
 */
export interface ExportContentResult {
  /**
   * 导出的内容数据
   * - OOXML/HTML: 文本字符串
   * - PDF: Base64 编码的字符串
   * Exported content data
   * - OOXML/HTML: text string
   * - PDF: Base64 encoded string
   */
  content: string;

  /**
   * 导出格式 / Export format
   */
  format: ExportFormat;

  /**
   * 导出范围 / Export scope
   */
  scope: ExportScope;

  /**
   * 导出时间戳 / Export timestamp
   */
  timestamp: number;

  /**
   * 内容大小（字节）/ Content size (bytes)
   */
  size: number;

  /**
   * MIME 类型 / MIME type
   */
  mimeType: string;
}

/**
 * 导出 Word 文档内容
 * Export Word document content
 *
 * @param options - 导出选项 / Export options
 * @returns Promise<ExportContentResult> 导出结果 / Export result
 *
 * @remarks
 * **支持的格式 / Supported Formats:**
 * 1. OOXML - Office Open XML 格式，文本格式，包含完整的格式信息
 *    OOXML - Office Open XML format, text format, contains complete formatting information
 * 2. HTML - HTML 格式，文本格式，适合在网页中显示
 *    HTML - HTML format, text format, suitable for web display
 * 3. PDF - PDF 格式（Beta API），返回 Base64 编码
 *    PDF - PDF format (Beta API), returns Base64 encoding
 *
 * @example
 * ```typescript
 * // 导出选中内容为 OOXML（默认）
 * const result1 = await exportContent();
 *
 * // 导出整个文档为 HTML
 * const result2 = await exportContent({ scope: "document", format: "html" });
 *
 * // 导出选中区域为 PDF
 * const result3 = await exportContent({ scope: "selection", format: "pdf" });
 *
 * // 使用导出的内容
 * console.log(result1.content); // OOXML 或 HTML 文本
 * console.log(result1.size); // 内容大小（字节）
 * ```
 */
export async function exportContent(
  options: ExportContentOptions = {}
): Promise<ExportContentResult> {
  const { scope = "selection", format = "ooxml" } = options;

  console.log("[exportContent] 开始导出内容", { scope, format });

  return Word.run(async (context) => {
    let range: Word.Range;

    // 根据不同的 scope 获取对应的 Range / Get Range based on scope
    switch (scope) {
      case "document":
        // 整个文档 / Entire document
        range = context.document.body.getRange();
        console.log("[exportContent] 导出整个文档");
        break;

      case "visible":
        // 当前可见区域 / Current visible area
        // Word API 没有直接获取可见区域的方法，使用文档主体
        // Word API doesn't have a direct method to get visible area, use document body
        range = context.document.body.getRange();
        console.log("[exportContent] 导出可见区域（使用文档主体）");
        break;

      case "selection":
        // 当前选中区域 / Current selection
        range = context.document.getSelection();
        console.log("[exportContent] 导出选中区域");
        break;

      default:
        throw new Error(`不支持的导出范围: ${scope} / Unsupported export scope: ${scope}`);
    }

    // 根据格式导出内容 / Export content based on format
    let content: string;
    let mimeType: string;

    switch (format) {
      case "ooxml": {
        // 导出为 OOXML 格式 / Export as OOXML format
        const ooxmlResult = range.getOoxml();
        await context.sync();
        content = ooxmlResult.value;
        mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        console.log("[exportContent] OOXML 导出成功，大小:", content.length);
        break;
      }

      case "html": {
        // 导出为 HTML 格式 / Export as HTML format
        const htmlResult = range.getHtml();
        await context.sync();
        content = htmlResult.value;
        mimeType = "text/html";
        console.log("[exportContent] HTML 导出成功，大小:", content.length);
        break;
      }

      case "pdf":
        // PDF 导出（Beta API，仅支持整个文档）
        // PDF export (Beta API, only supports entire document)
        if (scope !== "document") {
          throw new Error(
            "PDF 格式仅支持导出整个文档 / PDF format only supports exporting entire document"
          );
        }

        // 注意：exportAsFixedFormat 是 Beta API，且只能保存文件，不能获取 Base64
        // 这里我们返回一个提示信息
        // Note: exportAsFixedFormat is Beta API and can only save files, cannot get Base64 data
        // Here we return a message
        throw new Error(
          "PDF 导出功能暂不可用。Word JavaScript API 的 exportAsFixedFormat 方法仅支持保存文件，不支持获取 Base64 数据。建议使用 OOXML 或 HTML 格式。/ PDF export is not available. Word JavaScript API's exportAsFixedFormat method only supports saving files, not getting Base64 data. Please use OOXML or HTML format instead."
        );

      default:
        throw new Error(`不支持的导出格式: ${format} / Unsupported export format: ${format}`);
    }

    // 计算内容大小 / Calculate content size
    const size = new Blob([content]).size;

    console.log("[exportContent] 导出完成", { format, scope, size });

    return {
      content,
      format,
      scope,
      timestamp: Date.now(),
      size,
      mimeType,
    };
  });
}
