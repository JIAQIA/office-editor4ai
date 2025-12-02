/**
 * 文件名: headerFooterContent.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取文档页眉页脚内容的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

import type {
  HeaderFooterType,
  HeaderFooterContentItem,
  SectionHeaderFooterInfo,
  DocumentHeaderFooterInfo,
  GetHeaderFooterContentOptions,
  AnyContentElement,
  ParagraphElement,
  TableElement,
  InlinePictureElement,
  ContentControlElement,
} from "./types";

/**
 * 解析内容元素
 * Parse content elements
 */
async function parseContentElements(
  body: Word.Body,
  context: Word.RequestContext
): Promise<AnyContentElement[]> {
  const elements: AnyContentElement[] = [];

  try {
    // 获取所有段落 / Get all paragraphs
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    for (let i = 0; i < paragraphs.items.length; i++) {
      const para = paragraphs.items[i];
      para.load([
        "text",
        "style",
        "alignment",
        "firstLineIndent",
        "leftIndent",
        "rightIndent",
        "lineSpacing",
        "spaceAfter",
        "spaceBefore",
        "isListItem",
      ]);
    }

    await context.sync();

    // 处理段落 / Process paragraphs
    for (let i = 0; i < paragraphs.items.length; i++) {
      const para = paragraphs.items[i];

      const element: ParagraphElement = {
        id: `para-${i}`,
        type: "Paragraph",
        text: para.text || "",
        style: para.style || undefined,
        alignment: para.alignment || undefined,
        firstLineIndent: para.firstLineIndent || undefined,
        leftIndent: para.leftIndent || undefined,
        rightIndent: para.rightIndent || undefined,
        lineSpacing: para.lineSpacing || undefined,
        spaceAfter: para.spaceAfter || undefined,
        spaceBefore: para.spaceBefore || undefined,
        isListItem: para.isListItem || false,
      };

      elements.push(element);

      // 检查段落中的内联图片 / Check inline pictures in paragraph
      try {
        const inlinePictures = para.inlinePictures;
        inlinePictures.load("items");
        await context.sync();

        for (let j = 0; j < inlinePictures.items.length; j++) {
          const pic = inlinePictures.items[j];
          pic.load(["width", "height", "altTextDescription", "hyperlink"]);
        }

        await context.sync();

        for (let j = 0; j < inlinePictures.items.length; j++) {
          const pic = inlinePictures.items[j];
          const picElement: InlinePictureElement = {
            id: `para-${i}-pic-${j}`,
            type: "InlinePicture",
            width: pic.width || undefined,
            height: pic.height || undefined,
            altText: pic.altTextDescription || undefined,
            hyperlink: pic.hyperlink || undefined,
          };
          elements.push(picElement);
        }
      } catch (error) {
        // 忽略内联图片获取错误 / Ignore inline picture errors
      }
    }

    // 获取表格 / Get tables
    try {
      const tables = body.tables;
      tables.load("items");
      await context.sync();

      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        table.load("rowCount");
        table.columns.load("items");
      }

      await context.sync();

      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        const tableElement: TableElement = {
          id: `table-${i}`,
          type: "Table",
          rowCount: table.rowCount || 0,
          columnCount: table.columns.items.length || 0,
        };
        elements.push(tableElement);
      }
    } catch (error) {
      // 忽略表格获取错误 / Ignore table errors
    }

    // 获取内容控件 / Get content controls
    try {
      const contentControls = body.contentControls;
      contentControls.load("items");
      await context.sync();

      for (let i = 0; i < contentControls.items.length; i++) {
        const cc = contentControls.items[i];
        cc.load(["text", "title", "tag", "type", "cannotDelete", "cannotEdit", "placeholderText"]);
      }

      await context.sync();

      for (let i = 0; i < contentControls.items.length; i++) {
        const cc = contentControls.items[i];
        const ccElement: ContentControlElement = {
          id: `cc-${i}`,
          type: "ContentControl",
          text: cc.text || "",
          title: cc.title || undefined,
          tag: cc.tag || undefined,
          controlType: cc.type || undefined,
          cannotDelete: cc.cannotDelete || false,
          cannotEdit: cc.cannotEdit || false,
          placeholderText: cc.placeholderText || undefined,
        };
        elements.push(ccElement);
      }
    } catch (error) {
      // 忽略内容控件获取错误 / Ignore content control errors
    }
  } catch (error) {
    console.warn("解析内容元素失败:", error);
  }

  return elements;
}

/**
 * 获取单个页眉或页脚的内容信息
 * Get single header or footer content information
 */
async function getHeaderFooterContentItem(
  body: Word.Body,
  type: HeaderFooterType,
  includeElements: boolean,
  context: Word.RequestContext
): Promise<HeaderFooterContentItem> {
  const item: HeaderFooterContentItem = {
    type,
    exists: false,
  };

  try {
    // 加载 body 的 text 属性 / Load body's text property
    body.load("text");
    await context.sync();

    // 检查是否有内容 / Check if there is content
    if (body.text && body.text.trim().length > 0) {
      item.exists = true;
      item.text = body.text;

      // 如果需要详细元素，则解析内容 / Parse content if detailed elements needed
      if (includeElements) {
        item.elements = await parseContentElements(body, context);
      }
    }
  } catch (error) {
    // 如果获取失败，说明该页眉页脚不存在或为空
    // If retrieval fails, the header/footer doesn't exist or is empty
    item.exists = false;
    console.warn(`获取页眉页脚内容失败 (type: ${type}):`, error);
  }

  return item;
}

/**
 * 获取单个节的页眉页脚信息
 * Get single section's header footer information
 */
async function getSectionHeaderFooterInfo(
  section: Word.Section,
  sectionIndex: number,
  includeElements: boolean,
  context: Word.RequestContext
): Promise<SectionHeaderFooterInfo> {
  const info: SectionHeaderFooterInfo = {
    sectionIndex,
    headers: [],
    footers: [],
    differentFirstPage: false,
    differentOddAndEven: false,
  };

  try {
    // 获取页面设置信息 / Get page setup info
    const pageSetup = section.body.parentSection.pageSetup;
    pageSetup.load(["differentFirstPageHeaderFooter", "oddAndEvenPagesHeaderFooter"]);
    await context.sync();

    info.differentFirstPage = pageSetup.differentFirstPageHeaderFooter || false;
    info.differentOddAndEven = pageSetup.oddAndEvenPagesHeaderFooter || false;

    // 获取页眉 / Get headers
    const headerFirst = section.getHeader(Word.HeaderFooterType.firstPage);
    const headerPrimary = section.getHeader(Word.HeaderFooterType.primary);
    const headerEven = section.getHeader(Word.HeaderFooterType.evenPages);

    // 获取页脚 / Get footers
    const footerFirst = section.getFooter(Word.HeaderFooterType.firstPage);
    const footerPrimary = section.getFooter(Word.HeaderFooterType.primary);
    const footerEven = section.getFooter(Word.HeaderFooterType.evenPages);

    // 批量获取页眉内容 / Batch get header content
    info.headers.push(
      await getHeaderFooterContentItem(
        headerFirst,
        "firstPage" as HeaderFooterType,
        includeElements,
        context
      )
    );
    info.headers.push(
      await getHeaderFooterContentItem(
        headerPrimary,
        "oddPages" as HeaderFooterType,
        includeElements,
        context
      )
    );
    info.headers.push(
      await getHeaderFooterContentItem(
        headerEven,
        "evenPages" as HeaderFooterType,
        includeElements,
        context
      )
    );

    // 批量获取页脚内容 / Batch get footer content
    info.footers.push(
      await getHeaderFooterContentItem(
        footerFirst,
        "firstPage" as HeaderFooterType,
        includeElements,
        context
      )
    );
    info.footers.push(
      await getHeaderFooterContentItem(
        footerPrimary,
        "oddPages" as HeaderFooterType,
        includeElements,
        context
      )
    );
    info.footers.push(
      await getHeaderFooterContentItem(
        footerEven,
        "evenPages" as HeaderFooterType,
        includeElements,
        context
      )
    );
  } catch (error) {
    console.warn(`获取节 ${sectionIndex} 的页眉页脚信息失败:`, error);
  }

  return info;
}

/**
 * 获取文档页眉页脚内容
 * Get document header footer content
 *
 * @param options - 获取选项 / Get options
 * @returns 文档页眉页脚信息 / Document header footer information
 *
 * @example
 * ```typescript
 * // 获取所有节的页眉页脚内容
 * const result = await getHeaderFooterContent({
 *   includeElements: true,
 *   includeMetadata: true
 * });
 *
 * console.log(`文档共有 ${result.totalSections} 个节`);
 * console.log(`共有 ${result.metadata?.totalHeaders} 个页眉`);
 * console.log(`共有 ${result.metadata?.totalFooters} 个页脚`);
 *
 * // 获取指定节的页眉页脚内容
 * const sectionResult = await getHeaderFooterContent({
 *   sectionIndex: 0,
 *   includeElements: false
 * });
 * ```
 */
export async function getHeaderFooterContent(
  options: GetHeaderFooterContentOptions = {}
): Promise<DocumentHeaderFooterInfo> {
  const { sectionIndex, includeElements = false, includeMetadata = true } = options;

  return Word.run(async (context) => {
    try {
      // 获取文档的所有节 / Get all sections in the document
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      const result: DocumentHeaderFooterInfo = {
        sections: [],
        totalSections: sections.items.length,
      };

      // 确定要处理的节 / Determine which sections to process
      const sectionsToProcess: Array<{ section: Word.Section; index: number }> = [];

      if (sectionIndex !== undefined) {
        // 处理指定节 / Process specific section
        if (sectionIndex < 0 || sectionIndex >= sections.items.length) {
          throw new Error(`节索引 ${sectionIndex} 超出范围 (0-${sections.items.length - 1})`);
        }
        sectionsToProcess.push({
          section: sections.items[sectionIndex],
          index: sectionIndex,
        });
      } else {
        // 处理所有节 / Process all sections
        for (let i = 0; i < sections.items.length; i++) {
          sectionsToProcess.push({
            section: sections.items[i],
            index: i,
          });
        }
      }

      // 批量获取所有节的页眉页脚信息 / Batch get all sections' header footer info
      for (const { section, index } of sectionsToProcess) {
        const sectionInfo = await getSectionHeaderFooterInfo(
          section,
          index,
          includeElements,
          context
        );
        result.sections.push(sectionInfo);
      }

      // 计算元数据 / Calculate metadata
      if (includeMetadata) {
        let hasAnyHeader = false;
        let hasAnyFooter = false;
        let totalHeaders = 0;
        let totalFooters = 0;

        for (const section of result.sections) {
          for (const header of section.headers) {
            if (header.exists) {
              hasAnyHeader = true;
              totalHeaders++;
            }
          }
          for (const footer of section.footers) {
            if (footer.exists) {
              hasAnyFooter = true;
              totalFooters++;
            }
          }
        }

        result.metadata = {
          hasAnyHeader,
          hasAnyFooter,
          totalHeaders,
          totalFooters,
        };
      }

      return result;
    } catch (error) {
      console.error("获取页眉页脚内容失败:", error);
      throw new Error(
        `获取页眉页脚内容失败: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  });
}
