/**
 * 文件名: documentSections.ts
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取文档节信息（分节符、页眉页脚配置）的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

import { HeaderFooterType } from "./types";

/**
 * 页眉页脚信息 / Header Footer Info
 */
export interface HeaderFooterInfo {
  /** 类型 / Type */
  type: HeaderFooterType;
  /** 是否存在 / Exists */
  exists: boolean;
  /** 内容文本 / Content text */
  text?: string;
  /** 是否链接到上一节 / Link to previous section */
  linkToPrevious?: boolean;
}

/**
 * 节信息 / Section Info
 * 表示文档中的一个节（Section）
 */
export interface SectionInfo {
  /** 节索引（从0开始）/ Section index (0-based) */
  index: number;
  /** 页眉信息 / Header information */
  headers: HeaderFooterInfo[];
  /** 页脚信息 / Footer information */
  footers: HeaderFooterInfo[];
  /** 页面设置 / Page setup */
  pageSetup: {
    /** 页面宽度（磅）/ Page width in points */
    pageWidth: number;
    /** 页面高度（磅）/ Page height in points */
    pageHeight: number;
    /** 上边距（磅）/ Top margin in points */
    topMargin: number;
    /** 下边距（磅）/ Bottom margin in points */
    bottomMargin: number;
    /** 左边距（磅）/ Left margin in points */
    leftMargin: number;
    /** 右边距（磅）/ Right margin in points */
    rightMargin: number;
    /** 页面方向 / Page orientation */
    orientation: "portrait" | "landscape";
  };
  /** 节类型 / Section type */
  sectionType: "continuous" | "nextPage" | "oddPage" | "evenPage" | "nextColumn";
  /** 是否首页不同 / Different first page */
  differentFirstPage: boolean;
  /** 是否奇偶页不同 / Different odd and even pages */
  differentOddAndEven: boolean;
  /** 列数 / Number of columns */
  columnCount: number;
  /** 列间距（磅）/ Column spacing in points */
  columnSpacing?: number;
}

/**
 * 获取文档节信息选项 / Get Document Sections Options
 */
export interface GetDocumentSectionsOptions {
  /** 是否包含页眉页脚内容 / Include header footer content */
  includeContent?: boolean;
  /** 是否包含页面设置详情 / Include page setup details */
  includePageSetup?: boolean;
}

/**
 * 将 Word.SectionStart 转换为字符串
 * Convert Word.SectionStart to string
 */
function convertSectionStart(sectionStart: Word.SectionStart | string): SectionInfo["sectionType"] {
  if (typeof sectionStart === "string") {
    const lowerCase = sectionStart.toLowerCase();
    if (lowerCase === "continuous") return "continuous";
    if (lowerCase === "newcolumn") return "nextColumn";
    if (lowerCase === "newpage") return "nextPage";
    if (lowerCase === "oddpage") return "oddPage";
    if (lowerCase === "evenpage") return "evenPage";
  }
  return "nextPage";
}

/**
 * 将 Word.PageOrientation 转换为字符串
 * Convert Word.PageOrientation to string
 */
function convertPageOrientation(
  orientation: Word.PageOrientation | string
): "portrait" | "landscape" {
  if (typeof orientation === "string") {
    return orientation.toLowerCase() === "landscape" ? "landscape" : "portrait";
  }
  return "portrait";
}

/**
 * 获取页眉或页脚信息
 * Get header or footer information
 */
async function getHeaderFooterInfo(
  body: Word.Body,
  type: HeaderFooterType,
  includeContent: boolean,
  context: Word.RequestContext
): Promise<HeaderFooterInfo> {
  const info: HeaderFooterInfo = {
    type,
    exists: false,
  };

  try {
    // 加载 body 的 text 属性
    // Load body's text property
    body.load("text");
    await context.sync();

    // 检查是否有内容
    // Check if there is content
    if (body.text && body.text.trim().length > 0) {
      info.exists = true;

      if (includeContent) {
        info.text = body.text;
      }
    }
  } catch {
    // 如果获取失败，说明该页眉页脚不存在或为空
    // If retrieval fails, the header/footer doesn't exist or is empty
    info.exists = false;
  }

  return info;
}

/**
 * 获取文档节信息
 * Get document sections information
 *
 * @param options - 获取选项 / Get options
 * @returns 文档节信息列表 / Document sections information list
 *
 * @example
 * ```typescript
 * // 获取所有节信息，包含页眉页脚内容
 * const sections = await getDocumentSections({
 *   includeContent: true,
 *   includePageSetup: true
 * });
 *
 * console.log(`文档共有 ${sections.length} 个节`);
 * sections.forEach(section => {
 *   console.log(`节 ${section.index + 1}:`);
 *   console.log(`  分节类型: ${section.sectionType}`);
 *   console.log(`  页面方向: ${section.pageSetup.orientation}`);
 *   console.log(`  首页不同: ${section.differentFirstPage}`);
 * });
 * ```
 */
export async function getDocumentSections(
  options: GetDocumentSectionsOptions = {}
): Promise<SectionInfo[]> {
  const { includeContent = false, includePageSetup = true } = options;

  return Word.run(async (context) => {
    try {
      // 获取文档的所有节
      // Get all sections in the document
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      const sectionInfoList: SectionInfo[] = [];

      // 批量准备所有节的页眉页脚和页面设置对象
      // Batch prepare all section headers, footers and page setup objects
      const sectionData: Array<{
        section: Word.Section;
        headerPrimary: Word.Body;
        headerFirst: Word.Body;
        headerEven: Word.Body;
        footerPrimary: Word.Body;
        footerFirst: Word.Body;
        footerEven: Word.Body;
        pageSetup?: Word.Section;
        setup?: Word.PageSetup;
      }> = [];

      // 第一步：批量获取所有节的页眉页脚对象
      // Step 1: Batch get all section header and footer objects
      for (let i = 0; i < sections.items.length; i++) {
        const section = sections.items[i];

        // 加载页眉页脚 Body 对象
        // Load header and footer Body objects
        // 注意：Word.HeaderFooterType 有三个值：primary, firstPage, evenPages
        // Note: Word.HeaderFooterType has three values: primary, firstPage, evenPages
        const headerPrimary = section.getHeader(Word.HeaderFooterType.primary);
        const headerFirst = section.getHeader(Word.HeaderFooterType.firstPage);
        const headerEven = section.getHeader(Word.HeaderFooterType.evenPages);

        const footerPrimary = section.getFooter(Word.HeaderFooterType.primary);
        const footerFirst = section.getFooter(Word.HeaderFooterType.firstPage);
        const footerEven = section.getFooter(Word.HeaderFooterType.evenPages);

        const data: (typeof sectionData)[0] = {
          section,
          headerPrimary,
          headerFirst,
          headerEven,
          footerPrimary,
          footerFirst,
          footerEven,
        };

        // 如果需要页面设置，也批量准备
        // If page setup is needed, also prepare in batch
        if (includePageSetup) {
          try {
            const ps = section.body.parentSection;
            ps.load("pageSetup");
            data.pageSetup = ps;
          } catch (error) {
            console.warn(`准备节 ${i} 的页面设置失败:`, error);
          }
        }

        sectionData.push(data);
      }

      // 第二步：统一同步一次，获取所有页眉页脚对象和页面设置
      // Step 2: Single sync to get all header/footer objects and page setup
      await context.sync();

      // 第三步：批量加载页面设置的详细属性
      // Step 3: Batch load detailed page setup properties
      if (includePageSetup) {
        for (const data of sectionData) {
          if (data.pageSetup) {
            try {
              const setup = data.pageSetup.pageSetup;
              setup.load([
                "pageWidth",
                "pageHeight",
                "topMargin",
                "bottomMargin",
                "leftMargin",
                "rightMargin",
                "orientation",
                "sectionStart",
                "differentFirstPageHeaderFooter",
                "oddAndEvenPagesHeaderFooter",
              ]);
              data.setup = setup;
            } catch (error) {
              console.warn(`加载页面设置属性失败:`, error);
            }
          }
        }

        // 统一同步一次，获取所有页面设置属性
        // Single sync to get all page setup properties
        await context.sync();
      }

      // 第四步：处理每个节的数据
      // Step 4: Process each section's data
      for (let i = 0; i < sectionData.length; i++) {
        const data = sectionData[i];

        // 构建节信息
        // Build section information
        const sectionInfo: SectionInfo = {
          index: i,
          headers: [],
          footers: [],
          pageSetup: {
            pageWidth: 612, // 默认 8.5 英寸 * 72 点/英寸
            pageHeight: 792, // 默认 11 英寸 * 72 点/英寸
            topMargin: 72,
            bottomMargin: 72,
            leftMargin: 72,
            rightMargin: 72,
            orientation: "portrait",
          },
          sectionType: "nextPage",
          differentFirstPage: false,
          differentOddAndEven: false,
          columnCount: 1,
        };

        // 获取页眉信息
        // Get header information
        try {
          sectionInfo.headers.push(
            await getHeaderFooterInfo(
              data.headerFirst,
              HeaderFooterType.FirstPage,
              includeContent,
              context
            )
          );
          sectionInfo.headers.push(
            await getHeaderFooterInfo(
              data.headerPrimary,
              HeaderFooterType.OddPages,
              includeContent,
              context
            )
          );
          sectionInfo.headers.push(
            await getHeaderFooterInfo(
              data.headerEven,
              HeaderFooterType.EvenPages,
              includeContent,
              context
            )
          );
        } catch (error) {
          console.warn(`获取节 ${i} 的页眉信息失败:`, error);
        }

        // 获取页脚信息
        // Get footer information
        try {
          sectionInfo.footers.push(
            await getHeaderFooterInfo(
              data.footerFirst,
              HeaderFooterType.FirstPage,
              includeContent,
              context
            )
          );
          sectionInfo.footers.push(
            await getHeaderFooterInfo(
              data.footerPrimary,
              HeaderFooterType.OddPages,
              includeContent,
              context
            )
          );
          sectionInfo.footers.push(
            await getHeaderFooterInfo(
              data.footerEven,
              HeaderFooterType.EvenPages,
              includeContent,
              context
            )
          );
        } catch (error) {
          console.warn(`获取节 ${i} 的页脚信息失败:`, error);
        }

        // 获取页面设置详情（如果需要）
        // Get page setup details (if needed)
        if (includePageSetup && data.setup) {
          try {
            const setup = data.setup;

            sectionInfo.pageSetup = {
              pageWidth: setup.pageWidth || 612,
              pageHeight: setup.pageHeight || 792,
              topMargin: setup.topMargin || 72,
              bottomMargin: setup.bottomMargin || 72,
              leftMargin: setup.leftMargin || 72,
              rightMargin: setup.rightMargin || 72,
              orientation: convertPageOrientation(setup.orientation),
            };

            sectionInfo.sectionType = convertSectionStart(setup.sectionStart);
            sectionInfo.differentFirstPage = setup.differentFirstPageHeaderFooter || false;
            sectionInfo.differentOddAndEven = setup.oddAndEvenPagesHeaderFooter || false;
          } catch (error) {
            console.warn(`获取节 ${i} 的页面设置失败:`, error);
          }
        }

        sectionInfoList.push(sectionInfo);
      }

      return sectionInfoList;
    } catch (error) {
      console.error("获取文档节信息失败:", error);
      throw new Error(
        `获取文档节信息失败: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  });
}
