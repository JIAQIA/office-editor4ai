/**
 * 文件名: types.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: Word 工具集公共类型定义
 */

/**
 * 内容元素基础信息 / Base Content Element Info
 */
export interface ContentElement {
  /** 元素唯一标识 / Element unique ID */
  id: string;
  /** 元素类型 / Element type */
  type: ContentElementType;
  /** 文本内容 / Text content */
  text?: string;
  /** 元数据 / Metadata */
  metadata?: Record<string, unknown>;
}

/**
 * 段落元素 / Paragraph Element
 */
export interface ParagraphElement extends ContentElement {
  type: "Paragraph";
  /** 样式名称 / Style name */
  style?: string;
  /** 对齐方式 / Alignment */
  alignment?: string;
  /** 首行缩进（磅）/ First line indent in points */
  firstLineIndent?: number;
  /** 左缩进（磅）/ Left indent in points */
  leftIndent?: number;
  /** 右缩进（磅）/ Right indent in points */
  rightIndent?: number;
  /** 行间距 / Line spacing */
  lineSpacing?: number;
  /** 段后间距（磅）/ Space after in points */
  spaceAfter?: number;
  /** 段前间距（磅）/ Space before in points */
  spaceBefore?: number;
  /** 是否为列表项 / Is list item */
  isListItem?: boolean;
  /** 列表级别 / List level */
  listLevel?: number;
}

/**
 * 表格单元格信息 / Table Cell Info
 */
export interface TableCellInfo {
  /** 单元格文本 / Cell text */
  text: string;
  /** 行索引 / Row index */
  rowIndex: number;
  /** 列索引 / Column index */
  columnIndex: number;
  /** 单元格宽度（磅）/ Cell width in points */
  width?: number;
}

/**
 * 表格元素 / Table Element
 */
export interface TableElement extends ContentElement {
  type: "Table";
  /** 行数 / Row count */
  rowCount?: number;
  /** 列数 / Column count */
  columnCount?: number;
  /** 单元格数据 / Cell data */
  cells?: TableCellInfo[][];
}

/**
 * 图片元素 / Image Element
 */
export interface ImageElement extends ContentElement {
  type: "Image";
  /** 宽度（磅）/ Width in points */
  width?: number;
  /** 高度（磅）/ Height in points */
  height?: number;
  /** 替代文本 / Alt text */
  altText?: string;
  /** 超链接 / Hyperlink */
  hyperlink?: string;
  /** Base64 编码（可选）/ Base64 encoding (optional) */
  base64?: string;
}

/**
 * 内联图片元素 / Inline Picture Element
 */
export interface InlinePictureElement extends ContentElement {
  type: "InlinePicture";
  /** 宽度（磅）/ Width in points */
  width?: number;
  /** 高度（磅）/ Height in points */
  height?: number;
  /** 替代文本 / Alt text */
  altText?: string;
  /** 超链接 / Hyperlink */
  hyperlink?: string;
}

/**
 * 内容控件元素 / Content Control Element
 */
export interface ContentControlElement extends ContentElement {
  type: "ContentControl";
  /** 标题 / Title */
  title?: string;
  /** 标签 / Tag */
  tag?: string;
  /** 控件类型 / Control type */
  controlType?: string;
  /** 是否不可删除 / Cannot delete */
  cannotDelete?: boolean;
  /** 是否不可编辑 / Cannot edit */
  cannotEdit?: boolean;
  /** 占位符文本 / Placeholder text */
  placeholderText?: string;
}

/**
 * 文本框元素 / Text Box Element
 */
export interface TextBoxElement extends ContentElement {
  type: "TextBox";
  /** 名称 / Name */
  name?: string;
  /** 宽度（磅）/ Width in points */
  width?: number;
  /** 高度（磅）/ Height in points */
  height?: number;
  /** 左边距（磅）/ Left position in points */
  left?: number;
  /** 上边距（磅）/ Top position in points */
  top?: number;
  /** 旋转角度 / Rotation in degrees */
  rotation?: number;
  /** 是否可见 / Visible */
  visible?: boolean;
  /** 是否锁定纵横比 / Lock aspect ratio */
  lockAspectRatio?: boolean;
  /** 文本框内的段落列表 / Paragraphs in text box */
  paragraphs?: ParagraphElement[];
}

/**
 * 内容元素类型 / Content Element Type
 */
export type ContentElementType =
  | "Paragraph"
  | "Table"
  | "Image"
  | "ContentControl"
  | "InlinePicture"
  | "TextBox"
  | "Unknown";

/**
 * 所有内容元素的联合类型 / Union type of all content elements
 */
export type AnyContentElement =
  | ParagraphElement
  | TableElement
  | ImageElement
  | ContentControlElement
  | InlinePictureElement
  | TextBoxElement
  | ContentElement;

/**
 * 页面信息 / Page Info
 */
export interface PageInfo {
  /** 页面索引（从0开始）/ Page index (0-based) */
  index: number;
  /** 页面内的内容元素 / Content elements in the page */
  elements: AnyContentElement[];
  /** 页面的完整文本 / Complete text of the page */
  text?: string;
}

/**
 * 范围定位器类型 / Range Locator Type
 */
export type RangeLocatorType = "bookmark" | "heading" | "paragraph" | "section" | "contentControl";

/**
 * 书签定位器 / Bookmark Locator
 */
export interface BookmarkLocator {
  type: "bookmark";
  /** 书签名称 / Bookmark name */
  name: string;
}

/**
 * 标题定位器 / Heading Locator
 */
export interface HeadingLocator {
  type: "heading";
  /** 标题文本（精确匹配或部分匹配）/ Heading text (exact or partial match) */
  text?: string;
  /** 标题级别（1-9）/ Heading level (1-9) */
  level?: number;
  /** 标题索引（从0开始，同级别中的第几个）/ Heading index (0-based, which one among same level) */
  index?: number;
}

/**
 * 段落定位器 / Paragraph Locator
 */
export interface ParagraphLocator {
  type: "paragraph";
  /** 段落索引（从0开始）/ Paragraph index (0-based) */
  startIndex: number;
  /** 结束段落索引（可选，不指定则只获取单个段落）/ End paragraph index (optional, if not specified, only get single paragraph) */
  endIndex?: number;
}

/**
 * 节定位器 / Section Locator
 */
export interface SectionLocator {
  type: "section";
  /** 节索引（从0开始）/ Section index (0-based) */
  index: number;
}

/**
 * 内容控件定位器 / Content Control Locator
 */
export interface ContentControlLocator {
  type: "contentControl";
  /** 控件标题 / Control title */
  title?: string;
  /** 控件标签 / Control tag */
  tag?: string;
  /** 控件索引（从0开始）/ Control index (0-based) */
  index?: number;
}

/**
 * 范围定位器联合类型 / Range Locator Union Type
 */
export type RangeLocator =
  | BookmarkLocator
  | HeadingLocator
  | ParagraphLocator
  | SectionLocator
  | ContentControlLocator;

/**
 * 内容信息 / Content Info
 */
export interface ContentInfo {
  /** 文本内容 / Text content */
  text: string;
  /** 内容的元素列表 / List of elements in content */
  elements: AnyContentElement[];
  /** 内容的元数据 / Content metadata */
  metadata?: {
    /** 是否为空 / Is empty */
    isEmpty: boolean;
    /** 字符数 / Character count */
    characterCount: number;
    /** 段落数 / Paragraph count */
    paragraphCount: number;
    /** 表格数 / Table count */
    tableCount: number;
    /** 图片数 / Image count */
    imageCount: number;
    /** 定位器类型（仅用于范围内容）/ Locator type (only for range content) */
    locatorType?: string;
  };
}

/**
 * 获取内容的选项 / Get Content Options
 */
export interface GetContentOptions {
  /** 是否包含文本内容，默认为 true / Include text content, default true */
  includeText?: boolean;
  /** 是否包含图片信息，默认为 true / Include image info, default true */
  includeImages?: boolean;
  /** 是否包含表格信息，默认为 true / Include table info, default true */
  includeTables?: boolean;
  /** 是否包含内容控件，默认为 true / Include content controls, default true */
  includeContentControls?: boolean;
  /** 是否包含详细的元数据，默认为 false / Include detailed metadata, default false */
  detailedMetadata?: boolean;
  /** 文本内容的最大长度，默认不限制 / Max text length, default unlimited */
  maxTextLength?: number;
}

/**
 * 页眉页脚类型 / Header Footer Type
 */
export enum HeaderFooterType {
  /** 首页 / First page */
  FirstPage = "firstPage",
  /** 奇数页 / Odd pages */
  OddPages = "oddPages",
  /** 偶数页 / Even pages */
  EvenPages = "evenPages",
}

/**
 * 单个页眉或页脚的内容信息 / Single Header or Footer Content Info
 */
export interface HeaderFooterContentItem {
  /** 类型 / Type */
  type: HeaderFooterType;
  /** 是否存在 / Exists */
  exists: boolean;
  /** 文本内容 / Text content */
  text?: string;
  /** 内容元素列表 / Content elements list */
  elements?: AnyContentElement[];
  /** 是否链接到上一节 / Link to previous section */
  linkToPrevious?: boolean;
}

/**
 * 单个节的页眉页脚信息 / Single Section Header Footer Info
 */
export interface SectionHeaderFooterInfo {
  /** 节索引（从0开始）/ Section index (0-based) */
  sectionIndex: number;
  /** 页眉列表 / Headers list */
  headers: HeaderFooterContentItem[];
  /** 页脚列表 / Footers list */
  footers: HeaderFooterContentItem[];
  /** 是否首页不同 / Different first page */
  differentFirstPage: boolean;
  /** 是否奇偶页不同 / Different odd and even pages */
  differentOddAndEven: boolean;
}

/**
 * 文档所有页眉页脚信息 / Document All Headers Footers Info
 */
export interface DocumentHeaderFooterInfo {
  /** 所有节的页眉页脚信息 / All sections header footer info */
  sections: SectionHeaderFooterInfo[];
  /** 文档总节数 / Total section count */
  totalSections: number;
  /** 元数据 / Metadata */
  metadata?: {
    /** 是否有任何页眉 / Has any header */
    hasAnyHeader: boolean;
    /** 是否有任何页脚 / Has any footer */
    hasAnyFooter: boolean;
    /** 页眉总数 / Total header count */
    totalHeaders: number;
    /** 页脚总数 / Total footer count */
    totalFooters: number;
  };
}

/**
 * 获取页眉页脚内容的选项 / Get Header Footer Content Options
 */
export interface GetHeaderFooterContentOptions {
  /** 指定节索引（可选，不指定则获取所有节）/ Specific section index (optional, get all if not specified) */
  sectionIndex?: number;
  /** 是否包含详细内容元素，默认为 false / Include detailed content elements, default false */
  includeElements?: boolean;
  /** 是否包含元数据统计，默认为 true / Include metadata statistics, default true */
  includeMetadata?: boolean;
}

/**
 * 文本框信息 / Text Box Info
 */
export interface TextBoxInfo {
  /** 文本框唯一标识 / Text box unique ID */
  id: string;
  /** 名称 / Name */
  name?: string;
  /** 文本内容 / Text content */
  text?: string;
  /** 宽度（磅）/ Width in points */
  width?: number;
  /** 高度（磅）/ Height in points */
  height?: number;
  /** 左边距（磅）/ Left position in points */
  left?: number;
  /** 上边距（磅）/ Top position in points */
  top?: number;
  /** 旋转角度 / Rotation in degrees */
  rotation?: number;
  /** 是否可见 / Visible */
  visible?: boolean;
  /** 是否锁定纵横比 / Lock aspect ratio */
  lockAspectRatio?: boolean;
  /** 文本框内的段落列表 / Paragraphs in text box */
  paragraphs?: ParagraphElement[];
}

/**
 * 获取文本框内容的选项 / Get Text Box Content Options
 */
export interface GetTextBoxOptions {
  /** 是否包含文本内容，默认为 true / Include text content, default true */
  includeText?: boolean;
  /** 是否包含段落详情，默认为 false / Include paragraph details, default false */
  includeParagraphs?: boolean;
  /** 是否包含详细的元数据，默认为 false / Include detailed metadata, default false */
  detailedMetadata?: boolean;
  /** 文本内容的最大长度，默认不限制 / Max text length, default unlimited */
  maxTextLength?: number;
}

/**
 * 批注引用范围的位置信息 / Comment Range Location Info
 *
 * @remarks
 * 注意：start、end 等是导航属性，在 Office Add-ins 中加载这些属性会降低性能
 * 但为了提供精确的位置信息，这些属性是必需的
 * Note: start, end etc. are navigation properties, loading them in Office Add-ins will slow down performance
 * But these properties are necessary to provide precise location information
 */
export interface CommentRangeLocation {
  /** 范围样式 / Range style */
  style?: string;
  /** 范围所在段落索引（如果可获取）/ Paragraph index if available */
  paragraphIndex?: number;
  /** 范围起始位置（字符偏移）/ Range start position (character offset) */
  start?: number;
  /** 范围结束位置（字符偏移）/ Range end position (character offset) */
  end?: number;
  /** 故事类型（MainText, Footnotes, Headers 等）/ Story type (MainText, Footnotes, Headers, etc.) */
  storyType?: string;
  /** 文本哈希值，用于识别重复引用 / Text hash for identifying duplicate references */
  textHash?: string;
  /** 文本长度（字符数）/ Text length in characters */
  textLength?: number;
  /** 是否为列表项 / Is list item */
  isListItem?: boolean;
  /** 列表级别（如果是列表项）/ List level if is list item */
  listLevel?: number;
  /** 字体名称 / Font name */
  font?: string;
  /** 字体大小（磅）/ Font size in points */
  fontSize?: number;
  /** 是否加粗 / Is bold */
  isBold?: boolean;
  /** 是否斜体 / Is italic */
  isItalic?: boolean;
  /** 是否有下划线 / Is underlined */
  isUnderlined?: boolean;
  /** 高亮颜色 / Highlight color */
  highlightColor?: string;
}

/**
 * 批注回复信息 / Comment Reply Info
 */
export interface CommentReplyInfo {
  /** 回复唯一标识 / Reply unique ID */
  id: string;
  /** 回复内容 / Reply content */
  content: string;
  /** 回复作者 / Reply author */
  authorName?: string;
  /** 回复作者邮箱 / Reply author email */
  authorEmail?: string;
  /** 回复创建时间 / Reply creation date */
  creationDate?: Date;
}

/**
 * 批注信息 / Comment Info
 */
export interface CommentInfo {
  /** 批注唯一标识 / Comment unique ID */
  id: string;
  /** 批注内容 / Comment content */
  content: string;
  /** 批注作者 / Comment author */
  authorName?: string;
  /** 批注作者邮箱 / Comment author email */
  authorEmail?: string;
  /** 批注创建时间 / Comment creation date */
  creationDate?: Date;
  /** 批注是否已解决 / Comment is resolved */
  resolved?: boolean;
  /** 批注关联的文本 / Comment associated text */
  associatedText?: string;
  /** 批注引用范围的位置信息 / Comment range location info */
  rangeLocation?: CommentRangeLocation;
  /** 批注回复列表 / Comment replies */
  replies?: CommentReplyInfo[];
}

/**
 * 获取批注的选项 / Get Comments Options
 */
export interface GetCommentsOptions {
  /** 是否包含已解决的批注，默认为 true / Include resolved comments, default true */
  includeResolved?: boolean;
  /** 是否包含批注回复，默认为 true / Include comment replies, default true */
  includeReplies?: boolean;
  /** 是否包含关联文本，默认为 true / Include associated text, default true */
  includeAssociatedText?: boolean;
  /** 是否包含详细的元数据，默认为 false / Include detailed metadata, default false */
  detailedMetadata?: boolean;
  /** 文本内容的最大长度，默认不限制 / Max text length, default unlimited */
  maxTextLength?: number;
}

/**
 * 文本格式 / Text Format
 */
export interface TextFormat {
  /** 字体名称 / Font name */
  fontName?: string;
  /** 字体大小（磅）/ Font size in points */
  fontSize?: number;
  /** 是否加粗 / Is bold */
  bold?: boolean;
  /** 是否斜体 / Is italic */
  italic?: boolean;
  /** 下划线类型 / Underline type */
  underline?: Word.UnderlineType | string;
  /** 字体颜色 / Font color */
  color?: string;
  /** 高亮颜色 / Highlight color */
  highlightColor?: string;
  /** 删除线 / Strikethrough */
  strikeThrough?: boolean;
  /** 上标 / Superscript */
  superscript?: boolean;
  /** 下标 / Subscript */
  subscript?: boolean;
}

/**
 * 图片数据 / Image Data
 */
export interface ImageData {
  /** Base64 编码的图片数据 / Base64 encoded image data */
  base64: string;
  /** 图片宽度（磅），可选 / Image width in points, optional */
  width?: number;
  /** 图片高度（磅），可选 / Image height in points, optional */
  height?: number;
  /** 替代文本，可选 / Alt text, optional */
  altText?: string;
}

/**
 * 替换选中内容的选项 / Replace Selection Options
 */
export interface ReplaceSelectionOptions {
  /** 文本内容 / Text content */
  text?: string;
  /** 文本格式（可选，不指定则使用原格式）/ Text format (optional, use original format if not specified) */
  format?: TextFormat;
  /** 图片列表（可选，按顺序插入）/ Image list (optional, insert in order) */
  images?: ImageData[];
  /** 是否替换选中内容（默认为 true）/ Replace selection (default true) */
  replaceSelection?: boolean;
}
