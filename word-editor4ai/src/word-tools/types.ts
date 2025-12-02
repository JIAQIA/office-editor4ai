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
 * 内容元素类型 / Content Element Type
 */
export type ContentElementType =
  | "Paragraph"
  | "Table"
  | "Image"
  | "ContentControl"
  | "InlinePicture"
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
