# Word Tools - 可见内容获取工具

## 概述

这个工具集提供了获取 Word 文档中用户当前可见范围内容的功能，避免一次性加载整个大文档造成的性能问题。

## 核心功能

### 1. 获取可见内容 (`getVisibleContent`)

获取用户当前视口中可见的所有页面及其内容元素。

**支持的内容类型：**
- **段落 (Paragraph)**: 包含文本、样式、对齐方式、缩进、行距等信息
- **表格 (Table)**: 包含行数、列数、单元格内容等信息
- **图片 (Image/InlinePicture)**: 包含尺寸、替代文本、超链接等信息
- **内容控件 (ContentControl)**: 包含标题、标签、类型等信息

**使用示例：**

```typescript
import { getVisibleContent } from '../word-tools';

// 获取可见内容
const pages = await getVisibleContent({
  includeText: true,           // 包含文本内容
  includeImages: true,          // 包含图片信息
  includeTables: true,          // 包含表格信息
  includeContentControls: true, // 包含内容控件
  detailedMetadata: true,       // 包含详细元数据
  maxTextLength: 500           // 限制文本长度
});

// 遍历可见页面
for (const page of pages) {
  console.log(`页面 ${page.index + 1}:`);
  console.log(`元素数量: ${page.elements.length}`);
  
  // 遍历页面元素
  for (const element of page.elements) {
    console.log(`类型: ${element.type}`);
    if (element.text) {
      console.log(`文本: ${element.text}`);
    }
  }
}
```

### 2. 获取可见文本 (`getVisibleText`)

简化版本，仅获取可见内容的文本。

```typescript
import { getVisibleText } from '../word-tools';

const text = await getVisibleText();
console.log('可见文本:', text);
```

### 3. 获取统计信息 (`getVisibleContentStats`)

获取可见内容的统计数据。

```typescript
import { getVisibleContentStats } from '../word-tools';

const stats = await getVisibleContentStats();
console.log('统计信息:', {
  pageCount: stats.pageCount,           // 可见页数
  elementCount: stats.elementCount,     // 元素总数
  characterCount: stats.characterCount, // 字符数
  paragraphCount: stats.paragraphCount, // 段落数
  tableCount: stats.tableCount,         // 表格数
  imageCount: stats.imageCount,         // 图片数
  contentControlCount: stats.contentControlCount // 控件数
});
```

## API 要求

- **Word API 要求集**: `WordApiDesktop 1.2` 或更高版本
- **平台支持**: 主要支持 Word 桌面版

## 数据结构

### PageInfo

```typescript
interface PageInfo {
  index: number;                    // 页面索引
  elements: AnyContentElement[];    // 页面元素列表
  text?: string;                    // 页面完整文本
}
```

### ParagraphElement

```typescript
interface ParagraphElement {
  id: string;
  type: "Paragraph";
  text?: string;
  style?: string;              // 样式名称
  alignment?: string;          // 对齐方式
  firstLineIndent?: number;    // 首行缩进
  leftIndent?: number;         // 左缩进
  rightIndent?: number;        // 右缩进
  lineSpacing?: number;        // 行距
  spaceAfter?: number;         // 段后间距
  spaceBefore?: number;        // 段前间距
  isListItem?: boolean;        // 是否为列表项
  listLevel?: number;          // 列表级别
}
```

### TableElement

```typescript
interface TableElement {
  id: string;
  type: "Table";
  rowCount?: number;           // 行数（通过 table.rowCount 获取）
  columnCount?: number;        // 列数（通过 table.columns.items.length 获取）
  cells?: TableCellInfo[][];   // 单元格内容
}
```

**注意**: Word.Table API 没有直接的 `columnCount` 属性，需要通过 `table.columns` 集合的长度来获取列数。

### ImageElement / InlinePictureElement

```typescript
interface ImageElement {
  id: string;
  type: "Image" | "InlinePicture";
  width?: number;              // 宽度
  height?: number;             // 高度
  altText?: string;            // 替代文本
  hyperlink?: string;          // 超链接
}
```

### ContentControlElement

```typescript
interface ContentControlElement {
  id: string;
  type: "ContentControl";
  text?: string;
  title?: string;              // 标题
  tag?: string;                // 标签
  controlType?: string;        // 控件类型
  cannotDelete?: boolean;      // 是否可删除
  cannotEdit?: boolean;        // 是否可编辑
  placeholderText?: string;    // 占位符文本
}
```

## UI 组件

UI 组件位于 `src/taskpane/components/tools/VisibleContent.tsx`，提供了：

- **获取选项配置**: 可选择包含哪些类型的内容
- **可见内容展示**: 以卡片形式展示每个页面的元素
- **统计信息展示**: 显示可见内容的统计数据
- **详细元数据**: 可选显示元素的详细属性

## 注意事项

1. **性能优化**: 使用 `maxTextLength` 限制文本长度，避免处理超大文本
2. **错误处理**: 工具会捕获并记录错误，不会因单个元素失败而中断整个流程
3. **平台限制**: `pagesEnclosingViewport` API 主要支持桌面版 Word
4. **异步操作**: 所有 API 调用都是异步的，需要使用 `await` 或 `.then()`

## 使用场景

- **AI 内容分析**: 只分析用户当前查看的内容，提高响应速度
- **实时翻译**: 翻译可见范围的文本，避免全文档翻译的延迟
- **内容摘要**: 为当前可见内容生成摘要
- **格式检查**: 检查可见范围内的格式问题
- **内容导出**: 导出用户正在查看的部分内容
