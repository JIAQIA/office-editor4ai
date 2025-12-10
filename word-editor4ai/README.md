
## Word 工具封装规划

### 核心设计理念

Word 与 PPT 的本质区别：
- **PPT**: 离散的块状结构（幻灯片 → 形状/文本框/图片）
- **Word**: 连续的文档流结构（文档 → 节 → 段落 → 文本/图片/表格）

因此，Word 工具应该围绕**文档流的层次结构**和**定位方式**来设计。

---

## 一、内容获取工具（Query Tools）

### 1. **文档结构获取** `documentStructure.ts`
```typescript
// 获取文档大纲结构（标题层级）
getDocumentOutline(): Promise<OutlineNode[]>

// 获取文档节信息（分节符、页眉页脚配置）
getDocumentSections(): Promise<SectionInfo[]>

// 获取文档统计信息
getDocumentStats(): Promise<DocumentStats>
```

### 2. **内容范围获取** `contentRange.ts`
```typescript
// 获取可见内容（已实现）
getVisibleContent(options): Promise<PageInfo[]>

// 获取指定页面内容
getPageContent(pageNumber: number, options): Promise<PageInfo>

// 获取选中内容
getSelectedContent(options): Promise<ContentInfo>

// 获取指定范围内容（通过书签、标题、段落索引等）
getRangeContent(locator: RangeLocator, options): Promise<ContentInfo>
```

### 3. **特殊区域获取** `specialRegions.ts`
```typescript
// 获取页眉内容
getHeaderContent(sectionIndex?: number): Promise<HeaderFooterInfo>

// 获取页脚内容
getFooterContent(sectionIndex?: number): Promise<HeaderFooterInfo>

// 获取文本框内容
getTextBoxes(): Promise<TextBoxInfo[]>

// 获取批注内容
getComments(): Promise<CommentInfo[]>
```

---

## 二、内容插入工具（Insertion Tools）

### 4. **文本插入** [textInsertion.ts](cci:7://file:///Users/jqq/WebstormProjects/office-editor4ai/ppt-editor4ai/src/ppt-tools/textInsertion.ts:0:0-0:0)
```typescript
// 在指定位置插入文本
insertText(text: string, location: InsertLocation, format?: TextFormat): Promise<void>

// 在选中位置插入文本
insertTextAtSelection(text: string, format?: TextFormat): Promise<void>

// 在文档末尾追加文本
appendText(text: string, format?: TextFormat): Promise<void>
```

### 5. **段落插入** `paragraphInsertion.ts`
```typescript
// 插入段落
insertParagraph(text: string, location: InsertLocation, style?: ParagraphStyle): Promise<void>

// 插入标题
insertHeading(text: string, level: 1-9, location: InsertLocation): Promise<void>

// 插入列表项
insertListItem(text: string, listType: 'bullet' | 'number', location: InsertLocation): Promise<void>
```

### 6. **对象插入** `objectInsertion.ts`
```typescript
// 插入图片（内联或浮动）
insertImage(imageData: string, location: InsertLocation, options?: ImageOptions): Promise<void>

// 插入表格
table(rows: number, cols: number, location: InsertLocation, data?: string[][]): Promise<void>

// 插入文本框
insertTextBox(text: string, location: InsertLocation, options?: TextBoxOptions): Promise<void>

// 插入形状
insertShape(shapeType: string, location: InsertLocation, options?: ShapeOptions): Promise<void>
```

### 7. **特殊元素插入** `specialInsertion.ts`
```typescript
// 插入分页符
insertPageBreak(location: InsertLocation): Promise<void>

// 插入分节符
insertSectionBreak(breakType: SectionBreakType, location: InsertLocation): Promise<void>

// 插入封面页
insertCoverPage(template?: string): Promise<void>

// 插入目录
insertTableOfContents(location: InsertLocation, options?: TOCOptions): Promise<void>

// 插入公式
insertEquation(latex: string, location: InsertLocation): Promise<void>
```

---

## 三、内容更新工具（Update Tools）

### 8. **文本更新** [textUpdate.ts](cci:7://file:///Users/jqq/WebstormProjects/office-editor4ai/ppt-editor4ai/src/ppt-tools/textUpdate.ts:0:0-0:0)
```typescript
// 更新选中文本
updateSelectedText(newText: string, format?: TextFormat): Promise<void>

// 查找并替换文本
findAndReplace(searchText: string, replaceText: string, options?: FindReplaceOptions): Promise<number>

// 更新指定范围文本
updateRangeText(locator: RangeLocator, newText: string, format?: TextFormat): Promise<void>
```

### 9. **格式更新** `formatUpdate.ts`
```typescript
// 更新文本格式
updateTextFormat(locator: RangeLocator, format: TextFormat): Promise<void>

// 更新段落格式
updateParagraphFormat(locator: RangeLocator, format: ParagraphFormat): Promise<void>

// 应用样式
applyStyle(locator: RangeLocator, styleName: string): Promise<void>
```

### 10. **表格更新** `tableUpdate.ts`
```typescript
// 更新表格单元格
updateTableCell(tableLocator: TableLocator, row: number, col: number, content: string, format?: CellFormat): Promise<void>

// 批量更新表格数据
updateTableData(tableLocator: TableLocator, data: string[][]): Promise<void>

// 更新表格样式
updateTableStyle(tableLocator: TableLocator, style: TableStyle): Promise<void>

// 插入/删除行列
insertTableRow(tableLocator: TableLocator, rowIndex: number, data?: string[]): Promise<void>
deleteTableRow(tableLocator: TableLocator, rowIndex: number): Promise<void>
insertTableColumn(tableLocator: TableLocator, colIndex: number): Promise<void>
deleteTableColumn(tableLocator: TableLocator, colIndex: number): Promise<void>
```

### 11. **图片更新** `imageUpdate.ts`
```typescript
// 替换图片
replaceImage(imageLocator: ImageLocator, newImageData: string): Promise<void>

// 更新图片属性
updateImageProperties(imageLocator: ImageLocator, properties: ImageProperties): Promise<void>

// 批量替换图片
replaceImages(replacements: ImageReplacement[]): Promise<void>
```

### 12. **页眉页脚更新** `headerFooterUpdate.ts`
```typescript
// 更新页眉
updateHeader(content: string, sectionIndex?: number, type?: 'primary' | 'first' | 'even'): Promise<void>

// 更新页脚
updateFooter(content: string, sectionIndex?: number, type?: 'primary' | 'first' | 'even'): Promise<void>

// 插入页码
insertPageNumber(location: 'header' | 'footer', format?: PageNumberFormat): Promise<void>
```

---

## 四、内容删除工具（Deletion Tools）

### 13. **内容删除** `contentDeletion.ts`
```typescript
// 删除选中内容
deleteSelection(): Promise<void>

// 删除指定范围
deleteRange(locator: RangeLocator): Promise<void>

// 删除表格
deleteTable(tableLocator: TableLocator): Promise<void>

// 删除图片
deleteImage(imageLocator: ImageLocator): Promise<void>

// 清空页眉/页脚
clearHeader(sectionIndex?: number): Promise<void>
clearFooter(sectionIndex?: number): Promise<void>
```

---

## 五、布局与截图工具（Layout & Screenshot Tools）

### 14. **布局信息** `layoutInfo.ts`
```typescript
// 获取页面布局信息（页边距、纸张大小等）
getPageLayout(sectionIndex?: number): Promise<PageLayout>

// 获取元素位置信息
getElementPosition(locator: ElementLocator): Promise<PositionInfo>

// 获取文档视觉结构（用于 AI 理解布局）
getDocumentLayout(): Promise<DocumentLayoutInfo>
```

### 15. **内容导出工具** `exportContent.ts` ✅
```typescript
// 导出 Word 文档内容为 OOXML、HTML 等格式
exportContent(options?: ExportContentOptions): Promise<ExportContentResult>

// 导出选项
interface ExportContentOptions {
  scope?: 'document' | 'selection' | 'visible';  // 导出范围，默认 'selection'
  format?: 'ooxml' | 'html' | 'pdf';             // 导出格式，默认 'ooxml'
}

// 导出结果
interface ExportContentResult {
  content: string;          // 导出的内容数据（OOXML/HTML 为文本，PDF 为 Base64）
  format: ExportFormat;     // 导出格式
  scope: ExportScope;       // 导出范围
  timestamp: number;        // 导出时间戳
  size: number;             // 内容大小（字节）
  mimeType: string;         // MIME 类型
}

// 支持的格式
// - OOXML: Office Open XML 格式，保留完整格式信息
// - HTML: HTML 格式，适合网页显示
// - PDF: 暂不可用（Word API 限制）
```

---

## 六、核心类型定义

### **定位器类型** `types/locators.ts`

```typescript
// 插入位置定位器
type InsertLocation = 
  | { type: 'selection' }  // 当前选中位置
  | { type: 'start' }      // 文档开始
  | { type: 'end' }        // 文档末尾
  | { type: 'bookmark', name: string }  // 书签位置
  | { type: 'heading', text: string }   // 标题位置
  | { type: 'paragraph', index: number } // 段落索引
  | { type: 'page', number: number, position: 'start' | 'end' } // 页面位置

// 范围定位器
type RangeLocator =
  | { type: 'selection' }
  | { type: 'bookmark', name: string }
  | { type: 'heading', text: string }
  | { type: 'paragraph', index: number }
  | { type: 'paragraphRange', start: number, end: number }
  | { type: 'search', text: string, occurrence?: number }

// 表格定位器
type TableLocator =
  | { type: 'index', index: number }  // 第几个表格
  | { type: 'bookmark', name: string }
  | { type: 'nearHeading', headingText: string }  // 某标题附近的表格
  | { type: 'selection' }  // 选中的表格

// 图片定位器
type ImageLocator =
  | { type: 'index', index: number }
  | { type: 'altText', text: string }
  | { type: 'selection' }
```

---

## 工具分组总结

| 分类 | 工具文件 | 核心功能 |
|------|---------|---------|
| **查询** | `documentStructure.ts` | 大纲、节、统计 |
| | `contentRange.ts` | 可见内容、页面内容、选中内容 |
| | `specialRegions.ts` | 页眉页脚、文本框、批注 |
| **插入** | [textInsertion.ts](cci:7://file:///Users/jqq/WebstormProjects/office-editor4ai/ppt-editor4ai/src/ppt-tools/textInsertion.ts:0:0-0:0) | 文本插入 |
| | `paragraphInsertion.ts` | 段落、标题、列表 |
| | `objectInsertion.ts` | 图片、表格、文本框、形状 |
| | `specialInsertion.ts` | 分页符、封面、目录、公式 |
| **更新** | [textUpdate.ts](cci:7://file:///Users/jqq/WebstormProjects/office-editor4ai/ppt-editor4ai/src/ppt-tools/textUpdate.ts:0:0-0:0) | 文本内容更新、查找替换 |
| | `formatUpdate.ts` | 文本格式、段落格式、样式 |
| | `tableUpdate.ts` | 表格数据、样式、行列操作 |
| | `imageUpdate.ts` | 图片替换、属性更新 |
| | `headerFooterUpdate.ts` | 页眉页脚、页码 |
| **删除** | `contentDeletion.ts` | 删除各类内容 |
| **布局** | `layoutInfo.ts` | 页面布局、元素位置 |
| | `screenshot.ts` | 页面截图、区域截图 |

---

## 与 PPT 工具的对比

| 维度 | PPT 工具 | Word 工具 |
|------|---------|----------|
| **定位方式** | 页码 + 元素索引/ID | 定位器系统（书签、标题、段落索引等） |
| **内容结构** | 离散块状（Slide → Shape） | 连续流式（Document → Section → Paragraph） |
| **特殊区域** | 母版、备注 | 页眉页脚、文本框、批注 |
| **布局概念** | 幻灯片布局模板 | 页面设置、节布局 |
