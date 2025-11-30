# Word 可见内容获取工具 - 实现文档

## 概述

已成功为 Word Add-in 实现了类似 PPT `elementsList` 的可见内容获取工具。该工具利用 Office.js 的 `Word.Pane.pagesEnclosingViewport` API，获取用户当前可见范围内的文档内容，避免一次性加载整个大文档造成的性能问题。

## 技术实现

### 核心 API

使用 Microsoft Office.js Word API：
- **`Word.Document.activeWindow`**: 获取活动窗口
- **`Word.Window.activePane`**: 获取活动窗格
- **`Word.Pane.pagesEnclosingViewport`**: 获取视口中的页面集合（API 要求集: WordApiDesktop 1.2）
- **`Word.Page.getRange()`**: 获取页面的 Range 对象
- **`Word.Range`**: 访问段落、表格、图片、内容控件等

### 支持的内容类型

1. **段落 (Paragraph)**
   - 文本内容
   - 样式、对齐方式
   - 缩进（首行、左、右）
   - 行距、段前段后间距
   - 列表项信息

2. **表格 (Table)**
   - 行数、列数
   - 单元格内容
   - 单元格宽度

3. **图片 (Image/InlinePicture)**
   - 尺寸（宽度、高度）
   - 替代文本
   - 超链接

4. **内容控件 (ContentControl)**
   - 标题、标签
   - 控件类型
   - 编辑权限
   - 占位符文本

## 文件结构

```
word-editor4ai/
├── src/
│   ├── word-tools/                          # 核心工具目录（新增）
│   │   ├── index.ts                         # 导出文件
│   │   ├── visibleContent.ts                # 核心逻辑实现
│   │   ├── README.md                        # 工具文档
│   │   ├── __tests__/                       # 测试目录
│   │   │   └── visibleContent.test.ts       # 单元测试
│   │   └── examples/                        # 示例目录
│   │       └── visibleContentExample.ts     # 使用示例
│   │
│   └── taskpane/
│       └── components/
│           ├── Sidebar.tsx                  # 已更新：添加新工具菜单项
│           ├── ToolsDebugPage.tsx           # 已更新：添加工具路由
│           └── tools/                       # 工具组件目录（新增）
│               └── VisibleContent.tsx       # UI 组件
```

## 主要功能

### 1. `getVisibleContent(options)`

获取用户当前可见范围的所有内容元素。

**参数选项：**
```typescript
interface GetVisibleContentOptions {
  includeText?: boolean;           // 是否包含文本内容，默认 true
  includeImages?: boolean;          // 是否包含图片信息，默认 true
  includeTables?: boolean;          // 是否包含表格信息，默认 true
  includeContentControls?: boolean; // 是否包含内容控件，默认 true
  detailedMetadata?: boolean;       // 是否包含详细元数据，默认 false
  maxTextLength?: number;           // 文本内容的最大长度，默认不限制
}
```

**返回值：**
```typescript
PageInfo[] // 页面信息数组
```

### 2. `getVisibleText()`

简化版本，仅返回可见内容的文本字符串。

### 3. `getVisibleContentStats()`

返回可见内容的统计信息：
- 页面数
- 元素总数
- 字符数
- 段落数
- 表格数
- 图片数
- 内容控件数

## UI 组件特性

`VisibleContent.tsx` 提供了完整的用户界面：

1. **选项配置面板**
   - 可切换包含的内容类型
   - 可启用/禁用详细元数据

2. **操作按钮**
   - 获取可见内容
   - 获取统计信息

3. **内容展示**
   - 按页面分组显示
   - 每个元素以卡片形式展示
   - 显示元素类型图标
   - 文本内容预览（超过 200 字符自动截断）
   - 详细元数据展示

4. **统计信息展示**
   - 网格布局显示各项统计数据
   - 清晰的数值和标签

## 与 PPT elementsList 的对比

| 特性 | PPT elementsList | Word visibleContent |
|------|------------------|---------------------|
| 获取范围 | 单个幻灯片 | 用户可见的多个页面 |
| 元素类型 | Shape、Placeholder | Paragraph、Table、Image、ContentControl |
| 位置信息 | left、top、width、height | 无（Word 元素流式布局）|
| 文本内容 | ✓ | ✓ |
| 元数据 | placeholderType、name | style、alignment、indent 等 |
| 页码选择 | ✓（可指定页码）| ✗（仅可见范围）|

## 使用场景

1. **AI 内容分析**
   - 只分析用户当前查看的内容
   - 提高响应速度，降低 API 调用成本

2. **实时翻译**
   - 翻译可见范围的文本
   - 避免全文档翻译的延迟

3. **内容摘要**
   - 为当前可见内容生成摘要
   - 快速理解文档当前部分

4. **格式检查**
   - 检查可见范围内的格式问题
   - 实时反馈

5. **内容导出**
   - 导出用户正在查看的部分内容
   - 支持按需导出

## 注意事项

### 平台限制

- `pagesEnclosingViewport` API 主要支持 **Word 桌面版**
- Web 版和移动版可能不支持此 API
- 需要 **WordApiDesktop 1.2** 或更高版本

### 性能优化建议

1. 使用 `maxTextLength` 限制文本长度
2. 根据需求选择性包含内容类型
3. 避免频繁调用，建议添加防抖/节流
4. 大文档建议关闭 `detailedMetadata`

### 错误处理

- 工具内置了完善的错误处理
- 单个元素失败不会影响其他元素
- 所有错误都会记录到控制台

### API 特殊说明

- **表格列数**: Word.Table API 没有直接的 `columnCount` 属性，需要通过 `table.columns.items.length` 获取
- **表格行数**: 使用 `table.rowCount` 属性直接获取

## 测试指南

### 手动测试步骤

1. 准备测试文档：
   - 包含多页内容（至少 3 页）
   - 包含段落、表格、图片
   - 可选：添加内容控件

2. 启动 Word Add-in：
   ```bash
   cd word-editor4ai
   npm run dev-server
   npm start
   ```

3. 在 Word 中打开测试文档

4. 导航到"查询元素类" > "可见内容获取"

5. 测试场景：
   - ✓ 默认选项获取可见内容
   - ✓ 只获取文本（关闭其他选项）
   - ✓ 启用详细元数据
   - ✓ 滚动到不同位置，验证内容变化
   - ✓ 获取统计信息

### 自动化测试

测试文件位于 `src/word-tools/__tests__/visibleContent.test.ts`

运行测试：
```bash
npm test
```

## 代码示例

### 基本用法

```typescript
import { getVisibleContent } from '../word-tools';

const pages = await getVisibleContent({
  includeText: true,
  includeImages: true,
  includeTables: true,
});

console.log(`找到 ${pages.length} 个可见页面`);
```

### 获取统计信息

```typescript
import { getVisibleContentStats } from '../word-tools';

const stats = await getVisibleContentStats();
console.log(`可见内容包含 ${stats.paragraphCount} 个段落`);
```

### 查找特定内容

```typescript
import { getVisibleContent } from '../word-tools';

const pages = await getVisibleContent({ includeText: true });

for (const page of pages) {
  for (const element of page.elements) {
    if (element.text?.includes('重要')) {
      console.log('找到重要内容:', element.text);
    }
  }
}
```

## 后续优化建议

1. **缓存机制**
   - 缓存已获取的页面内容
   - 减少重复 API 调用

2. **增量更新**
   - 监听文档变化事件
   - 只更新变化的部分

3. **导出功能**
   - 支持导出为 JSON、Markdown 等格式
   - 支持导出到剪贴板

4. **搜索功能**
   - 在可见内容中搜索关键词
   - 高亮显示匹配项

5. **过滤功能**
   - 按元素类型过滤
   - 按内容长度过滤

## 参考资料

- [Word JavaScript API 文档](https://learn.microsoft.com/en-us/javascript/api/word)
- [Word.Pane 类文档](https://learn.microsoft.com/en-us/javascript/api/word/word.pane)
- [Word.Page 类文档](https://learn.microsoft.com/en-us/javascript/api/word/word.page)
- [Office Add-ins 开发指南](https://learn.microsoft.com/en-us/office/dev/add-ins/)

## 版本历史

- **v1.0.0** (2025/11/30)
  - 初始版本发布
  - 实现核心功能
  - 完成 UI 组件
  - 添加测试和示例
