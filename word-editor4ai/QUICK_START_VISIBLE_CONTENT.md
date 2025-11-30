# Word 可见内容获取工具 - 快速开始

## 🚀 快速启动

### 1. 启动开发服务器

```bash
cd word-editor4ai
npm run dev-server
```

### 2. 启动 Word Add-in

在另一个终端窗口：

```bash
npm start
```

### 3. 在 Word 中使用

1. Word 会自动打开并加载 Add-in
2. 点击侧边栏的"查询元素类"菜单
3. 选择"可见内容获取"工具
4. 配置获取选项并点击"获取可见内容"按钮

## 📋 功能演示

### 场景 1: 获取当前可见内容

1. 在 Word 中打开一个多页文档
2. 滚动到你想分析的位置
3. 在 Add-in 中点击"获取可见内容"
4. 查看返回的页面和元素信息

### 场景 2: 获取统计信息

1. 点击"获取统计信息"按钮
2. 查看可见范围内的：
   - 页面数
   - 段落数
   - 表格数
   - 图片数
   - 字符数等

### 场景 3: 自定义获取选项

1. 取消勾选"包含图片信息"（如果不需要图片）
2. 勾选"详细元数据"（获取更多信息）
3. 点击"获取可见内容"
4. 查看定制化的结果

## 💻 代码使用示例

### 在你的代码中使用

```typescript
import { getVisibleContent, getVisibleContentStats } from './word-tools';

// 获取可见内容
async function analyzeVisibleContent() {
  const pages = await getVisibleContent({
    includeText: true,
    includeImages: true,
    includeTables: true,
    maxTextLength: 500
  });
  
  console.log(`找到 ${pages.length} 个可见页面`);
  
  for (const page of pages) {
    console.log(`页面 ${page.index + 1}:`);
    console.log(`- ${page.elements.length} 个元素`);
    console.log(`- ${page.text?.length || 0} 个字符`);
  }
}

// 获取统计信息
async function getStats() {
  const stats = await getVisibleContentStats();
  console.log('统计信息:', stats);
}
```

## 🎯 常见使用场景

### 1. AI 内容分析

```typescript
import { getVisibleText } from './word-tools';

async function analyzeWithAI() {
  // 只获取可见文本，发送给 AI 分析
  const text = await getVisibleText();
  
  // 调用 AI API
  const analysis = await callAIAPI(text);
  
  return analysis;
}
```

### 2. 实时翻译

```typescript
import { getVisibleContent } from './word-tools';

async function translateVisible() {
  const pages = await getVisibleContent({
    includeText: true,
    includeImages: false,
    includeTables: false
  });
  
  for (const page of pages) {
    for (const element of page.elements) {
      if (element.text) {
        const translated = await translate(element.text);
        console.log('翻译:', translated);
      }
    }
  }
}
```

### 3. 内容摘要

```typescript
import { getVisibleContent } from './word-tools';

async function summarizeVisible() {
  const pages = await getVisibleContent({
    includeText: true,
    maxTextLength: 200 // 限制长度
  });
  
  const allText = pages
    .map(page => page.text)
    .join('\n\n');
  
  const summary = await generateSummary(allText);
  return summary;
}
```

## 🔧 配置选项说明

| 选项 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `includeText` | boolean | true | 是否包含文本内容 |
| `includeImages` | boolean | true | 是否包含图片信息 |
| `includeTables` | boolean | true | 是否包含表格信息 |
| `includeContentControls` | boolean | true | 是否包含内容控件 |
| `detailedMetadata` | boolean | false | 是否包含详细元数据 |
| `maxTextLength` | number | undefined | 文本内容的最大长度 |

## 📊 返回数据结构

### PageInfo

```typescript
{
  index: 0,              // 页面索引
  elements: [...],       // 元素数组
  text: "页面文本..."    // 页面完整文本
}
```

### 元素类型

- **Paragraph**: 段落（包含文本、样式、对齐等）
- **Table**: 表格（包含行列数、单元格内容）
- **Image/InlinePicture**: 图片（包含尺寸、替代文本）
- **ContentControl**: 内容控件（包含标题、标签、类型）

## ⚠️ 注意事项

1. **平台限制**: 此功能主要支持 Word 桌面版（需要 WordApiDesktop 1.2+）
2. **性能优化**: 大文档建议使用 `maxTextLength` 限制文本长度
3. **错误处理**: 工具内置错误处理，单个元素失败不影响其他元素

## 🐛 故障排除

### 问题 1: "未检测到可见内容"

**解决方案:**
- 确保文档已打开且有内容
- 确保使用的是 Word 桌面版
- 检查 Word 版本是否支持 WordApiDesktop 1.2

### 问题 2: 获取的内容不完整

**解决方案:**
- 检查获取选项是否正确配置
- 确认所有需要的选项都已启用
- 查看控制台是否有错误信息

### 问题 3: 性能较慢

**解决方案:**
- 使用 `maxTextLength` 限制文本长度
- 关闭不需要的选项（如 `detailedMetadata`）
- 减少获取的内容类型

## 📚 更多资源

- [完整文档](./docs/WORD_VISIBLE_CONTENT_TOOL.md)
- [API 文档](./src/word-tools/README.md)
- [使用示例](./src/word-tools/examples/visibleContentExample.ts)
- [测试文件](./src/word-tools/__tests__/visibleContent.test.ts)

## 🤝 反馈与支持

如有问题或建议，请联系开发团队或提交 Issue。
