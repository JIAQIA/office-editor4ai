# replaceText 工具使用示例

`replaceText` 是一个统一的文本替换工具，整合了三种常见的文本替换场景：
1. 替换当前选中的文本
2. 查找并替换匹配的文本
3. 替换指定范围的文本

## 核心优势

- **统一接口**：一个工具解决三种需求，降低学习成本
- **灵活定位**：支持选区、搜索、范围（书签、标题、段落、节、内容控件）
- **格式控制**：可选的文本格式应用
- **批量替换**：搜索模式支持替换所有匹配项

## 使用场景

### 场景 1: 替换选中文本

```typescript
import { replaceText } from './word-tools';

// 替换当前选中的文本
await replaceText({
  locator: { type: "selection" },
  newText: "新文本内容"
});

// 替换选中文本并应用格式
await replaceText({
  locator: { type: "selection" },
  newText: "重要提示",
  format: {
    bold: true,
    color: "#FF0000",
    fontSize: 16
  }
});
```

### 场景 2: 查找并替换

```typescript
// 查找并替换第一个匹配项
await replaceText({
  locator: {
    type: "search",
    searchText: "旧公司名称",
    searchOptions: {
      matchCase: true,
      matchWholeWord: true
    }
  },
  newText: "新公司名称"
});

// 查找并替换所有匹配项
await replaceText({
  locator: {
    type: "search",
    searchText: "TODO",
    searchOptions: {
      matchCase: false
    }
  },
  newText: "✓ 已完成",
  replaceAll: true,
  format: {
    color: "#00AA00",
    strikeThrough: true
  }
});
```

### 场景 3: 替换指定范围

#### 3.1 替换书签位置的文本

```typescript
// 替换书签内容（常用于模板填充）
await replaceText({
  locator: {
    type: "range",
    rangeLocator: {
      type: "bookmark",
      name: "签名位置"
    }
  },
  newText: "张三",
  format: {
    fontName: "楷体",
    fontSize: 14
  }
});
```

#### 3.2 替换标题文本

```typescript
// 替换第一个一级标题
await replaceText({
  locator: {
    type: "range",
    rangeLocator: {
      type: "heading",
      level: 1,
      index: 0
    }
  },
  newText: "新的文档标题"
});

// 替换包含特定文本的标题
await replaceText({
  locator: {
    type: "range",
    rangeLocator: {
      type: "heading",
      text: "第一章",
      level: 1
    }
  },
  newText: "第一章：引言"
});
```

#### 3.3 替换段落文本

```typescript
// 替换第一个段落
await replaceText({
  locator: {
    type: "range",
    rangeLocator: {
      type: "paragraph",
      startIndex: 0
    }
  },
  newText: "这是新的第一段内容"
});

// 替换段落范围（第 5-10 段）
await replaceText({
  locator: {
    type: "range",
    rangeLocator: {
      type: "paragraph",
      startIndex: 4,
      endIndex: 9
    }
  },
  newText: "这段文本将替换第5到第10段的所有内容"
});
```

#### 3.4 替换节内容

```typescript
// 替换第一节的所有内容
await replaceText({
  locator: {
    type: "range",
    rangeLocator: {
      type: "section",
      index: 0
    }
  },
  newText: "新的节内容"
});
```

#### 3.5 替换内容控件

```typescript
// 通过标题查找内容控件
await replaceText({
  locator: {
    type: "range",
    rangeLocator: {
      type: "contentControl",
      title: "客户姓名"
    }
  },
  newText: "李四"
});

// 通过标签查找内容控件
await replaceText({
  locator: {
    type: "range",
    rangeLocator: {
      type: "contentControl",
      tag: "customer_name",
      index: 0
    }
  },
  newText: "王五"
});
```

## 返回值

```typescript
interface ReplaceResult {
  count: number;      // 替换的数量
  success: boolean;   // 是否成功
  error?: string;     // 错误信息（如果有）
}

// 示例
const result = await replaceText({
  locator: {
    type: "search",
    searchText: "错误"
  },
  newText: "正确",
  replaceAll: true
});

console.log(`成功替换 ${result.count} 处文本`);
// 输出: 成功替换 5 处文本
```

## 参数说明

### ReplaceTextOptions

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| locator | ReplaceTextLocator | ✓ | 定位方式 |
| newText | string | ✓ | 新文本内容 |
| format | TextFormat | ✗ | 文本格式（可选） |
| replaceAll | boolean | ✗ | 是否替换所有匹配项（仅 search 模式） |

### TextFormat

| 参数 | 类型 | 说明 |
|------|------|------|
| fontName | string | 字体名称 |
| fontSize | number | 字号（磅） |
| bold | boolean | 是否加粗 |
| italic | boolean | 是否斜体 |
| color | string | 字体颜色（十六进制） |
| highlightColor | string | 高亮颜色 |
| strikeThrough | boolean | 删除线 |
| superscript | boolean | 上标 |
| subscript | boolean | 下标 |

## 最佳实践

1. **选区模式**：适合用户手动选择后的快速替换
2. **搜索模式**：适合批量替换相同文本，建议使用 `matchWholeWord` 避免误替换
3. **范围模式**：适合模板填充、结构化文档更新

## 错误处理

```typescript
const result = await replaceText({
  locator: {
    type: "range",
    rangeLocator: {
      type: "bookmark",
      name: "不存在的书签"
    }
  },
  newText: "测试"
});

if (!result.success) {
  console.error(`替换失败: ${result.error}`);
  // 输出: 替换失败: 书签 "不存在的书签" 不存在
}
```

## 与其他工具的对比

| 工具 | 适用场景 | 定位方式 |
|------|----------|----------|
| replaceSelection | 替换选中内容，支持图片 | 仅选区 |
| replaceText | 统一的文本替换 | 选区/搜索/范围 |
| updateSelectedText | 已废弃，使用 replaceText | - |
| findAndReplace | 已废弃，使用 replaceText | - |
