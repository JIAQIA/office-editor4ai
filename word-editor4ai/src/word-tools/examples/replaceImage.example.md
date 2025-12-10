# replaceImage 工具使用示例

## 功能概述

`replaceImage` 是一个统一的图片替换工具，支持四种定位方式和灵活的操作组合。

## 核心特性

- ✅ 支持四种定位方式：选区、索引、搜索、范围
- ✅ 可替换图片内容（Base64）
- ✅ 可更新图片属性（尺寸、替代文本、超链接等）
- ✅ 支持批量替换
- ✅ 统一的错误处理和结果返回

## 使用示例

### 1. 替换选中的图片

```typescript
import { replaceImage } from "./replaceImage";

// 替换选中图片的内容
await replaceImage({
  locator: { type: "selection" },
  newImageData: "iVBORw0KGgo...", // 纯 Base64 字符串，不包含 data:image/png;base64, 前缀
  properties: {
    width: 200,
    height: 150,
    altText: "新图片"
  }
});

// 仅更新选中图片的属性
await replaceImage({
  locator: { type: "selection" },
  properties: {
    altText: "更新的替代文本",
    hyperlink: "https://example.com"
  }
});
```

### 2. 按索引替换图片

```typescript
// 替换文档中第一张图片
await replaceImage({
  locator: { type: "index", index: 0 },
  newImageData: "iVBORw0KGgo...", // 纯 Base64 字符串
  properties: {
    lockAspectRatio: true
  }
});

// 仅更新第三张图片的尺寸
await replaceImage({
  locator: { type: "index", index: 2 },
  properties: {
    width: 300,
    height: 200
  }
});
```

### 3. 搜索并替换图片

```typescript
// 按替代文本搜索并替换
await replaceImage({
  locator: {
    type: "search",
    searchOptions: {
      altText: "旧图片"
    }
  },
  newImageData: "iVBORw0KGgo...", // 纯 Base64 字符串
  replaceAll: true // 替换所有匹配的图片
});

// 按尺寸范围搜索并更新属性
await replaceImage({
  locator: {
    type: "search",
    searchOptions: {
      minWidth: 100,
      maxWidth: 300,
      minHeight: 100,
      maxHeight: 300
    }
  },
  properties: {
    width: 200,
    height: 200,
    lockAspectRatio: true
  },
  replaceAll: true
});

// 仅替换第一个匹配的图片
await replaceImage({
  locator: {
    type: "search",
    searchOptions: {
      altText: "logo"
    }
  },
  newImageData: "iVBORw0KGgo...", // 纯 Base64 字符串
  replaceAll: false // 默认值，仅替换第一个
});
```

### 4. 替换指定范围内的图片

```typescript
// 替换第一段中的所有图片
await replaceImage({
  locator: {
    type: "range",
    rangeLocator: {
      type: "paragraph",
      startIndex: 0
    }
  },
  properties: {
    width: 150,
    height: 150
  },
  replaceAll: true
});

// 替换书签范围内的第一张图片
await replaceImage({
  locator: {
    type: "range",
    rangeLocator: {
      type: "bookmark",
      name: "图片区域"
    }
  },
  newImageData: "iVBORw0KGgo...", // 纯 Base64 字符串
  replaceAll: false
});

// 替换标题下的所有图片
await replaceImage({
  locator: {
    type: "range",
    rangeLocator: {
      type: "heading",
      level: 1,
      text: "图片展示"
    }
  },
  properties: {
    altText: "示例图片"
  },
  replaceAll: true
});
```

## 参数说明

### ReplaceImageOptions

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| locator | ReplaceImageLocator | ✅ | 图片定位器 |
| newImageData | string | ❌ | 新图片的 Base64 数据 |
| properties | ImageProperties | ❌ | 图片属性 |
| replaceAll | boolean | ❌ | 是否替换所有匹配项（默认 false） |

**注意**：`newImageData` 和 `properties` 至少需要提供一个。

### ImageProperties

| 属性 | 类型 | 说明 |
|------|------|------|
| width | number | 宽度（磅） |
| height | number | 高度（磅） |
| altText | string | 替代文本 |
| hyperlink | string | 超链接 |
| lockAspectRatio | boolean | 是否锁定纵横比 |

### ImageSearchOptions

| 属性 | 类型 | 说明 |
|------|------|------|
| altText | string | 按替代文本搜索（部分匹配） |
| minWidth | number | 最小宽度（磅） |
| maxWidth | number | 最大宽度（磅） |
| minHeight | number | 最小高度（磅） |
| maxHeight | number | 最大高度（磅） |

## 返回值

```typescript
interface ReplaceImageResult {
  count: number;      // 替换的图片数量
  success: boolean;   // 是否成功
  error?: string;     // 错误信息（如果有）
}
```

## 错误处理

```typescript
const result = await replaceImage({
  locator: { type: "index", index: 0 },
  newImageData: "iVBORw0KGgo..." // 纯 Base64 字符串
});

if (result.success) {
  console.log(`成功替换 ${result.count} 张图片`);
} else {
  console.error(`替换失败: ${result.error}`);
}
```

## 常见使用场景

### 场景 1: 批量更新图片尺寸

```typescript
// 将所有大图片缩小到统一尺寸
await replaceImage({
  locator: {
    type: "search",
    searchOptions: {
      minWidth: 500
    }
  },
  properties: {
    width: 400,
    lockAspectRatio: true
  },
  replaceAll: true
});
```

### 场景 2: 替换特定标记的图片

```typescript
// 替换所有标记为 "placeholder" 的图片
await replaceImage({
  locator: {
    type: "search",
    searchOptions: {
      altText: "placeholder"
    }
  },
  newImageData: actualImageBase64,
  replaceAll: true
});
```

### 场景 3: 为图片添加超链接

```typescript
// 为所有图片添加超链接
await replaceImage({
  locator: {
    type: "search",
    searchOptions: {}
  },
  properties: {
    hyperlink: "https://example.com"
  },
  replaceAll: true
});
```

## 注意事项

1. **Base64 格式**：`newImageData` 应为**纯 Base64 字符串**，**不要**包含数据 URI 前缀（如 `data:image/png;base64,`）。如果从 FileReader 获取，需要使用 `dataUrl.split(',')[1]` 提取纯 Base64 部分
2. **尺寸单位**：所有尺寸参数使用磅（points）作为单位
3. **批量操作**：使用 `replaceAll: true` 时要谨慎，确保搜索条件准确
4. **性能考虑**：替换大量图片时可能需要较长时间，建议分批处理
5. **锁定纵横比**：设置 `lockAspectRatio: true` 后，仅设置宽度或高度会自动调整另一维度
