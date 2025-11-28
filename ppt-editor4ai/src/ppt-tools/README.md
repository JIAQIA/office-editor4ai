# PPT Tools

这个目录包含所有与 PowerPoint Office API 交互的核心工具逻辑。

## 目录结构

```
ppt-tools/
├── index.ts              # 统一导出入口
├── textInsertion.ts      # 文本插入工具
├── elementsList.ts       # 元素列表获取工具
└── README.md            # 本文档
```

## 设计原则

### 1. 独立性
每个工具的核心逻辑都封装在独立的 `.ts` 文件中，不依赖 React 或 UI 组件。这样设计的好处：
- 可以在不同的入口调用（UI 组件、Socket.IO、命令行等）
- 便于单元测试
- 逻辑清晰，职责单一

### 2. 类型安全
所有工具都提供完整的 TypeScript 类型定义：
- 输入参数类型（Options 接口）
- 返回值类型
- 导出的类型定义供外部使用

### 3. 错误处理
每个工具都包含完善的错误处理：
- try-catch 捕获异常
- console.error 记录错误
- throw error 向上传递异常

## 使用示例

### 在 React 组件中使用

```typescript
import { insertText, getCurrentSlideElements } from '../../../ppt-tools';

// 插入文本
await insertText("Hello World", 100, 100);

// 获取元素列表
const elements = await getCurrentSlideElements();
```

### 在 Socket.IO 中使用

```typescript
import { insertText, getSlideElements } from './ppt-tools';

socket.on('insertText', async (data) => {
  try {
    await insertText(data.text, data.left, data.top);
    socket.emit('success', { message: '文本插入成功' });
  } catch (error) {
    socket.emit('error', { message: error.message });
  }
});
```

## 添加新工具

1. 在 `ppt-tools/` 目录创建新的 `.ts` 文件
2. 定义工具的接口和函数
3. 在 `index.ts` 中导出
4. 在 `taskpane/components/tools/` 创建对应的 UI 组件
5. 在 `toolsConfig.tsx` 中注册工具

## 现有工具

### textInsertion.ts
- `insertText(text, left?, top?)` - 简化版本
- `insertTextToSlide(options)` - 完整版本，支持更多配置

### elementsList.ts
- `getCurrentSlideElements()` - 获取当前幻灯片元素
- `getSlideElements(options)` - 获取指定幻灯片元素，支持配置

## 注意事项

1. 所有工具函数都是异步的（返回 Promise）
2. 工具函数内部使用 `PowerPoint.run()` 与 Office API 交互
3. 坐标和尺寸单位统一使用"磅"（points）
4. 工具函数应该专注于单一职责，复杂功能可以组合多个工具
