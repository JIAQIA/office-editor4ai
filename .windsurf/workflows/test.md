---
description: 测试最佳实践
---

## 项目测试最佳实践总结

### 1. 测试框架与核心依赖包

**主测试框架：**
- **Vitest 4.0.14** - 现代化的单元测试框架，基于 Vite 构建
- **@vitest/ui 4.0.14** - Vitest 的可视化测试界面

**测试工具库：**
- **@testing-library/react 16.3.0** - React 组件测试工具
- **@testing-library/user-event 14.6.1** - 模拟用户交互
- **@testing-library/jest-dom 6.9.1** - DOM 断言扩展

**Mock 工具：**
- **office-addin-mock 3.0.6** - Office.js API 的 Mock 工具
- **jsdom 27.2.0** - 浏览器环境模拟

**其他工具：**
- **@vitejs/plugin-react 5.1.1** - React 支持插件

### 2. 测试文件目录编排模式

```
tests/
├── setup.ts                          # 全局测试配置文件
├── utils/                            # 测试工具函数
│   └── test-utils.tsx               # 自定义渲染函数、Mock 工具
├── unit/                            # 单元测试
│   ├── components/                  # 组件单元测试
│   │   ├── App.test.tsx
│   │   ├── HomePage.test.tsx
│   │   └── TextInsertion.test.tsx
│   └── ppt-tools/                   # 工具函数单元测试
│       ├── elementsList.test.ts
│       └── textInsertion.test.ts
└── integration/                     # 集成测试
    ├── app-navigation.integration.test.tsx
    └── text-insertion.integration.test.tsx
```

**命名规范：**
- 单元测试：`[功能名].test.tsx` 或 `[功能名].test.ts`
- 集成测试：`[功能名].integration.test.tsx`
- 测试文件与源文件路径对应：`src/taskpane/components/App.tsx` → [tests/unit/components/App.test.tsx](cci:7://file:///Users/JQQ/WebstormProjects/office-editor4ai/ppt-editor4ai/tests/unit/components/App.test.tsx:0:0-0:0)

### 3. 测试配置要点

**vitest.config.ts 关键配置：**
```typescript
{
  environment: 'jsdom',           // 浏览器环境模拟
  globals: true,                  // 全局 API 支持
  setupFiles: ['./tests/setup.ts'], // 全局设置文件
  coverage: {
    provider: 'v8',
    thresholds: {                 // 覆盖率阈值
      lines: 60,
      functions: 60,
      branches: 60,
      statements: 60
    }
  },
  testTimeout: 10000,             // 测试超时 10 秒
}
```

### 4. 测试编写模式

关于用例中的Mock方法你需要阅读测试附近的兄弟用例的Mock方式然后进行撰写

### 5. 测试工具函数

**核心工具函数（tests/utils/test-utils.tsx）：**
- [renderWithProviders()](cci:1://file:///Users/JQQ/WebstormProjects/office-editor4ai/ppt-editor4ai/tests/utils/test-utils.tsx:24:0-35:1) - 包装 FluentUI Provider 的自定义渲染函数
- [createMockOfficeContext()](cci:1://file:///Users/JQQ/WebstormProjects/office-editor4ai/ppt-editor4ai/tests/utils/test-utils.tsx:37:0-66:1) - 创建 Office 上下文 Mock
- [mockPowerPointRun()](cci:1://file:///Users/JQQ/WebstormProjects/office-editor4ai/ppt-editor4ai/tests/utils/test-utils.tsx:74:0-81:1) - Mock PowerPoint.run 调用
- [waitForAsync()](cci:1://file:///Users/JQQ/WebstormProjects/office-editor4ai/ppt-editor4ai/tests/utils/test-utils.tsx:68:0-72:83) - 等待异步操作

### 6. 运行测试的命令

```bash
pnpm test              # 运行所有测试
pnpm test:run          # 单次运行测试
pnpm test:watch        # 监听模式
pnpm test:ui           # 可视化界面
pnpm test:coverage     # 生成覆盖率报告
pnpm test:unit         # 仅运行单元测试
pnpm test:integration  # 仅运行集成测试
```

### 7. 最佳实践要点

1. **测试分层明确**：单元测试（unit）和集成测试（integration）分开
2. **Mock 策略**：使用 `office-addin-mock` 模拟 Office.js API，使用 `vi.mock()` 模拟组件依赖
3. **全局设置**：在 [setup.ts](cci:7://file:///Users/JQQ/WebstormProjects/office-editor4ai/ppt-editor4ai/tests/setup.ts:0:0-0:0) 中配置 Office.js 和 PowerPoint 全局对象
4. **自定义工具**：封装常用的测试工具函数到 [test-utils.tsx](cci:7://file:///Users/JQQ/WebstormProjects/office-editor4ai/ppt-editor4ai/tests/utils/test-utils.tsx:0:0-0:0)
5. **中英双语注释**：所有测试用例使用中英文双语描述
6. **覆盖率要求**：设置 60% 的覆盖率阈值
7. **边界测试**：包含边界情况测试（空值、特殊字符、负数等）