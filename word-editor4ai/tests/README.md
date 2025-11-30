# Word Editor4AI 测试指南

## 测试目录结构

```
tests/
├── setup.ts              # 测试环境设置文件
├── unit/                 # 单元测试
│   └── word-tools/       # Word 工具测试
│       └── visibleContent.test.ts
├── integration/          # 集成测试
└── utils/                # 测试工具函数
```

## 运行测试

```bash
# 运行所有测试
pnpm test

# 运行测试（单次）
pnpm test:run

# 运行测试（监听模式）
pnpm test:watch

# 运行测试 UI
pnpm test:ui

# 运行单元测试
pnpm test:unit

# 运行集成测试
pnpm test:integration

# 生成覆盖率报告
pnpm test:coverage
```

## 测试配置

- **测试框架**: Vitest
- **测试环境**: jsdom (模拟浏览器环境)
- **全局变量**: 已启用 (describe, it, expect 等)
- **设置文件**: `tests/setup.ts`

## 编写测试

### 单元测试示例

```typescript
import { describe, it, expect } from 'vitest';
import { yourFunction } from '../../../src/your-module';

describe('yourFunction', () => {
  it('should do something', () => {
    const result = yourFunction();
    expect(result).toBe(expected);
  });
});
```

### 测试 Office.js 功能

测试环境已经配置了 Office.js 和 Word 对象的 mock，可以直接使用：

```typescript
it('should work with Word API', async () => {
  await Word.run(async (context) => {
    // 你的测试代码
  });
});
```

## 覆盖率要求

- 行覆盖率: 60%
- 函数覆盖率: 60%
- 分支覆盖率: 60%
- 语句覆盖率: 60%
