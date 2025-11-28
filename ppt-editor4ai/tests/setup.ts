/**
 * 文件名: setup.ts
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/jest-dom, vitest
 * 描述: 测试环境设置文件 | Test environment setup file
 */

import '@testing-library/jest-dom';
import { expect, afterEach, vi } from 'vitest';
import { cleanup } from '@testing-library/react';

// 每个测试后自动清理 | Automatically cleanup after each test
afterEach(() => {
  cleanup();
});

// 模拟 Office.js 全局对象 | Mock Office.js global object
global.Office = {
  onReady: vi.fn((callback) => {
    if (typeof callback === 'function') {
      callback({ host: 'PowerPoint', platform: 'PC' });
    }
    return Promise.resolve({ host: 'PowerPoint', platform: 'PC' });
  }),
  context: {
    document: {},
    mailbox: {},
  },
  actions: {
    associate: vi.fn(),
  },
} as any;

// 模拟 PowerPoint 对象 | Mock PowerPoint object
global.PowerPoint = {
  run: vi.fn((callback) => {
    const context = {
      presentation: {
        slides: {
          getItemAt: vi.fn(),
          add: vi.fn(),
        },
        load: vi.fn(),
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    return Promise.resolve(callback(context));
  }),
} as any;

// 扩展 expect 匹配器 | Extend expect matchers
expect.extend({
  toBeInTheDocument(received) {
    const pass = received !== null && received !== undefined;
    return {
      pass,
      message: () => `expected element ${pass ? 'not ' : ''}to be in the document`,
    };
  },
});
