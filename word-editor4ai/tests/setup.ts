/**
 * 文件名: setup.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
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
      callback({ host: 'Word', platform: 'PC' });
    }
    return Promise.resolve({ host: 'Word', platform: 'PC' });
  }),
  context: {
    document: {},
    mailbox: {},
  },
  actions: {
    associate: vi.fn(),
  },
} as any;

// 模拟 Word 对象 | Mock Word object
global.Word = {
  run: vi.fn((callback) => {
    const mockParagraph = {
      text: 'Mock paragraph text',
      style: 'Normal',
      alignment: Word.Alignment.left,
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
      lineSpacing: 1.5,
      spaceAfter: 10,
      spaceBefore: 0,
      isListItem: false,
      load: vi.fn(),
    };

    const mockRange = {
      text: 'Mock text',
      paragraphs: {
        items: [mockParagraph],
        load: vi.fn(),
      },
      load: vi.fn(),
    };

    const mockBody = {
      paragraphs: {
        items: [mockParagraph],
        load: vi.fn(),
      },
      getRange: vi.fn().mockReturnValue(mockRange),
      load: vi.fn(),
    };

    const context = {
      document: {
        body: mockBody,
        sections: {
          items: [],
          load: vi.fn(),
        },
        load: vi.fn(),
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    return Promise.resolve(callback(context));
  }),
  Alignment: {
    left: 'Left',
    centered: 'Centered',
    right: 'Right',
    justified: 'Justified',
  },
  ShapeType: {
    unsupported: 'Unsupported',
    textBox: 'TextBox',
    geometricShape: 'GeometricShape',
    group: 'Group',
    picture: 'Picture',
    canvas: 'Canvas',
  },
} as any;

// Mock ResizeObserver for FluentUI components
global.ResizeObserver = class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
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
