/**
 * 文件名: test-utils.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/react, @fluentui/react-components
 * 描述: 测试工具函数 | Test utility functions
 */

import * as React from 'react';
import { ReactElement } from 'react';
import { render, RenderOptions } from '@testing-library/react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { vi } from 'vitest';

/**
 * 自定义渲染函数，包装 FluentUI Provider
 * Custom render function with FluentUI Provider wrapper
 */
interface CustomRenderOptions extends Omit<RenderOptions, 'wrapper'> {
  theme?: typeof webLightTheme;
}

export function renderWithProviders(
  ui: ReactElement,
  options?: CustomRenderOptions
) {
  const { theme = webLightTheme, ...renderOptions } = options || {};

  function Wrapper({ children }: { children: React.ReactNode }) {
    return <FluentProvider theme={theme}>{children}</FluentProvider>;
  }

  return render(ui, { wrapper: Wrapper, ...renderOptions });
}

/**
 * 创建模拟的 Word 上下文
 * Create mock Word context
 */
export function createMockWordContext() {
  const mockParagraph = {
    text: 'Mock paragraph text',
    style: 'Normal',
    styleBuiltIn: 'Normal',
    alignment: 'Left',
    firstLineIndent: 0,
    leftIndent: 0,
    rightIndent: 0,
    lineSpacing: 1.5,
    spaceAfter: 10,
    spaceBefore: 0,
    isListItem: false,
    font: {
      name: 'Arial',
      size: 12,
      bold: false,
      italic: false,
      color: '#000000',
      load: vi.fn(),
    },
    load: vi.fn(),
    select: vi.fn(),
  };

  const mockBody = {
    paragraphs: {
      items: [mockParagraph],
      load: vi.fn(),
    },
    getRange: vi.fn().mockReturnValue({
      text: 'Mock text',
      paragraphs: {
        items: [mockParagraph],
        load: vi.fn(),
      },
      load: vi.fn(),
    }),
    load: vi.fn(),
  };

  return {
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
}

/**
 * 创建包含标题段落的模拟 Word 上下文
 * Create mock Word context with heading paragraphs
 */
export function createMockWordContextWithHeadings(headings: Array<{ text: string; level: number; index: number }>) {
  const mockParagraphs = headings.map((heading) => ({
    text: heading.text,
    style: `Heading ${heading.level}`,
    styleBuiltIn: `Heading${heading.level}`,
    alignment: 'Left',
    font: {
      name: 'Arial',
      size: 14 + (3 - heading.level) * 2, // 标题级别越高字体越大
      bold: true,
      italic: false,
      color: '#000000',
      load: vi.fn(),
    },
    load: vi.fn(),
    select: vi.fn(),
  }));

  const mockBody = {
    paragraphs: {
      items: mockParagraphs,
      load: vi.fn(),
    },
    load: vi.fn(),
  };

  return {
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
}

/**
 * 等待异步操作完成
 * Wait for async operations to complete
 */
export const waitForAsync = () => new Promise((resolve) => setTimeout(resolve, 0));

/**
 * 模拟 Word.run 调用
 * Mock Word.run call
 */
export function mockWordRun(mockContext?: any) {
  const context = mockContext || createMockWordContext();
  return vi.fn((callback) => Promise.resolve(callback(context)));
}

// 重新导出所有 testing-library 工具 | Re-export all testing-library utilities
export * from '@testing-library/react';
export { default as userEvent } from '@testing-library/user-event';
