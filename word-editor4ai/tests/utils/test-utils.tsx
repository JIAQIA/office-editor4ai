/**
 * 文件名: test-utils.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/react, @fluentui/react-components
 * 描述: 测试工具函数 | Test utility functions
 */

import * as React from "react";
import { ReactElement } from "react";
import { render, RenderOptions } from "@testing-library/react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { vi } from "vitest";

/**
 * 自定义渲染函数，包装 FluentUI Provider
 * Custom render function with FluentUI Provider wrapper
 */
interface CustomRenderOptions extends Omit<RenderOptions, "wrapper"> {
  theme?: typeof webLightTheme;
}

export function renderWithProviders(ui: ReactElement, options?: CustomRenderOptions) {
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
    text: "Mock paragraph text",
    style: "Normal",
    styleBuiltIn: "Normal",
    alignment: "Left",
    firstLineIndent: 0,
    leftIndent: 0,
    rightIndent: 0,
    lineSpacing: 1.5,
    spaceAfter: 10,
    spaceBefore: 0,
    isListItem: false,
    font: {
      name: "Arial",
      size: 12,
      bold: false,
      italic: false,
      color: "#000000",
      load: vi.fn(),
    },
    load: vi.fn(),
    select: vi.fn(),
  };

  // 创建模拟表格 / Create mock table
  const mockTable = {
    rowCount: 3,
    alignment: "Left",
    style: "TableGrid",
    styleBuiltIn: "TableGrid",
    width: 400,
    values: [
      ["1", "2", "3"],
      ["4", "5", "6"],
      ["7", "8", "9"],
    ],
    columns: {
      items: [
        { width: 100, load: vi.fn() },
        { width: 100, load: vi.fn() },
        { width: 100, load: vi.fn() },
      ],
      load: vi.fn(),
    },
    rows: {
      items: [
        {
          cells: { items: [], load: vi.fn() },
          load: vi.fn(),
        },
        {
          cells: { items: [], load: vi.fn() },
          load: vi.fn(),
        },
        {
          cells: { items: [], load: vi.fn() },
          load: vi.fn(),
        },
      ],
      load: vi.fn(),
      getFirst: vi.fn().mockReturnValue({
        cells: { items: [], load: vi.fn() },
        load: vi.fn(),
      }),
    },
    getCell: vi.fn((row: number, col: number) => ({
      value: `${row},${col}`,
      horizontalAlignment: "Left",
      verticalAlignment: "Top",
      shadingColor: "#FFFFFF",
      body: {
        font: {
          name: "Arial",
          size: 12,
          bold: false,
          italic: false,
          color: "#000000",
          load: vi.fn(),
        },
        load: vi.fn(),
      },
      load: vi.fn(),
    })),
    getBorder: vi.fn(() => ({
      type: "Single",
      width: 1,
      color: "#000000",
      load: vi.fn(),
    })),
    addRows: vi.fn(),
    addColumns: vi.fn(),
    delete: vi.fn(),
    load: vi.fn(),
  };

  const mockBody = {
    paragraphs: {
      items: [mockParagraph],
      load: vi.fn(),
    },
    tables: {
      items: [mockTable],
      load: vi.fn(),
    },
    getRange: vi.fn().mockReturnValue({
      text: "Mock text",
      paragraphs: {
        items: [mockParagraph],
        load: vi.fn(),
      },
      insertTable: vi.fn(
        (_rows: number, _cols: number, _location: string, _values?: string[][]) => mockTable
      ),
      load: vi.fn(),
    }),
    load: vi.fn(),
  };

  const mockSelection = {
    text: "Selected text",
    parentTableOrNullObject: {
      isNullObject: false,
      rowCount: 3,
      columns: {
        items: [
          { width: 100, load: vi.fn() },
          { width: 100, load: vi.fn() },
          { width: 100, load: vi.fn() },
        ],
        load: vi.fn(),
      },
      getCell: mockTable.getCell,
      getBorder: mockTable.getBorder,
      delete: vi.fn(),
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
      getSelection: vi.fn(() => mockSelection),
      load: vi.fn(),
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 创建包含标题段落的模拟 Word 上下文
 * Create mock Word context with heading paragraphs
 */
export function createMockWordContextWithHeadings(
  headings: Array<{ text: string; level: number; index: number }>
) {
  const mockParagraphs = headings.map((heading) => ({
    text: heading.text,
    style: `Heading ${heading.level}`,
    styleBuiltIn: `Heading${heading.level}`,
    alignment: "Left",
    font: {
      name: "Arial",
      size: 14 + (3 - heading.level) * 2, // 标题级别越高字体越大
      bold: true,
      italic: false,
      color: "#000000",
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

  // 模拟全局 Word 对象 / Mock global Word object
  (global as any).Word = {
    run: vi.fn((callback) => Promise.resolve(callback(context))),
    BuiltInStyleName: {
      tableGrid: "TableGrid",
      plainTable1: "PlainTable1",
      plainTable2: "PlainTable2",
      gridTable1Light: "GridTable1Light",
      gridTable2: "GridTable2",
      gridTable3: "GridTable3",
      gridTable4: "GridTable4",
      gridTable5Dark: "GridTable5Dark",
      gridTable6Colorful: "GridTable6Colorful",
      gridTable7Colorful: "GridTable7Colorful",
      listTable1Light: "ListTable1Light",
      listTable2: "ListTable2",
      listTable3: "ListTable3",
      listTable4: "ListTable4",
      listTable5Dark: "ListTable5Dark",
      listTable6Colorful: "ListTable6Colorful",
      listTable7Colorful: "ListTable7Colorful",
    },
    InsertLocation: {
      before: "Before",
      after: "After",
      start: "Start",
      end: "End",
      replace: "Replace",
    },
    Alignment: {
      left: "Left",
      centered: "Centered",
      right: "Right",
      justified: "Justified",
    },
    BorderLocation: {
      all: "All",
      top: "Top",
      bottom: "Bottom",
      left: "Left",
      right: "Right",
      inside: "Inside",
      outside: "Outside",
    },
    BorderType: {
      single: "Single",
      dotted: "Dotted",
      dashed: "Dashed",
      double: "Double",
      none: "None",
    },
    VerticalAlignment: {
      top: "Top",
      center: "Center",
      bottom: "Bottom",
    },
  };

  return (global as any).Word.run;
}

// 重新导出所有 testing-library 工具 | Re-export all testing-library utilities
export * from "@testing-library/react";
export { default as userEvent } from "@testing-library/user-event";
