/**
 * 文件名: headerFooterContent.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 页眉页脚内容获取工具的单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import {
  getHeaderFooterContent,
  type GetHeaderFooterContentOptions,
  HeaderFooterType,
} from "../../../src/word-tools";

// Mock Word API
const mockContext = {
  sync: vi.fn().mockResolvedValue(undefined),
  document: {
    sections: {
      load: vi.fn(),
      items: [] as any[],
    },
  },
};

const mockSection = {
  body: {
    parentSection: {
      pageSetup: {
        load: vi.fn(),
        differentFirstPageHeaderFooter: false,
        oddAndEvenPagesHeaderFooter: false,
      },
    },
  },
  getHeader: vi.fn(),
  getFooter: vi.fn(),
};

const mockHeaderFooterBody = {
  load: vi.fn(),
  text: "",
  paragraphs: {
    load: vi.fn(),
    items: [] as any[],
  },
  tables: {
    load: vi.fn(),
    items: [] as any[],
  },
  contentControls: {
    load: vi.fn(),
    items: [] as any[],
  },
};

// Mock Word.run
const mockWordRun = vi.fn((callback: (context: any) => Promise<any>) => {
  return callback(mockContext);
});

// 设置全局 Word 对象
(global as any).Word = {
  run: mockWordRun,
  HeaderFooterType: {
    primary: "primary",
    firstPage: "firstPage",
    evenPages: "evenPages",
  },
};

describe("getHeaderFooterContent", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockContext.document.sections.items = [];
  });

  it("应该成功获取单个节的页眉页脚内容", async () => {
    // 准备测试数据 / Prepare test data
    const headerBody = {
      ...mockHeaderFooterBody,
      text: "测试页眉内容",
    };
    const footerBody = {
      ...mockHeaderFooterBody,
      text: "测试页脚内容",
    };

    mockSection.getHeader.mockReturnValue(headerBody);
    mockSection.getFooter.mockReturnValue(footerBody);
    mockContext.document.sections.items = [mockSection];

    const options: GetHeaderFooterContentOptions = {
      includeElements: false,
      includeMetadata: true,
    };

    const result = await getHeaderFooterContent(options);

    expect(result).toBeDefined();
    expect(result.totalSections).toBe(1);
    expect(result.sections).toHaveLength(1);
    expect(mockContext.sync).toHaveBeenCalled();
  });

  it("应该支持获取指定节的页眉页脚内容", async () => {
    // 准备多个节 / Prepare multiple sections
    const section1 = { ...mockSection };
    const section2 = { ...mockSection };

    mockContext.document.sections.items = [section1, section2];

    const headerBody = {
      ...mockHeaderFooterBody,
      text: "节1页眉",
    };

    section1.getHeader.mockReturnValue(headerBody);
    section1.getFooter.mockReturnValue(mockHeaderFooterBody);

    const options: GetHeaderFooterContentOptions = {
      sectionIndex: 0,
      includeElements: false,
    };

    const result = await getHeaderFooterContent(options);

    expect(result).toBeDefined();
    expect(result.sections).toHaveLength(1);
    expect(result.sections[0].sectionIndex).toBe(0);
  });

  it("应该在节索引超出范围时抛出错误", async () => {
    mockContext.document.sections.items = [mockSection];

    const options: GetHeaderFooterContentOptions = {
      sectionIndex: 5,
    };

    await expect(getHeaderFooterContent(options)).rejects.toThrow();
  });

  it("应该正确处理空页眉页脚", async () => {
    const emptyBody = {
      ...mockHeaderFooterBody,
      text: "",
    };

    mockSection.getHeader.mockReturnValue(emptyBody);
    mockSection.getFooter.mockReturnValue(emptyBody);
    mockContext.document.sections.items = [mockSection];

    const result = await getHeaderFooterContent({
      includeMetadata: true,
    });

    expect(result).toBeDefined();
    expect(result.sections[0].headers).toBeDefined();
    expect(result.sections[0].footers).toBeDefined();
  });

  it("应该在包含元数据时计算统计信息", async () => {
    const headerBody = {
      ...mockHeaderFooterBody,
      text: "页眉内容",
    };
    const footerBody = {
      ...mockHeaderFooterBody,
      text: "页脚内容",
    };

    mockSection.getHeader.mockReturnValue(headerBody);
    mockSection.getFooter.mockReturnValue(footerBody);
    mockContext.document.sections.items = [mockSection];

    const result = await getHeaderFooterContent({
      includeMetadata: true,
    });

    expect(result.metadata).toBeDefined();
    expect(result.metadata?.hasAnyHeader).toBeDefined();
    expect(result.metadata?.hasAnyFooter).toBeDefined();
    expect(result.metadata?.totalHeaders).toBeGreaterThanOrEqual(0);
    expect(result.metadata?.totalFooters).toBeGreaterThanOrEqual(0);
  });

  it("应该支持包含详细内容元素", async () => {
    const paragraph = {
      load: vi.fn(),
      text: "段落文本",
      style: "Normal",
      alignment: "Left",
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
      lineSpacing: 1.15,
      spaceAfter: 8,
      spaceBefore: 0,
      isListItem: false,
      inlinePictures: {
        load: vi.fn(),
        items: [],
      },
    };

    const bodyWithElements = {
      ...mockHeaderFooterBody,
      text: "页眉内容",
      paragraphs: {
        load: vi.fn(),
        items: [paragraph],
      },
    };

    mockSection.getHeader.mockReturnValue(bodyWithElements);
    mockSection.getFooter.mockReturnValue(mockHeaderFooterBody);
    mockContext.document.sections.items = [mockSection];

    const result = await getHeaderFooterContent({
      includeElements: true,
    });

    expect(result).toBeDefined();
    expect(result.sections[0].headers).toBeDefined();
  });

  it("应该正确处理不同首页和奇偶页设置", async () => {
    const sectionWithDifferentPages = {
      ...mockSection,
      body: {
        parentSection: {
          pageSetup: {
            load: vi.fn(),
            differentFirstPageHeaderFooter: true,
            oddAndEvenPagesHeaderFooter: true,
          },
        },
      },
    };

    mockContext.document.sections.items = [sectionWithDifferentPages];

    const result = await getHeaderFooterContent({});

    expect(result).toBeDefined();
    expect(result.sections[0].differentFirstPage).toBe(true);
    expect(result.sections[0].differentOddAndEven).toBe(true);
  });

  it("应该处理多个节的情况", async () => {
    const section1 = { ...mockSection };
    const section2 = { ...mockSection };
    const section3 = { ...mockSection };

    mockContext.document.sections.items = [section1, section2, section3];

    const result = await getHeaderFooterContent({});

    expect(result).toBeDefined();
    expect(result.totalSections).toBe(3);
    expect(result.sections).toHaveLength(3);
    expect(result.sections[0].sectionIndex).toBe(0);
    expect(result.sections[1].sectionIndex).toBe(1);
    expect(result.sections[2].sectionIndex).toBe(2);
  });

  it("应该正确识别页眉页脚类型", async () => {
    const headerFirst = { ...mockHeaderFooterBody, text: "首页页眉" };
    const headerPrimary = { ...mockHeaderFooterBody, text: "奇数页页眉" };
    const headerEven = { ...mockHeaderFooterBody, text: "偶数页页眉" };

    mockSection.getHeader
      .mockReturnValueOnce(headerFirst)
      .mockReturnValueOnce(headerPrimary)
      .mockReturnValueOnce(headerEven);

    mockSection.getFooter.mockReturnValue(mockHeaderFooterBody);

    mockContext.document.sections.items = [mockSection];

    const result = await getHeaderFooterContent({});

    expect(result).toBeDefined();
    expect(result.sections[0].headers).toHaveLength(3);
    expect(result.sections[0].headers[0].type).toBe(HeaderFooterType.FirstPage);
    expect(result.sections[0].headers[1].type).toBe(HeaderFooterType.OddPages);
    expect(result.sections[0].headers[2].type).toBe(HeaderFooterType.EvenPages);
  });

  it("应该在发生错误时抛出有意义的错误信息", async () => {
    mockContext.sync.mockRejectedValueOnce(new Error("同步失败"));

    await expect(getHeaderFooterContent({})).rejects.toThrow("获取页眉页脚内容失败");
  });
});
