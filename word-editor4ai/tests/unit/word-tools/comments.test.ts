/**
 * 文件名: comments.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 批注内容获取工具的测试文件 | Test file for comments tool
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { getComments, findDuplicateReferences } from "../../../src/word-tools";

describe("comments - 批注内容获取工具", () => {
  let originalWordRun: any;

  beforeEach(() => {
    // 保存原始的 Word.run / Save original Word.run
    originalWordRun = global.Word.run;
  });

  afterEach(() => {
    // 恢复原始的 Word.run / Restore original Word.run
    global.Word.run = originalWordRun;
    vi.clearAllMocks();
  });

  describe("getComments - 获取批注内容", () => {
    it("应该在没有选择时返回文档所有批注 | Should return all comments in document when no selection", async () => {
      const mockContext = createMockContextWithComments(false, 2);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getComments();

      expect(result).toBeDefined();
      expect(Array.isArray(result)).toBe(true);
      expect(result.length).toBe(2);
    });

    it("应该在有选择时返回选择范围内的批注 | Should return comments in selection when selection exists", async () => {
      const mockContext = createMockContextWithComments(true, 3);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getComments();

      expect(result).toBeDefined();
      expect(result.length).toBe(3);
    });

    it("应该在没有批注时返回空数组 | Should return empty array when no comments found", async () => {
      const mockContext = createMockContextWithComments(false, 0);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getComments();

      expect(result).toBeDefined();
      expect(result.length).toBe(0);
    });

    it("应该支持 includeResolved 选项 | Should support includeResolved option", async () => {
      const mockContext = createMockContextWithComments(false, 2, true);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithResolved = await getComments({ includeResolved: true });
      const resultWithoutResolved = await getComments({ includeResolved: false });

      expect(resultWithResolved.length).toBe(2);
      expect(resultWithoutResolved.length).toBe(1); // 只有一个未解决的批注
    });

    it("应该支持 includeReplies 选项 | Should support includeReplies option", async () => {
      const mockContext = createMockContextWithComments(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithReplies = await getComments({ includeReplies: true });
      const resultWithoutReplies = await getComments({ includeReplies: false });

      expect(resultWithReplies[0].replies).toBeDefined();
      expect(resultWithReplies[0].replies!.length).toBeGreaterThan(0);
      expect(resultWithoutReplies[0].replies).toBeUndefined();
    });

    it("应该支持 includeAssociatedText 选项 | Should support includeAssociatedText option", async () => {
      const mockContext = createMockContextWithComments(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithText = await getComments({ includeAssociatedText: true });
      const resultWithoutText = await getComments({ includeAssociatedText: false });

      expect(resultWithText[0].associatedText).toBeDefined();
      expect(resultWithoutText[0].associatedText).toBeUndefined();
    });

    it("应该支持 detailedMetadata 选项 | Should support detailedMetadata option", async () => {
      const mockContext = createMockContextWithComments(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const resultWithMetadata = await getComments({ detailedMetadata: true });
      const resultWithoutMetadata = await getComments({ detailedMetadata: false });

      expect(resultWithMetadata[0].authorName).toBeDefined();
      expect(resultWithMetadata[0].authorEmail).toBeDefined();
      expect(resultWithMetadata[0].creationDate).toBeDefined();
      expect(resultWithoutMetadata[0].authorName).toBeUndefined();
      expect(resultWithoutMetadata[0].authorEmail).toBeUndefined();
    });

    it("应该支持 maxTextLength 选项 | Should support maxTextLength option", async () => {
      const mockContext = createMockContextWithComments(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getComments({ maxTextLength: 10 });

      expect(result[0].content).toBeDefined();
      expect(result[0].content.length).toBeLessThanOrEqual(13); // 10 + "..."
    });

    it("应该正确处理批注属性 | Should handle comment properties correctly", async () => {
      const mockContext = createMockContextWithComments(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getComments({ detailedMetadata: true });

      expect(result[0]).toHaveProperty("id");
      expect(result[0]).toHaveProperty("content");
      expect(result[0]).toHaveProperty("resolved");
      expect(result[0]).toHaveProperty("authorName");
      expect(result[0]).toHaveProperty("authorEmail");
      expect(result[0]).toHaveProperty("creationDate");
    });

    it("应该正确处理批注回复 | Should handle comment replies correctly", async () => {
      const mockContext = createMockContextWithComments(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getComments({
        includeReplies: true,
        detailedMetadata: true,
      });

      expect(result[0].replies).toBeDefined();
      expect(result[0].replies!.length).toBeGreaterThan(0);

      const reply = result[0].replies![0];
      expect(reply).toHaveProperty("id");
      expect(reply).toHaveProperty("content");
      expect(reply).toHaveProperty("authorName");
    });

    it("应该正确处理错误 | Should handle errors correctly", async () => {
      global.Word.run = vi.fn(async () => {
        throw new Error("API Error");
      });

      await expect(getComments()).rejects.toThrow();
    });

    it("应该正确处理多条批注 | Should handle multiple comments correctly", async () => {
      const mockContext = createMockContextWithComments(false, 5);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getComments({ includeAssociatedText: true });

      expect(result.length).toBe(5);
      result.forEach((comment, index) => {
        expect(comment.id).toBe(`comment-${index + 1}`);
        expect(comment.content).toBeDefined();
      });
    });
  });

  describe("CommentInfo 类型验证 | CommentInfo type validation", () => {
    it("应该包含所有必需的字段 | Should contain all required fields", async () => {
      const mockContext = createMockContextWithComments(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getComments({ detailedMetadata: true, includeAssociatedText: true });

      expect(result[0]).toHaveProperty("id");
      expect(typeof result[0].id).toBe("string");
      expect(result[0]).toHaveProperty("content");
      expect(typeof result[0].content).toBe("string");
    });

    it("应该包含位置信息和元数据 | Should contain location info and metadata", async () => {
      const mockContext = createMockContextWithComments(false, 1);

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getComments({ includeAssociatedText: true });

      expect(result[0]).toHaveProperty("rangeLocation");
      expect(result[0].rangeLocation).toBeDefined();
      
      // 基本位置信息 / Basic location info
      expect(result[0].rangeLocation?.style).toBe("Normal");
      
      // 文本元数据 / Text metadata
      expect(result[0].rangeLocation?.textHash).toBeDefined();
      expect(result[0].rangeLocation?.textLength).toBe(29); // "Associated text for comment 1"
      
      // Range 位置信息 / Range position info
      expect(result[0].rangeLocation?.start).toBe(0);
      expect(result[0].rangeLocation?.end).toBe(29);
      expect(result[0].rangeLocation?.storyType).toBe("MainText");
      
      // 段落位置信息 / Paragraph position info
      expect(result[0].rangeLocation?.paragraphIndex).toBe(0);
      
      // 字体信息 / Font info
      expect(result[0].rangeLocation?.font).toBe("Calibri");
      expect(result[0].rangeLocation?.fontSize).toBe(11);
      expect(result[0].rangeLocation?.isBold).toBe(false);
      expect(result[0].rangeLocation?.isItalic).toBe(false);
      expect(result[0].rangeLocation?.isUnderlined).toBe(false);
    });
  });

  describe("findDuplicateReferences - 识别重复引用", () => {
    it("应该识别重复的批注引用 | Should identify duplicate comment references", async () => {
      // 创建包含重复引用的批注 / Create comments with duplicate references
      const comments = [
        {
          id: "comment-1",
          content: "First comment",
          resolved: false,
          associatedText: "Same text",
          rangeLocation: {
            textHash: "abc123",
            textLength: 9,
          },
        },
        {
          id: "comment-2",
          content: "Second comment",
          resolved: false,
          associatedText: "Same text",
          rangeLocation: {
            textHash: "abc123",
            textLength: 9,
          },
        },
        {
          id: "comment-3",
          content: "Third comment",
          resolved: false,
          associatedText: "Different text",
          rangeLocation: {
            textHash: "def456",
            textLength: 14,
          },
        },
      ];

      const duplicates = findDuplicateReferences(comments);

      expect(duplicates).toHaveLength(1);
      expect(duplicates[0].textHash).toBe("abc123");
      expect(duplicates[0].text).toBe("Same text");
      expect(duplicates[0].count).toBe(2);
      expect(duplicates[0].comments).toHaveLength(2);
      expect(duplicates[0].comments[0].id).toBe("comment-1");
      expect(duplicates[0].comments[1].id).toBe("comment-2");
    });

    it("应该在没有重复时返回空数组 | Should return empty array when no duplicates", () => {
      const comments = [
        {
          id: "comment-1",
          content: "First comment",
          resolved: false,
          associatedText: "Text 1",
          rangeLocation: {
            textHash: "abc123",
            textLength: 6,
          },
        },
        {
          id: "comment-2",
          content: "Second comment",
          resolved: false,
          associatedText: "Text 2",
          rangeLocation: {
            textHash: "def456",
            textLength: 6,
          },
        },
      ];

      const duplicates = findDuplicateReferences(comments);

      expect(duplicates).toHaveLength(0);
    });
  });
});

/**
 * 创建包含批注的模拟上下文
 * Create mock context with comments
 */
function createMockContextWithComments(
  hasSelection: boolean,
  commentCount: number,
  hasResolvedComments: boolean = false
) {
  // 创建模拟的批注 / Create mock comments
  const mockComments = Array.from({ length: commentCount }, (_, i) => {
    const mockReply = {
      id: `reply-${i + 1}`,
      content: `Reply ${i + 1} content`,
      authorName: `Reply Author ${i + 1}`,
      authorEmail: `reply${i + 1}@example.com`,
      creationDate: new Date(),
      load: vi.fn().mockReturnThis(),
    };

    return {
      id: `comment-${i + 1}`,
      content: `Comment ${i + 1} content - This is a test comment`,
      authorName: `Author ${i + 1}`,
      authorEmail: `author${i + 1}@example.com`,
      creationDate: new Date(),
      resolved: hasResolvedComments && i === 0, // 第一条批注已解决
      getRange: vi.fn().mockReturnValue({
        text: `Associated text for comment ${i + 1}`,
        start: i * 100,
        end: i * 100 + 29, // "Associated text for comment X" 长度为 29
        style: "Normal",
        isEmpty: false,
        font: {
          name: "Calibri",
          size: 11,
          bold: false,
          italic: false,
          underline: "None",
          highlightColor: "None",
          load: vi.fn().mockReturnThis(),
        },
        parentBody: {
          type: "MainText",
          paragraphs: {
            items: [
              {
                text: `Associated text for comment ${i + 1}`,
                load: vi.fn().mockReturnThis(),
              },
            ],
            load: vi.fn().mockReturnThis(),
          },
        },
        load: vi.fn().mockReturnThis(),
      }),
      replies: {
        items: [mockReply],
        load: vi.fn().mockReturnThis(),
      },
      load: vi.fn().mockReturnThis(),
    };
  });

  // 创建模拟的选择范围 / Create mock selection range
  const mockSelection = {
    isEmpty: !hasSelection,
    getComments: vi.fn().mockReturnValue({
      items: hasSelection ? mockComments : [],
      load: vi.fn().mockReturnThis(),
    }),
    load: vi.fn().mockReturnThis(),
  };

  // 创建模拟的段落 / Create mock paragraphs
  const mockParagraphs = Array.from({ length: commentCount }, (_, i) => ({
    text: `Associated text for comment ${i + 1}`,
    isListItem: false,
    listItem: null,
    load: vi.fn().mockReturnThis(),
  }));

  return {
    document: {
      getSelection: vi.fn().mockReturnValue(mockSelection),
      body: {
        getComments: vi.fn().mockReturnValue({
          items: mockComments,
          load: vi.fn().mockReturnThis(),
        }),
        paragraphs: {
          items: mockParagraphs,
          load: vi.fn().mockReturnThis(),
        },
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}
