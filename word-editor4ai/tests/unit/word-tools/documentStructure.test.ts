/**
 * 文件名: documentStructure.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 文档结构工具的测试文件 | Test file for document structure tool
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  getDocumentOutline,
  getDocumentOutlineFlat,
  navigateToOutlineNode,
  exportOutlineAsMarkdown,
  exportOutlineAsJSON,
  type DocumentOutline,
  type OutlineNode,
} from "../../../src/word-tools";
import { createMockWordContextWithHeadings } from '../../utils/test-utils';

describe('documentStructure - 文档结构工具', () => {
  let originalWordRun: any;

  beforeEach(() => {
    // 保存原始的 Word.run
    originalWordRun = global.Word.run;
  });

  afterEach(() => {
    // 恢复原始的 Word.run
    global.Word.run = originalWordRun;
    vi.clearAllMocks();
  });

  describe('getDocumentOutline - 获取文档大纲（树形结构）', () => {
    it('应该返回正确的文档大纲结构 | Should return correct document outline structure', async () => {
      // 模拟包含多级标题的文档
      const mockHeadings = [
        { text: '第一章', level: 1, index: 0 },
        { text: '1.1 节', level: 2, index: 1 },
        { text: '1.2 节', level: 2, index: 2 },
        { text: '第二章', level: 1, index: 3 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutline();

      expect(result).toBeDefined();
      expect(result.nodes).toBeDefined();
      expect(result.totalHeadings).toBe(4);
      expect(result.maxDepth).toBe(2);
      expect(result.levelCounts).toEqual({ 1: 2, 2: 2 });
    });

    it('应该正确构建层级关系 | Should correctly build hierarchy', async () => {
      const mockHeadings = [
        { text: 'H1', level: 1, index: 0 },
        { text: 'H1.1', level: 2, index: 1 },
        { text: 'H1.1.1', level: 3, index: 2 },
        { text: 'H1.2', level: 2, index: 3 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutline();

      // 验证根节点数量
      expect(result.nodes).toHaveLength(1);
      
      // 验证第一级子节点
      const rootNode = result.nodes[0];
      expect(rootNode.text).toBe('H1');
      expect(rootNode.level).toBe(1);
      expect(rootNode.children).toHaveLength(2);

      // 验证第二级子节点
      expect(rootNode.children[0].text).toBe('H1.1');
      expect(rootNode.children[0].children).toHaveLength(1);
      expect(rootNode.children[0].children[0].text).toBe('H1.1.1');
      
      expect(rootNode.children[1].text).toBe('H1.2');
      expect(rootNode.children[1].children).toHaveLength(0);
    });

    it('应该支持 includeFormat 选项 | Should support includeFormat option', async () => {
      const mockHeadings = [
        { text: '标题', level: 1, index: 0 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutline({ includeFormat: true });

      expect(result.nodes[0].format).toBeDefined();
      expect(result.nodes[0].format?.font).toBe('Arial');
      expect(result.nodes[0].format?.bold).toBe(true);
    });

    it('应该支持 maxDepth 选项 | Should support maxDepth option', async () => {
      const mockHeadings = [
        { text: 'H1', level: 1, index: 0 },
        { text: 'H2', level: 2, index: 1 },
        { text: 'H3', level: 3, index: 2 },
        { text: 'H4', level: 4, index: 3 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutline({ maxDepth: 2 });

      // 只应该包含 H1 和 H2
      expect(result.totalHeadings).toBe(2);
      expect(result.maxDepth).toBe(2);
    });

    it('应该支持 specificLevels 选项 | Should support specificLevels option', async () => {
      const mockHeadings = [
        { text: 'H1', level: 1, index: 0 },
        { text: 'H2', level: 2, index: 1 },
        { text: 'H3', level: 3, index: 2 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutline({ specificLevels: [1, 3] });

      // 只应该包含 H1 和 H3
      expect(result.totalHeadings).toBe(2);
      expect(result.levelCounts).toEqual({ 1: 1, 3: 1 });
    });

    it('应该处理空文档 | Should handle empty document', async () => {
      const mockContext = createMockWordContextWithHeadings([]);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutline();

      expect(result.nodes).toHaveLength(0);
      expect(result.totalHeadings).toBe(0);
      expect(result.maxDepth).toBe(0);
    });

    it('应该处理错误情况 | Should handle error cases', async () => {
      global.Word.run = vi.fn(() => Promise.reject(new Error('Word API error')));

      await expect(getDocumentOutline()).rejects.toThrow();
    });
  });

  describe('getDocumentOutlineFlat - 获取文档大纲（扁平列表）', () => {
    it('应该返回扁平的大纲节点列表 | Should return flat outline node list', async () => {
      const mockHeadings = [
        { text: 'H1', level: 1, index: 0 },
        { text: 'H1.1', level: 2, index: 1 },
        { text: 'H1.2', level: 2, index: 2 },
        { text: 'H2', level: 1, index: 3 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutlineFlat();

      expect(result).toHaveLength(4);
      expect(result[0].text).toBe('H1');
      expect(result[0].children).toHaveLength(0); // 扁平列表中没有子节点
      expect(result[1].text).toBe('H1.1');
      expect(result[2].text).toBe('H1.2');
      expect(result[3].text).toBe('H2');
    });

    it('应该支持 includeFormat 选项 | Should support includeFormat option', async () => {
      const mockHeadings = [
        { text: '标题', level: 1, index: 0 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutlineFlat({ includeFormat: true });

      expect(result[0].format).toBeDefined();
      expect(result[0].format?.font).toBe('Arial');
    });

    it('应该支持过滤选项 | Should support filter options', async () => {
      const mockHeadings = [
        { text: 'H1', level: 1, index: 0 },
        { text: 'H2', level: 2, index: 1 },
        { text: 'H3', level: 3, index: 2 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutlineFlat({ maxDepth: 2 });

      expect(result).toHaveLength(2);
    });

    it('应该处理错误情况 | Should handle error cases', async () => {
      global.Word.run = vi.fn(() => Promise.reject(new Error('Word API error')));

      await expect(getDocumentOutlineFlat()).rejects.toThrow();
    });
  });

  describe('navigateToOutlineNode - 跳转到大纲节点', () => {
    it('应该正确跳转到指定节点 | Should navigate to specified node correctly', async () => {
      const mockParagraph = {
        select: vi.fn(),
      };

      const mockContext = {
        document: {
          body: {
            paragraphs: {
              items: [mockParagraph, mockParagraph, mockParagraph],
              load: vi.fn(),
            },
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await navigateToOutlineNode('heading-1');

      expect(mockContext.document.body.paragraphs.load).toHaveBeenCalled();
      expect(mockParagraph.select).toHaveBeenCalledWith('Start');
      expect(mockContext.sync).toHaveBeenCalled();
    });

    it('应该处理无效的节点ID | Should handle invalid node ID', async () => {
      const mockContext = createMockWordContextWithHeadings([]);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await expect(navigateToOutlineNode('invalid-id')).rejects.toThrow('无效的节点ID');
    });

    it('应该处理超出范围的索引 | Should handle out of range index', async () => {
      const mockContext = {
        document: {
          body: {
            paragraphs: {
              items: [],
              load: vi.fn(),
            },
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      await expect(navigateToOutlineNode('heading-10')).rejects.toThrow('段落索引超出范围');
    });

    it('应该处理错误情况 | Should handle error cases', async () => {
      global.Word.run = vi.fn(() => Promise.reject(new Error('Word API error')));

      await expect(navigateToOutlineNode('heading-0')).rejects.toThrow();
    });
  });

  describe('exportOutlineAsMarkdown - 导出为 Markdown', () => {
    it('应该正确导出为 Markdown 格式 | Should export to Markdown format correctly', () => {
      const outline: DocumentOutline = {
        nodes: [
          {
            id: 'heading-0',
            text: '第一章',
            level: 1,
            style: 'Heading 1',
            children: [
              {
                id: 'heading-1',
                text: '1.1 节',
                level: 2,
                style: 'Heading 2',
                children: [],
                index: 1,
              },
            ],
            index: 0,
          },
          {
            id: 'heading-2',
            text: '第二章',
            level: 1,
            style: 'Heading 1',
            children: [],
            index: 2,
          },
        ],
        totalHeadings: 3,
        maxDepth: 2,
        levelCounts: { 1: 2, 2: 1 },
      };

      const markdown = exportOutlineAsMarkdown(outline);

      expect(markdown).toContain('# 第一章');
      expect(markdown).toContain('## 1.1 节');
      expect(markdown).toContain('# 第二章');
    });

    it('应该处理空大纲 | Should handle empty outline', () => {
      const outline: DocumentOutline = {
        nodes: [],
        totalHeadings: 0,
        maxDepth: 0,
        levelCounts: {},
      };

      const markdown = exportOutlineAsMarkdown(outline);

      expect(markdown).toBe('');
    });

    it('应该正确处理多级嵌套 | Should handle multi-level nesting correctly', () => {
      const outline: DocumentOutline = {
        nodes: [
          {
            id: 'heading-0',
            text: 'H1',
            level: 1,
            style: 'Heading 1',
            children: [
              {
                id: 'heading-1',
                text: 'H2',
                level: 2,
                style: 'Heading 2',
                children: [
                  {
                    id: 'heading-2',
                    text: 'H3',
                    level: 3,
                    style: 'Heading 3',
                    children: [],
                    index: 2,
                  },
                ],
                index: 1,
              },
            ],
            index: 0,
          },
        ],
        totalHeadings: 3,
        maxDepth: 3,
        levelCounts: { 1: 1, 2: 1, 3: 1 },
      };

      const markdown = exportOutlineAsMarkdown(outline);

      expect(markdown).toContain('# H1');
      expect(markdown).toContain('## H2');
      expect(markdown).toContain('### H3');
    });
  });

  describe('exportOutlineAsJSON - 导出为 JSON', () => {
    it('应该正确导出为 JSON 格式 | Should export to JSON format correctly', () => {
      const outline: DocumentOutline = {
        nodes: [
          {
            id: 'heading-0',
            text: '标题',
            level: 1,
            style: 'Heading 1',
            children: [],
            index: 0,
          },
        ],
        totalHeadings: 1,
        maxDepth: 1,
        levelCounts: { 1: 1 },
      };

      const json = exportOutlineAsJSON(outline);
      const parsed = JSON.parse(json);

      expect(parsed).toEqual(outline);
      expect(parsed.nodes).toHaveLength(1);
      expect(parsed.nodes[0].text).toBe('标题');
    });

    it('应该包含所有属性 | Should include all properties', () => {
      const outline: DocumentOutline = {
        nodes: [
          {
            id: 'heading-0',
            text: '标题',
            level: 1,
            style: 'Heading 1',
            children: [],
            index: 0,
            format: {
              font: 'Arial',
              fontSize: 16,
              bold: true,
              italic: false,
              color: '#000000',
              alignment: 'Left',
            },
          },
        ],
        totalHeadings: 1,
        maxDepth: 1,
        levelCounts: { 1: 1 },
      };

      const json = exportOutlineAsJSON(outline);
      const parsed = JSON.parse(json);

      expect(parsed.nodes[0].format).toBeDefined();
      expect(parsed.nodes[0].format.font).toBe('Arial');
      expect(parsed.nodes[0].format.bold).toBe(true);
    });

    it('应该格式化 JSON 输出 | Should format JSON output', () => {
      const outline: DocumentOutline = {
        nodes: [],
        totalHeadings: 0,
        maxDepth: 0,
        levelCounts: {},
      };

      const json = exportOutlineAsJSON(outline);

      // 检查是否包含换行和缩进
      expect(json).toContain('\n');
      expect(json).toContain('  ');
    });
  });

  describe('边界情况测试 | Edge cases', () => {
    it('应该处理中文标题样式 | Should handle Chinese heading styles', async () => {
      const mockParagraphs = [
        {
          text: '中文标题',
          style: '标题 1',
          styleBuiltIn: 'Heading1',
          font: {
            name: '宋体',
            size: 16,
            bold: true,
            italic: false,
            color: '#000000',
            load: vi.fn(),
          },
          load: vi.fn(),
        },
      ];

      const mockContext = {
        document: {
          body: {
            paragraphs: {
              items: mockParagraphs,
              load: vi.fn(),
            },
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutline();

      expect(result.totalHeadings).toBe(1);
      expect(result.nodes[0].text).toBe('中文标题');
    });

    it('应该处理空标题文本 | Should handle empty heading text', async () => {
      const mockHeadings = [
        { text: '', level: 1, index: 0 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutline();

      expect(result.totalHeadings).toBe(1);
      expect(result.nodes[0].text).toBe('');
    });

    it('应该处理连续相同级别的标题 | Should handle consecutive same-level headings', async () => {
      const mockHeadings = [
        { text: 'H1-1', level: 1, index: 0 },
        { text: 'H1-2', level: 1, index: 1 },
        { text: 'H1-3', level: 1, index: 2 },
      ];

      const mockContext = createMockWordContextWithHeadings(mockHeadings);
      global.Word.run = vi.fn((callback) => Promise.resolve(callback(mockContext)));

      const result = await getDocumentOutline();

      expect(result.nodes).toHaveLength(3);
      expect(result.nodes[0].children).toHaveLength(0);
      expect(result.nodes[1].children).toHaveLength(0);
      expect(result.nodes[2].children).toHaveLength(0);
    });
  });
});
