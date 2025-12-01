/**
 * 文件名: documentStructure.ts
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取文档结构（大纲）的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

/**
 * 大纲节点 / Outline Node
 * 表示文档中的一个标题节点
 */
export interface OutlineNode {
  /** 节点唯一标识 / Unique identifier */
  id: string;
  /** 标题文本 / Heading text */
  text: string;
  /** 标题级别 (1-9) / Heading level (1-9) */
  level: number;
  /** 标题样式名称 / Heading style name */
  style: string;
  /** 子节点 / Child nodes */
  children: OutlineNode[];
  /** 段落索引（在文档中的位置）/ Paragraph index in document */
  index: number;
  /** 格式信息 / Format information */
  format?: {
    /** 字体名称 / Font name */
    font?: string;
    /** 字体大小 / Font size */
    fontSize?: number;
    /** 是否加粗 / Is bold */
    bold?: boolean;
    /** 是否斜体 / Is italic */
    italic?: boolean;
    /** 字体颜色 / Font color */
    color?: string;
    /** 对齐方式 / Alignment */
    alignment?: string;
  };
}

/**
 * 文档大纲结构 / Document Outline Structure
 */
export interface DocumentOutline {
  /** 根节点列表 / Root nodes */
  nodes: OutlineNode[];
  /** 总标题数量 / Total heading count */
  totalHeadings: number;
  /** 最大层级深度 / Maximum depth level */
  maxDepth: number;
  /** 各层级标题数量统计 / Heading count by level */
  levelCounts: Record<number, number>;
}

/**
 * 获取文档大纲选项 / Get Document Outline Options
 */
export interface GetDocumentOutlineOptions {
  /** 是否包含格式信息 / Include format information */
  includeFormat?: boolean;
  /** 最大层级深度（0表示不限制）/ Maximum depth (0 means no limit) */
  maxDepth?: number;
  /** 是否只获取特定层级 / Only get specific levels */
  specificLevels?: number[];
}

/**
 * 判断段落是否为标题样式
 * Determine if a paragraph is a heading style
 */
function isHeadingStyle(style: string): boolean {
  // 标准标题样式：Heading 1-9, 标题 1-9
  const headingPattern = /^(Heading|标题)\s*(\d)$/i;
  return headingPattern.test(style);
}

/**
 * 从样式名称中提取标题级别
 * Extract heading level from style name
 */
function extractHeadingLevel(style: string): number {
  const headingPattern = /^(Heading|标题)\s*(\d)$/i;
  const match = style.match(headingPattern);
  if (match && match[2]) {
    return parseInt(match[2], 10);
  }
  return 0;
}

/**
 * 获取文档大纲结构
 * Get document outline structure
 *
 * @param options - 获取选项 / Get options
 * @returns 文档大纲结构 / Document outline structure
 */
export async function getDocumentOutline(
  options: GetDocumentOutlineOptions = {}
): Promise<DocumentOutline> {
  const { includeFormat = false, maxDepth = 0, specificLevels } = options;

  return Word.run(async (context) => {
    try {
      // 获取文档的所有段落 / Get all paragraphs in the document
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");

      await context.sync();

      // 加载段落的样式和文本 / Load paragraph styles and text
      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("text,style,styleBuiltIn");

        if (includeFormat) {
          para.font.load("name,size,bold,italic,color");
          para.load("alignment");
        }
      }

      await context.sync();

      // 过滤出标题段落 / Filter heading paragraphs
      const headingParagraphs: Array<{
        paragraph: Word.Paragraph;
        index: number;
        level: number;
        text: string;
        style: string;
      }> = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const style = para.style;

        if (isHeadingStyle(style)) {
          const level = extractHeadingLevel(style);

          // 应用层级过滤 / Apply level filtering
          if (specificLevels && !specificLevels.includes(level)) {
            continue;
          }
          if (maxDepth > 0 && level > maxDepth) {
            continue;
          }

          headingParagraphs.push({
            paragraph: para,
            index: i,
            level,
            text: para.text.trim(),
            style,
          });
        }
      }

      // 构建大纲节点 / Build outline nodes
      const nodes: OutlineNode[] = [];
      const levelCounts: Record<number, number> = {};
      let maxDepthFound = 0;

      // 使用栈来构建树形结构 / Use stack to build tree structure
      const stack: OutlineNode[] = [];

      for (const heading of headingParagraphs) {
        const { paragraph, index, level, text, style } = heading;

        // 创建节点 / Create node
        const node: OutlineNode = {
          id: `heading-${index}`,
          text,
          level,
          style,
          children: [],
          index,
        };

        // 添加格式信息 / Add format information
        if (includeFormat) {
          node.format = {
            font: paragraph.font.name,
            fontSize: paragraph.font.size,
            bold: paragraph.font.bold,
            italic: paragraph.font.italic,
            color: paragraph.font.color,
            alignment: paragraph.alignment,
          };
        }

        // 更新统计信息 / Update statistics
        levelCounts[level] = (levelCounts[level] || 0) + 1;
        maxDepthFound = Math.max(maxDepthFound, level);

        // 构建层级关系 / Build hierarchy
        // 弹出栈中所有级别大于等于当前级别的节点
        // Pop all nodes with level >= current level from stack
        while (stack.length > 0 && stack[stack.length - 1].level >= level) {
          stack.pop();
        }

        if (stack.length === 0) {
          // 当前节点是根节点 / Current node is root
          nodes.push(node);
        } else {
          // 当前节点是栈顶节点的子节点 / Current node is child of stack top
          stack[stack.length - 1].children.push(node);
        }

        // 将当前节点入栈 / Push current node to stack
        stack.push(node);
      }

      return {
        nodes,
        totalHeadings: headingParagraphs.length,
        maxDepth: maxDepthFound,
        levelCounts,
      };
    } catch (error) {
      console.error("获取文档大纲失败 / Failed to get document outline:", error);
      throw new Error(`获取文档大纲失败 / Failed to get document outline: ${error.message}`);
    }
  });
}

/**
 * 获取文档大纲的扁平列表（不构建树形结构）
 * Get document outline as flat list (without building tree structure)
 *
 * @param options - 获取选项 / Get options
 * @returns 大纲节点数组 / Array of outline nodes
 */
export async function getDocumentOutlineFlat(
  options: GetDocumentOutlineOptions = {}
): Promise<OutlineNode[]> {
  const { includeFormat = false, maxDepth = 0, specificLevels } = options;

  return Word.run(async (context) => {
    try {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");

      await context.sync();

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("text,style,styleBuiltIn");

        if (includeFormat) {
          para.font.load("name,size,bold,italic,color");
          para.load("alignment");
        }
      }

      await context.sync();

      const outlineNodes: OutlineNode[] = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const style = para.style;

        if (isHeadingStyle(style)) {
          const level = extractHeadingLevel(style);

          if (specificLevels && !specificLevels.includes(level)) {
            continue;
          }
          if (maxDepth > 0 && level > maxDepth) {
            continue;
          }

          const node: OutlineNode = {
            id: `heading-${i}`,
            text: para.text.trim(),
            level,
            style,
            children: [], // 扁平列表中不包含子节点 / No children in flat list
            index: i,
          };

          if (includeFormat) {
            node.format = {
              font: para.font.name,
              fontSize: para.font.size,
              bold: para.font.bold,
              italic: para.font.italic,
              color: para.font.color,
              alignment: para.alignment,
            };
          }

          outlineNodes.push(node);
        }
      }

      return outlineNodes;
    } catch (error) {
      console.error("获取文档大纲失败 / Failed to get document outline:", error);
      throw new Error(`获取文档大纲失败 / Failed to get document outline: ${error.message}`);
    }
  });
}

/**
 * 跳转到指定的大纲节点
 * Navigate to a specific outline node
 *
 * @param nodeId - 节点ID / Node ID
 */
export async function navigateToOutlineNode(nodeId: string): Promise<void> {
  return Word.run(async (context) => {
    try {
      // 从节点ID中提取段落索引 / Extract paragraph index from node ID
      const match = nodeId.match(/^heading-(\d+)$/);
      if (!match) {
        throw new Error("无效的节点ID / Invalid node ID");
      }

      const paragraphIndex = parseInt(match[1], 10);

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");

      await context.sync();

      if (paragraphIndex >= paragraphs.items.length) {
        throw new Error("段落索引超出范围 / Paragraph index out of range");
      }

      const targetParagraph = paragraphs.items[paragraphIndex];

      // 选中该段落 / Select the paragraph
      targetParagraph.select("Start");

      await context.sync();
    } catch (error) {
      console.error("跳转到大纲节点失败 / Failed to navigate to outline node:", error);
      throw new Error(`跳转到大纲节点失败 / Failed to navigate to outline node: ${error.message}`);
    }
  });
}

/**
 * 导出文档大纲为 Markdown 格式
 * Export document outline as Markdown format
 *
 * @param outline - 文档大纲 / Document outline
 * @returns Markdown 格式的大纲文本 / Outline text in Markdown format
 */
export function exportOutlineAsMarkdown(outline: DocumentOutline): string {
  const lines: string[] = [];

  function processNode(node: OutlineNode, depth: number = 0) {
    const indent = "  ".repeat(depth);
    const prefix = "#".repeat(node.level);
    lines.push(`${indent}${prefix} ${node.text}`);

    for (const child of node.children) {
      processNode(child, depth + 1);
    }
  }

  for (const node of outline.nodes) {
    processNode(node);
  }

  return lines.join("\n");
}

/**
 * 导出文档大纲为 JSON 格式
 * Export document outline as JSON format
 *
 * @param outline - 文档大纲 / Document outline
 * @returns JSON 格式的大纲文本 / Outline text in JSON format
 */
export function exportOutlineAsJSON(outline: DocumentOutline): string {
  return JSON.stringify(outline, null, 2);
}
