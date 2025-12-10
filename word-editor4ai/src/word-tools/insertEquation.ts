/**
 * 文件名: insertEquation.ts
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: None
 * 描述: 插入公式工具核心逻辑（支持 LaTeX 格式）
 */

/* global Word, console */

import type { InsertLocation } from "./types";

export type { InsertLocation };

/**
 * 插入公式结果 / Insert Equation Result
 */
export interface InsertEquationResult {
  /** 是否成功 / Success */
  success: boolean;
  /** 公式内容（LaTeX 格式）/ Equation content (LaTeX format) */
  latex?: string;
  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;
}

/**
 * 在文档中插入公式（使用 LaTeX 格式）
 * Insert equation in document (using LaTeX format)
 *
 * @remarks
 * - 使用 Word.js API 的 insertOoxml 方法插入 OMML（Office Math Markup Language）格式的公式
 * - 简单的 LaTeX 表达式会被转换为 OMML 格式
 * - 公式以内联方式插入到文档流中
 * - Uses Word.js API's insertOoxml method to insert OMML (Office Math Markup Language) format equations
 * - Simple LaTeX expressions will be converted to OMML format
 * - Equations are inserted inline into the document flow
 *
 * @param latex - LaTeX 格式的公式字符串 / LaTeX format equation string
 * @param location - 插入位置 / Insert location
 * @returns Promise<InsertEquationResult> - 插入结果 / Insert result
 *
 * @example
 * ```typescript
 * // 插入简单公式 / Insert simple equation
 * await insertEquation("E = mc^2", "End");
 *
 * // 插入分数公式 / Insert fraction equation
 * await insertEquation("\\frac{a}{b}", "End");
 *
 * // 插入求和公式 / Insert summation equation
 * await insertEquation("\\sum_{i=1}^{n} x_i", "End");
 * ```
 */
export async function insertEquation(
  latex: string,
  location: InsertLocation = "End"
): Promise<InsertEquationResult> {
  try {
    await Word.run(async (context) => {
      // 获取插入范围 / Get insert range
      let insertRange: Word.Range;
      const selection = context.document.getSelection();

      switch (location) {
        case "Start":
          insertRange = context.document.body.getRange("Start");
          break;
        case "End":
          insertRange = context.document.body.getRange("End");
          break;
        case "Before":
          insertRange = selection;
          break;
        case "After":
          insertRange = selection;
          break;
        case "Replace":
          insertRange = selection;
          break;
        default:
          insertRange = context.document.body.getRange("End");
      }

      // 将 LaTeX 转换为 OMML 格式
      // Convert LaTeX to OMML format
      const ooxml = convertLatexToOoxml(latex);

      // 插入公式 / Insert equation
      const apiInsertLocation:
        | Word.InsertLocation.before
        | Word.InsertLocation.after
        | Word.InsertLocation.replace
        | "Before"
        | "After"
        | "Replace" =
        location === "Start" || location === "Before"
          ? "Before"
          : location === "Replace"
            ? "Replace"
            : "After";

      insertRange.insertOoxml(ooxml, apiInsertLocation);

      await context.sync();
    });

    return {
      success: true,
      latex,
    };
  } catch (error) {
    console.error("插入公式失败 / Failed to insert equation:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 将 LaTeX 格式转换为 OOXML 格式（包含 OMML 数学标记）
 * Convert LaTeX format to OOXML format (containing OMML math markup)
 *
 * @param latex - LaTeX 格式的公式字符串 / LaTeX format equation string
 * @returns OOXML 格式的 XML 字符串 / OOXML format XML string
 */
function convertLatexToOoxml(latex: string): string {
  // 将 LaTeX 转换为 OMML 数学内容
  // Convert LaTeX to OMML math content
  const ommlMath = convertLatexToOmml(latex);

  // 完整的 OOXML 包结构（基于 Office Open XML 标准）
  // Complete OOXML package structure (based on Office Open XML standard)
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body>
          <w:p>
            <w:r>
              ${ommlMath}
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

/**
 * 数学符号映射表 / Math symbol mapping
 */
const MATH_SYMBOLS: Record<string, string> = {
  "\\pm": "±",
  "\\mp": "∓",
  "\\times": "×",
  "\\div": "÷",
  "\\cdot": "⋅",
  "\\leq": "≤",
  "\\geq": "≥",
  "\\neq": "≠",
  "\\approx": "≈",
  "\\infty": "∞",
  "\\to": "→",
  "\\rightarrow": "→",
  "\\leftarrow": "←",
  "\\leftrightarrow": "↔",
  "\\Rightarrow": "⇒",
  "\\Leftarrow": "⇐",
  "\\alpha": "α",
  "\\beta": "β",
  "\\gamma": "γ",
  "\\delta": "δ",
  "\\theta": "θ",
  "\\pi": "π",
  "\\sigma": "σ",
  "\\omega": "ω",
};

/**
 * 提取大括号内的内容（支持嵌套）
 * Extract content within braces (supports nesting)
 */
function extractBracedContent(
  str: string,
  startIndex: number
): { content: string; endIndex: number } | null {
  if (str[startIndex] !== "{") return null;

  let braceCount = 1;
  let i = startIndex + 1;

  while (i < str.length && braceCount > 0) {
    if (str[i] === "\\") {
      i += 2; // 跳过转义字符
      continue;
    }
    if (str[i] === "{") braceCount++;
    if (str[i] === "}") braceCount--;
    i++;
  }

  if (braceCount !== 0) return null; // 不匹配

  return {
    content: str.substring(startIndex + 1, i - 1),
    endIndex: i - 1,
  };
}

/**
 * 将 LaTeX 表达式转换为 OMML（Office Math Markup Language）格式
 * Convert LaTeX expression to OMML (Office Math Markup Language) format
 *
 * @param latex - LaTeX 表达式 / LaTeX expression
 * @returns OMML 格式的数学标记 / OMML format math markup
 */
function convertLatexToOmml(latex: string): string {
  // 使用占位符标记特殊结构，避免正则表达式冲突
  // Use placeholders to mark special structures to avoid regex conflicts
  // 使用特殊字符避免与下标语法冲突 / Use special characters to avoid conflicts with subscript syntax
  const placeholders: string[] = [];
  let result = latex;

  console.log("原始输入 / Original input:", latex);

  // 先替换数学符号 / Replace math symbols first
  for (const [latexSymbol, unicodeSymbol] of Object.entries(MATH_SYMBOLS)) {
    result = result.replace(new RegExp(latexSymbol.replace(/\\/g, "\\\\"), "g"), unicodeSymbol);
  }

  // 处理极限符号 \lim_{x \to \infty} f(x) / Handle limit \lim_{x \to \infty} f(x)
  result = result.replace(/\\lim_\{([^}]+)\}/g, (_, condition) => {
    const placeholder = `§PLACEHOLDER§${placeholders.length}§`;
    placeholders.push(
      `<m:sSub>` +
        `<m:e><m:r><m:t>lim</m:t></m:r></m:e>` +
        `<m:sub><m:r><m:t>${condition}</m:t></m:r></m:sub>` +
        `</m:sSub>`
    );
    console.log("极限替换 / Limit replaced:", placeholder);
    return placeholder;
  });

  // 处理求和符号 \sum_{i=1}^{n} x_i / Handle summation \sum_{i=1}^{n} x_i
  result = result.replace(/\\sum_\{([^}]+)\}\^\{([^}]+)\}\s*([^ ]+)/g, (_, lower, upper, body) => {
    const placeholder = `§PLACEHOLDER§${placeholders.length}§`;
    placeholders.push(
      `<m:nary>` +
        `<m:naryPr>` +
        `<m:chr m:val="∑"/>` +
        `<m:limLoc m:val="undOvr"/>` +
        `</m:naryPr>` +
        `<m:sub><m:r><m:t>${lower}</m:t></m:r></m:sub>` +
        `<m:sup><m:r><m:t>${upper}</m:t></m:r></m:sup>` +
        `<m:e><m:r><m:t>${body}</m:t></m:r></m:e>` +
        `</m:nary>`
    );
    console.log("求和替换 / Summation replaced:", placeholder);
    return placeholder;
  });

  // 处理积分符号 \int_a^{b} f(x) dx / Handle integral \int_a^{b} f(x) dx
  result = result.replace(/\\int_\{([^}]+)\}\^\{([^}]+)\}\s*([^ ]+)/g, (_, lower, upper, body) => {
    const placeholder = `§PLACEHOLDER§${placeholders.length}§`;
    placeholders.push(
      `<m:nary>` +
        `<m:naryPr>` +
        `<m:chr m:val="∫"/>` +
        `<m:limLoc m:val="undOvr"/>` +
        `</m:naryPr>` +
        `<m:sub><m:r><m:t>${lower}</m:t></m:r></m:sub>` +
        `<m:sup><m:r><m:t>${upper}</m:t></m:r></m:sup>` +
        `<m:e><m:r><m:t>${body}</m:t></m:r></m:e>` +
        `</m:nary>`
    );
    console.log("积分替换 / Integral replaced:", placeholder);
    return placeholder;
  });

  // 处理分数 \frac{}{} / Handle fraction \frac{}{}
  // 需要递归处理，因为分子分母可能包含其他结构
  let i = 0;
  while (i < result.length) {
    if (result.substring(i, i + 6) === "\\frac{") {
      const numResult = extractBracedContent(result, i + 5);
      if (numResult) {
        const denResult = extractBracedContent(result, numResult.endIndex + 1);
        if (denResult) {
          // 递归处理分子和分母
          const numOmml = convertLatexToOmml(numResult.content);
          const denOmml = convertLatexToOmml(denResult.content);

          const placeholder = `§PLACEHOLDER§${placeholders.length}§`;
          placeholders.push(
            `<m:f>` +
              `<m:num>${numOmml.replace(/<m:oMath[^>]*>|<\/m:oMath>/g, "")}</m:num>` +
              `<m:den>${denOmml.replace(/<m:oMath[^>]*>|<\/m:oMath>/g, "")}</m:den>` +
              `</m:f>`
          );
          console.log("分数替换 / Fraction replaced:", placeholder);

          // 替换整个 \frac{...}{...}
          result = result.substring(0, i) + placeholder + result.substring(denResult.endIndex + 2);
          i += placeholder.length;
          continue;
        }
      }
    }
    i++;
  }

  // 处理根号 \sqrt{} / Handle square root \sqrt{}
  // 需要递归处理，因为根号内可能包含其他结构
  i = 0;
  while (i < result.length) {
    if (result.substring(i, i + 6) === "\\sqrt{") {
      const contentResult = extractBracedContent(result, i + 5);
      if (contentResult) {
        // 递归处理根号内容
        const contentOmml = convertLatexToOmml(contentResult.content);

        const placeholder = `§PLACEHOLDER§${placeholders.length}§`;
        placeholders.push(
          `<m:rad>` +
            `<m:radPr><m:degHide m:val="1"/></m:radPr>` +
            `<m:deg/>` +
            `<m:e>${contentOmml.replace(/<m:oMath[^>]*>|<\/m:oMath>/g, "")}</m:e>` +
            `</m:rad>`
        );
        console.log("根号替换 / Sqrt replaced:", placeholder);

        // 替换整个 \sqrt{...}
        result =
          result.substring(0, i) + placeholder + result.substring(contentResult.endIndex + 2);
        i += placeholder.length;
        continue;
      }
    }
    i++;
  }

  // 处理上标 ^{} / Handle superscript ^{}
  result = result.replace(/(\w)\^{([^}]+)}/g, (_, base, sup) => {
    const placeholder = `§PLACEHOLDER§${placeholders.length}§`;
    placeholders.push(
      `<m:sSup><m:e><m:r><m:t>${base}</m:t></m:r></m:e><m:sup><m:r><m:t>${sup}</m:t></m:r></m:sup></m:sSup>`
    );
    console.log("上标{}替换 / Superscript{} replaced:", placeholder);
    return placeholder;
  });

  // 处理简单上标 ^x / Handle simple superscript ^x
  result = result.replace(/(\w)\^(\w)/g, (_, base, sup) => {
    const placeholder = `§PLACEHOLDER§${placeholders.length}§`;
    placeholders.push(
      `<m:sSup><m:e><m:r><m:t>${base}</m:t></m:r></m:e><m:sup><m:r><m:t>${sup}</m:t></m:r></m:sup></m:sSup>`
    );
    console.log("简单上标替换 / Simple superscript replaced:", base, "^", sup, "->", placeholder);
    return placeholder;
  });

  console.log("替换后的字符串 / After replacement:", result);
  console.log("占位符数组 / Placeholders array:", placeholders);

  // 处理下标 _{} / Handle subscript _{}
  result = result.replace(/(\w)_{([^}]+)}/g, (_, base, sub) => {
    const placeholder = `§PLACEHOLDER§${placeholders.length}§`;
    placeholders.push(
      `<m:sSub><m:e><m:r><m:t>${base}</m:t></m:r></m:e><m:sub><m:r><m:t>${sub}</m:t></m:r></m:sub></m:sSub>`
    );
    return placeholder;
  });

  // 处理简单下标 _x / Handle simple subscript _x
  result = result.replace(/(\w)_(\w)/g, (_, base, sub) => {
    const placeholder = `§PLACEHOLDER§${placeholders.length}§`;
    placeholders.push(
      `<m:sSub><m:e><m:r><m:t>${base}</m:t></m:r></m:e><m:sub><m:r><m:t>${sub}</m:t></m:r></m:sub></m:sSub>`
    );
    return placeholder;
  });

  // 将占位符替换回 OMML 标记，并包装剩余文本
  // Replace placeholders with OMML tags and wrap remaining text
  console.log("开始处理占位符替换 / Start processing placeholder replacement");
  console.log("结果字符串 / Result string:", result);
  console.log("占位符数量 / Placeholder count:", placeholders.length);

  // 使用正则表达式分割字符串，保留占位符
  const parts = result.split(/(§PLACEHOLDER§\d+§)/);
  console.log("分割后的部分 / Split parts:", parts);

  const ommlParts: string[] = [];

  for (const part of parts) {
    if (!part) continue; // 跳过空字符串

    // 检查是否是占位符
    const match = part.match(/^§PLACEHOLDER§(\d+)§$/);
    if (match) {
      const placeholderIndex = parseInt(match[1]);
      console.log("替换占位符 / Replace placeholder:", placeholderIndex);
      ommlParts.push(placeholders[placeholderIndex]);
    } else {
      // 普通文本
      ommlParts.push(`<m:r><m:t>${part}</m:t></m:r>`);
    }
  }

  // 组合所有部分
  const ommlContent = ommlParts.join("");

  console.log("OMML 部分 / OMML parts:", ommlParts);
  console.log("最终 OMML 内容 / Final OMML content:", ommlContent);

  return `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${ommlContent}</m:oMath>`;
}
