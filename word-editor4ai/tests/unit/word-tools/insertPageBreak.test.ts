/**
 * 文件名: insertPageBreak.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: insertPageBreak工具函数的单元测试 / Unit tests for insertPageBreak tool function
 */

import { describe, it, expect, beforeEach } from "vitest";
import { insertPageBreak } from "../../../src/word-tools";

describe("insertPageBreak", () => {
  beforeEach(() => {
    // 清理工作在全局 mock 中处理
    // Cleanup is handled in global mock
  });

  it("应该在文档末尾成功插入分页符", async () => {
    const result = await insertPageBreak("End");
    expect(result.success).toBe(true);
    expect(result.error).toBeUndefined();
  });

  it("应该在文档开头成功插入分页符", async () => {
    const result = await insertPageBreak("Start");
    expect(result.success).toBe(true);
    expect(result.error).toBeUndefined();
  });

  it("应该在选中内容之前插入分页符", async () => {
    const result = await insertPageBreak("Before");
    expect(result.success).toBe(true);
    expect(result.error).toBeUndefined();
  });

  it("应该在选中内容之后插入分页符", async () => {
    const result = await insertPageBreak("After");
    expect(result.success).toBe(true);
    expect(result.error).toBeUndefined();
  });

  it("应该替换选中内容为分页符", async () => {
    const result = await insertPageBreak("Replace");
    expect(result.success).toBe(true);
    expect(result.error).toBeUndefined();
  });

  it("应该支持所有插入位置", async () => {
    const locations: Array<"Start" | "End" | "Before" | "After" | "Replace"> = [
      "Start",
      "End",
      "Before",
      "After",
      "Replace",
    ];

    for (const location of locations) {
      const result = await insertPageBreak(location);
      expect(result.success).toBe(true);
      expect(result.error).toBeUndefined();
    }
  });
});
