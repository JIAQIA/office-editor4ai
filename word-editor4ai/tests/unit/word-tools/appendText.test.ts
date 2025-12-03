/**
 * 文件名: appendText.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: appendText工具函数的单元测试
 */

import { describe, it, expect, beforeEach } from "vitest";
import { appendText } from "../../../src/word-tools";

describe("appendText", () => {
  beforeEach(() => {
    // 清理 mock 不需要了，因为使用全局 mock
  });

  it("应该拒绝空参数", async () => {
    await expect(appendText({})).rejects.toThrow("必须提供文本、图片或文件");
  });

  it("应该成功追加文本", async () => {
    const text = "测试文本";
    await expect(appendText({ text })).resolves.not.toThrow();
  });

  it("应该应用文本格式", async () => {
    const options = {
      text: "格式化文本",
      format: {
        fontName: "宋体",
        fontSize: 14,
        bold: true,
        italic: true,
        color: "#FF0000",
      },
    };

    await expect(appendText(options)).resolves.not.toThrow();
  });

  it("应该处理图片插入", async () => {
    const images = [
      {
        base64:
          "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgAF/6vKq1QAAAABJRU5ErkJggg==",
        width: 100,
        height: 100,
        altText: "测试图片",
      },
    ];

    await expect(appendText({ images })).resolves.not.toThrow();
  });
});
