/**
 * 文件名: insertImage.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: insertImage工具函数的单元测试
 */

import { describe, it, expect, beforeEach } from "vitest";
import { insertImage, insertImages } from "../../../src/word-tools";

describe("insertImage", () => {
  // 测试用的 base64 图片数据（1x1 像素的透明 PNG）
  // Test base64 image data (1x1 pixel transparent PNG)
  const testBase64 =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgAF/6vKq1QAAAABJRU5ErkJggg==";

  beforeEach(() => {
    // 清理工作在全局 mock 中处理
    // Cleanup is handled in global mock
  });

  it("应该拒绝空的 base64 数据", async () => {
    const result = await insertImage({ base64: "" });
    expect(result.success).toBe(false);
    expect(result.error).toContain("base64");
  });

  it("应该成功插入内联图片", async () => {
    const result = await insertImage({
      base64: testBase64,
      width: 300,
      height: 200,
      altText: "测试图片",
    });

    expect(result.success).toBe(true);
    // imageId 使用 altTextTitle 作为标识符
    // imageId uses altTextTitle as identifier
    expect(result.imageId).toBe("测试图片");
  });

  it("应该成功插入浮动图片", async () => {
    const result = await insertImage({
      base64: testBase64,
      width: 300,
      height: 200,
      altText: "浮动图片",
      layoutType: "floating",
      floatingOptions: {
        wrapType: "Square",
        position: {
          left: 100,
          top: 100,
        },
      },
    });

    expect(result.success).toBe(true);
    // 浮动图片返回 shape ID / Floating pictures return shape ID
    expect(result.imageId).toBe("shape-mock-shape-id");
  });

  it("应该支持不同的插入位置", async () => {
    const locations: Array<"Start" | "End" | "Before" | "After" | "Replace"> = [
      "Start",
      "End",
      "Before",
      "After",
      "Replace",
    ];

    for (const location of locations) {
      const result = await insertImage({
        base64: testBase64,
        insertLocation: location,
      });

      expect(result.success).toBe(true);
    }
  });

  it("应该应用图片属性", async () => {
    const result = await insertImage({
      base64: testBase64,
      width: 400,
      height: 300,
      altText: "替代文本",
      description: "详细描述",
      hyperlink: "https://example.com",
      keepAspectRatio: true,
    });

    expect(result.success).toBe(true);
  });

  it("应该处理带有数据URL前缀的 base64", async () => {
    const result = await insertImage({
      base64: testBase64, // 包含 data:image/png;base64, 前缀
    });

    expect(result.success).toBe(true);
  });

  it("应该处理不带前缀的纯 base64", async () => {
    const pureBase64 = testBase64.split(",")[1];
    const result = await insertImage({
      base64: pureBase64,
    });

    expect(result.success).toBe(true);
  });

  it("没有 altText 时 imageId 应该为 undefined", async () => {
    const result = await insertImage({
      base64: testBase64,
      width: 300,
      height: 200,
    });

    expect(result.success).toBe(true);
    expect(result.imageId).toBeUndefined();
  });

  it("应该支持浮动图片的不同环绕方式", async () => {
    const wrapTypes: Array<
      "Square" | "Tight" | "Through" | "TopBottom" | "Behind" | "Front"
    > = ["Square", "Tight", "Through", "TopBottom", "Behind", "Front"];

    for (const wrapType of wrapTypes) {
      const result = await insertImage({
        base64: testBase64,
        layoutType: "floating",
        floatingOptions: {
          wrapType,
        },
      });

      expect(result.success).toBe(true);
    }
  });
});

describe("insertImages", () => {
  const testBase64 =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgAF/6vKq1QAAAABJRU5ErkJggg==";

  it("应该批量插入多张图片", async () => {
    const images = [
      {
        base64: testBase64,
        width: 200,
        height: 150,
        altText: "图片1",
      },
      {
        base64: testBase64,
        width: 300,
        height: 200,
        altText: "图片2",
      },
      {
        base64: testBase64,
        width: 400,
        height: 300,
        altText: "图片3",
      },
    ];

    const results = await insertImages(images);

    expect(results).toHaveLength(3);
    expect(results.every((r) => r.success)).toBe(true);
  });

  it("应该处理部分失败的情况", async () => {
    const images = [
      {
        base64: testBase64,
        altText: "有效图片",
      },
      {
        base64: "", // 无效的 base64
        altText: "无效图片",
      },
    ];

    const results = await insertImages(images);

    expect(results).toHaveLength(2);
    expect(results[0].success).toBe(true);
    expect(results[1].success).toBe(false);
  });

  it("应该处理空数组", async () => {
    const results = await insertImages([]);
    expect(results).toHaveLength(0);
  });
});
