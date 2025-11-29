/**
 * 文件名: imageInsertion.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 描述: imageInsertion 工具的单元测试 | Unit tests for imageInsertion tool
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import {
  insertImageToSlide,
  insertImage,
  readImageAsBase64,
  fetchImageAsBase64,
} from "../../../src/ppt-tools";

// Mock Office.context.document
const createMockOfficeContext = () => {
  const mockAsyncResult: Office.AsyncResult<void> = {
    status: Office.AsyncResultStatus.Succeeded,
    value: undefined,
    error: undefined,
    asyncContext: undefined,
    diagnostics: undefined,
  };

  return {
    setSelectedDataAsync: vi.fn(
      (
        _data: string,
        _options: Office.SetSelectedDataOptions,
        callback: (result: Office.AsyncResult<void>) => void
      ) => {
        callback(mockAsyncResult);
      }
    ),
  };
};

describe("imageInsertion 工具测试 | imageInsertion Tool Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();

    // Mock Office.AsyncResultStatus first (before createMockOfficeContext uses it)
    if (!global.Office) {
      global.Office = {} as typeof Office;
    }
    if (!global.Office.AsyncResultStatus) {
      global.Office.AsyncResultStatus = {
        Succeeded: 0,
        Failed: 1,
      } as typeof Office.AsyncResultStatus;
    }

    // Mock Office.context.document
    if (!global.Office.context) {
      global.Office.context = {} as Office.Context;
    }
    global.Office.context.document = createMockOfficeContext() as unknown as Office.Document;

    // Mock Office.CoercionType
    if (!global.Office.CoercionType) {
      global.Office.CoercionType = {
        Image: "image",
      } as unknown as typeof Office.CoercionType;
    }
  });

  describe("insertImageToSlide", () => {
    it("应该能够插入 Base64 图片（带 data URL 前缀）| should insert Base64 image with data URL prefix", async () => {
      const base64WithPrefix =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

      const result = await insertImageToSlide({
        imageSource: base64WithPrefix,
        left: 100,
        top: 100,
        width: 200,
        height: 150,
      });

      expect(global.Office.context.document.setSelectedDataAsync).toHaveBeenCalledTimes(1);
      expect(result).toEqual({
        imageId: "",
        width: 200,
        height: 150,
      });
    });

    it("应该能够插入纯 Base64 图片（不带前缀）| should insert pure Base64 image without prefix", async () => {
      const pureBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

      const result = await insertImageToSlide({
        imageSource: pureBase64,
      });

      expect(global.Office.context.document.setSelectedDataAsync).toHaveBeenCalledTimes(1);
      expect(result).toEqual({
        imageId: "",
        width: 200,
        height: 150,
      });
    });

    it("应该能够插入带有指定位置的图片 | should insert image with specified position", async () => {
      const base64Data =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

      await insertImageToSlide({
        imageSource: base64Data,
        left: 150,
        top: 250,
      });

      const mockFn = vi.mocked(global.Office.context.document.setSelectedDataAsync);
      expect(mockFn).toHaveBeenCalledWith(
        base64Data,
        expect.objectContaining({
          coercionType: "image",
          imageLeft: 150,
          imageTop: 250,
        }),
        expect.any(Function)
      );
    });

    it("应该能够插入带有自定义尺寸的图片 | should insert image with custom dimensions", async () => {
      const base64Data =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

      await insertImageToSlide({
        imageSource: base64Data,
        width: 300,
        height: 200,
      });

      const mockFn = vi.mocked(global.Office.context.document.setSelectedDataAsync);
      expect(mockFn).toHaveBeenCalledWith(
        base64Data,
        expect.objectContaining({
          coercionType: "image",
          imageWidth: 300,
          imageHeight: 200,
        }),
        expect.any(Function)
      );
    });

    it("应该能够插入带有完整配置的图片 | should insert image with full configuration", async () => {
      const base64Data =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

      const result = await insertImageToSlide({
        imageSource: base64Data,
        left: 100,
        top: 200,
        width: 400,
        height: 300,
      });

      const mockFn = vi.mocked(global.Office.context.document.setSelectedDataAsync);
      expect(mockFn).toHaveBeenCalledWith(
        base64Data,
        expect.objectContaining({
          coercionType: "image",
          imageLeft: 100,
          imageTop: 200,
          imageWidth: 400,
          imageHeight: 300,
        }),
        expect.any(Function)
      );

      expect(result).toEqual({
        imageId: "",
        width: 400,
        height: 300,
      });
    });

    it("应该正确提取 data URL 前缀中的 Base64 数据 | should correctly extract Base64 data from data URL prefix", async () => {
      const base64WithPrefix = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBD";

      await insertImageToSlide({
        imageSource: base64WithPrefix,
      });

      const mockFn = vi.mocked(global.Office.context.document.setSelectedDataAsync);
      const calledData = mockFn.mock.calls[0][0];

      // 应该只传递纯 Base64 部分 | Should only pass pure Base64 part
      expect(calledData).toBe("/9j/4AAQSkZJRgABAQEAYABgAAD/2wBD");
      expect(calledData).not.toContain("data:");
    });

    it("应该在插入失败时抛出错误 | should throw error when insertion fails", async () => {
      const mockErrorResult: Office.AsyncResult<void> = {
        status: Office.AsyncResultStatus.Failed,
        value: undefined,
        error: {
          message: "插入图片失败",
          name: "Error",
          code: 0,
        },
        asyncContext: undefined,
        diagnostics: undefined,
      };

      global.Office.context.document.setSelectedDataAsync = vi.fn((_data, _options, callback) => {
        callback(mockErrorResult);
      }) as unknown as Office.Document["setSelectedDataAsync"];

      await expect(
        insertImageToSlide({
          imageSource: "invalid-base64",
        })
      ).rejects.toThrow("插入图片失败");
    });

    it("应该返回默认尺寸当未指定尺寸时 | should return default dimensions when not specified", async () => {
      const result = await insertImageToSlide({
        imageSource: "base64data",
      });

      expect(result.width).toBe(200);
      expect(result.height).toBe(150);
    });

    it("应该处理零坐标 | should handle zero coordinates", async () => {
      await insertImageToSlide({
        imageSource: "base64data",
        left: 0,
        top: 0,
      });

      const mockFn = vi.mocked(global.Office.context.document.setSelectedDataAsync);
      expect(mockFn).toHaveBeenCalledWith(
        "base64data",
        expect.objectContaining({
          imageLeft: 0,
          imageTop: 0,
        }),
        expect.any(Function)
      );
    });
  });

  describe("insertImage", () => {
    it("应该能够插入简单图片 | should insert simple image", async () => {
      const base64Data =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

      const result = await insertImage(base64Data);

      expect(global.Office.context.document.setSelectedDataAsync).toHaveBeenCalledTimes(1);
      expect(result).toEqual({
        imageId: "",
        width: 200,
        height: 150,
      });
    });

    it("应该能够插入带有位置的图片 | should insert image with position", async () => {
      const base64Data = "base64string";

      await insertImage(base64Data, 100, 200);

      const mockFn = vi.mocked(global.Office.context.document.setSelectedDataAsync);
      expect(mockFn).toHaveBeenCalledWith(
        base64Data,
        expect.objectContaining({
          imageLeft: 100,
          imageTop: 200,
        }),
        expect.any(Function)
      );
    });

    it("应该正确调用 insertImageToSlide | should correctly call insertImageToSlide", async () => {
      const base64Data = "testbase64";
      const left = 50;
      const top = 75;

      const result = await insertImage(base64Data, left, top);

      expect(result).toEqual({
        imageId: "",
        width: 200,
        height: 150,
      });
    });
  });

  describe("readImageAsBase64", () => {
    it("应该能够读取文件并转换为 Base64 | should read file and convert to Base64", async () => {
      const mockFile = new File(["test content"], "test.png", { type: "image/png" });
      const expectedBase64 = "data:image/png;base64,dGVzdCBjb250ZW50";

      // Mock FileReader
      class MockFileReader {
        result: string | null = "";
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;

        readAsDataURL = vi.fn(function (this: MockFileReader) {
          this.result = expectedBase64;
          setTimeout(() => this.onload?.(), 0);
        });
      }

      global.FileReader = MockFileReader as unknown as typeof FileReader;

      const result = await readImageAsBase64(mockFile);

      expect(result).toBe(expectedBase64);
    });

    it("应该在读取失败时抛出错误 | should throw error when reading fails", async () => {
      const mockFile = new File(["test"], "test.png", { type: "image/png" });

      class MockFileReader {
        result: string | null = "";
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;

        readAsDataURL = vi.fn(function (this: MockFileReader) {
          setTimeout(() => this.onerror?.(), 0);
        });
      }

      global.FileReader = MockFileReader as unknown as typeof FileReader;

      await expect(readImageAsBase64(mockFile)).rejects.toThrow("读取文件失败");
    });

    it("应该在结果不是字符串时抛出错误 | should throw error when result is not string", async () => {
      const mockFile = new File(["test"], "test.png", { type: "image/png" });

      class MockFileReader {
        result: string | ArrayBuffer | null = null;
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;

        readAsDataURL = vi.fn(function (this: MockFileReader) {
          this.result = null;
          setTimeout(() => this.onload?.(), 0);
        });
      }

      global.FileReader = MockFileReader as unknown as typeof FileReader;

      await expect(readImageAsBase64(mockFile)).rejects.toThrow("读取文件失败：结果不是字符串");
    });
  });

  describe("fetchImageAsBase64", () => {
    it("应该能够从 URL 加载图片并转换为 Base64 | should fetch image from URL and convert to Base64", async () => {
      const imageUrl = "https://example.com/image.png";
      const mockBlob = new Blob(["image data"], { type: "image/png" });
      const expectedBase64 = "data:image/png;base64,aW1hZ2UgZGF0YQ==";

      // Mock fetch
      global.fetch = vi.fn().mockResolvedValue({
        ok: true,
        blob: () => Promise.resolve(mockBlob),
      });

      // Mock FileReader
      class MockFileReader {
        result: string | null = "";
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;

        readAsDataURL = vi.fn(function (this: MockFileReader) {
          this.result = expectedBase64;
          setTimeout(() => this.onload?.(), 0);
        });
      }

      global.FileReader = MockFileReader as unknown as typeof FileReader;

      const result = await fetchImageAsBase64(imageUrl);

      expect(result).toBe(expectedBase64);
      expect(global.fetch).toHaveBeenCalledWith(imageUrl);
    });

    it("应该在 fetch 失败时抛出错误 | should throw error when fetch fails", async () => {
      const imageUrl = "https://example.com/image.png";

      global.fetch = vi.fn().mockResolvedValue({
        ok: false,
        status: 404,
        statusText: "Not Found",
      });

      await expect(fetchImageAsBase64(imageUrl)).rejects.toThrow("获取图片失败: 404 Not Found");
    });

    it("应该在网络错误时抛出错误 | should throw error on network error", async () => {
      const imageUrl = "https://example.com/image.png";

      global.fetch = vi.fn().mockRejectedValue(new Error("Network error"));

      await expect(fetchImageAsBase64(imageUrl)).rejects.toThrow(
        "无法从 URL 加载图片: Network error"
      );
    });

    it("应该在 FileReader 失败时抛出错误 | should throw error when FileReader fails", async () => {
      const imageUrl = "https://example.com/image.png";
      const mockBlob = new Blob(["image data"], { type: "image/png" });

      global.fetch = vi.fn().mockResolvedValue({
        ok: true,
        blob: () => Promise.resolve(mockBlob),
      });

      class MockFileReader {
        result: string | null = "";
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;

        readAsDataURL = vi.fn(function (this: MockFileReader) {
          setTimeout(() => this.onerror?.(), 0);
        });
      }

      global.FileReader = MockFileReader as unknown as typeof FileReader;

      await expect(fetchImageAsBase64(imageUrl)).rejects.toThrow("转换失败");
    });
  });

  describe("边界情况测试 | Edge Cases", () => {
    it("应该处理空 Base64 字符串 | should handle empty Base64 string", async () => {
      const result = await insertImageToSlide({
        imageSource: "",
      });

      expect(result).toEqual({
        imageId: "",
        width: 200,
        height: 150,
      });
    });

    it("应该处理非常大的尺寸值 | should handle very large dimension values", async () => {
      const result = await insertImageToSlide({
        imageSource: "base64data",
        width: 10000,
        height: 10000,
      });

      expect(result.width).toBe(10000);
      expect(result.height).toBe(10000);
    });

    it("应该处理负数尺寸（虽然不推荐）| should handle negative dimensions (not recommended)", async () => {
      await insertImageToSlide({
        imageSource: "base64data",
        left: -10,
        top: -20,
        width: -100,
        height: -50,
      });

      const mockFn = vi.mocked(global.Office.context.document.setSelectedDataAsync);
      expect(mockFn).toHaveBeenCalledWith(
        "base64data",
        expect.objectContaining({
          imageLeft: -10,
          imageTop: -20,
          imageWidth: -100,
          imageHeight: -50,
        }),
        expect.any(Function)
      );
    });

    it("应该处理包含多个逗号的 data URL | should handle data URL with multiple commas", async () => {
      const complexDataUrl = "data:image/png;base64,abc,def,ghi";

      await insertImageToSlide({
        imageSource: complexDataUrl,
      });

      const mockFn = vi.mocked(global.Office.context.document.setSelectedDataAsync);
      const calledData = mockFn.mock.calls[0][0];

      // split(",")[1] 只会取第一个逗号后的第一个部分 | split(",")[1] only takes the first part after first comma
      expect(calledData).toBe("abc");
    });
  });
});
