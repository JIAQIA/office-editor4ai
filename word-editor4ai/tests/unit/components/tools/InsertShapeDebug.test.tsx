/**
 * 文件名: InsertShapeDebug.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: InsertShapeDebug 组件的单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { render, screen, fireEvent, waitFor } from "@testing-library/react";
import { InsertShapeDebug } from "../../../../src/taskpane/components/tools/InsertShapeDebug";
import * as wordTools from "../../../../src/word-tools";

// Mock word-tools
vi.mock("../../../../src/word-tools", () => ({
  insertShape: vi.fn(),
}));

describe("InsertShapeDebug", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("渲染 / Rendering", () => {
    it("应该正确渲染组件 / Should render component correctly", () => {
      render(<InsertShapeDebug />);

      expect(screen.getByText("形状类型")).toBeInTheDocument();
      expect(screen.getByText("基本选项")).toBeInTheDocument();
      expect(screen.getByText("位置和旋转")).toBeInTheDocument();
      expect(screen.getByText("插入形状")).toBeInTheDocument();
    });

    it("应该显示形状类型下拉列表 / Should display shape type dropdown", () => {
      render(<InsertShapeDebug />);

      const dropdown = screen.getByRole("combobox", { name: /形状类型/i });
      expect(dropdown).toBeInTheDocument();
    });

    it("应该显示插入按钮 / Should display insert button", () => {
      render(<InsertShapeDebug />);

      const button = screen.getByRole("button", { name: /插入形状/i });
      expect(button).toBeInTheDocument();
    });
  });

  describe("基本功能 / Basic functionality", () => {
    it("应该成功插入形状 / Should successfully insert shape", async () => {
      const mockInsertShape = vi.mocked(wordTools.insertShape);
      mockInsertShape.mockResolvedValue({
        success: true,
        shapeId: "shape-123",
      });

      render(<InsertShapeDebug />);

      const button = screen.getByRole("button", { name: /插入形状/i });
      fireEvent.click(button);

      await waitFor(() => {
        expect(mockInsertShape).toHaveBeenCalledWith(
          "Rectangle",
          "End",
          expect.objectContaining({
            width: 100,
            height: 100,
          })
        );
      });

      await waitFor(() => {
        expect(screen.getByText(/形状插入成功/i)).toBeInTheDocument();
      });
    });

    it("应该处理插入失败 / Should handle insert failure", async () => {
      const mockInsertShape = vi.mocked(wordTools.insertShape);
      mockInsertShape.mockResolvedValue({
        success: false,
        error: "Insert failed",
      });

      render(<InsertShapeDebug />);

      const button = screen.getByRole("button", { name: /插入形状/i });
      fireEvent.click(button);

      await waitFor(() => {
        expect(screen.getByText(/形状插入失败/i)).toBeInTheDocument();
      });
    });
  });

  describe("样式选项 / Style options", () => {
    it("应该在启用样式时传递样式选项 / Should pass style options when enabled", async () => {
      const mockInsertShape = vi.mocked(wordTools.insertShape);
      mockInsertShape.mockResolvedValue({
        success: true,
        shapeId: "shape-123",
      });

      render(<InsertShapeDebug />);

      // 启用样式设置 / Enable style settings
      const styleSwitch = screen.getByRole("switch", { name: /启用样式设置/i });
      fireEvent.click(styleSwitch);

      // 插入形状 / Insert shape
      const button = screen.getByRole("button", { name: /插入形状/i });
      fireEvent.click(button);

      await waitFor(() => {
        expect(mockInsertShape).toHaveBeenCalledWith(
          "Rectangle",
          "End",
          expect.objectContaining({
            fillColor: expect.any(String),
            lineColor: expect.any(String),
          })
        );
      });
    });
  });

  describe("文本选项 / Text options", () => {
    it("应该在启用文本时传递文本内容 / Should pass text content when enabled", async () => {
      const mockInsertShape = vi.mocked(wordTools.insertShape);
      mockInsertShape.mockResolvedValue({
        success: true,
        shapeId: "shape-123",
      });

      render(<InsertShapeDebug />);

      // 启用文本内容 / Enable text content
      const textSwitch = screen.getByRole("switch", { name: /添加文本内容/i });
      fireEvent.click(textSwitch);

      // 插入形状 / Insert shape
      const button = screen.getByRole("button", { name: /插入形状/i });
      fireEvent.click(button);

      await waitFor(() => {
        expect(mockInsertShape).toHaveBeenCalledWith(
          "Rectangle",
          "End",
          expect.objectContaining({
            text: expect.any(String),
          })
        );
      });
    });
  });

  describe("加载状态 / Loading state", () => {
    it("应该在插入时显示加载状态 / Should show loading state during insert", async () => {
      const mockInsertShape = vi.mocked(wordTools.insertShape);
      mockInsertShape.mockImplementation(
        () => new Promise((resolve) => setTimeout(() => resolve({ success: true }), 100))
      );

      render(<InsertShapeDebug />);

      const button = screen.getByRole("button", { name: /插入形状/i });
      fireEvent.click(button);

      await waitFor(() => {
        expect(screen.getByText(/正在插入/i)).toBeInTheDocument();
      });
    });

    it("应该在加载时禁用按钮 / Should disable button during loading", async () => {
      const mockInsertShape = vi.mocked(wordTools.insertShape);
      mockInsertShape.mockImplementation(
        () => new Promise((resolve) => setTimeout(() => resolve({ success: true }), 100))
      );

      render(<InsertShapeDebug />);

      const button = screen.getByRole("button", { name: /插入形状/i });
      fireEvent.click(button);

      await waitFor(() => {
        expect(button).toBeDisabled();
      });
    });
  });

  describe("错误处理 / Error handling", () => {
    it("应该处理异常错误 / Should handle exception errors", async () => {
      const mockInsertShape = vi.mocked(wordTools.insertShape);
      mockInsertShape.mockRejectedValue(new Error("Unexpected error"));

      render(<InsertShapeDebug />);

      const button = screen.getByRole("button", { name: /插入形状/i });
      fireEvent.click(button);

      await waitFor(() => {
        expect(screen.getByText(/插入形状失败/i)).toBeInTheDocument();
      });
    });
  });
});
