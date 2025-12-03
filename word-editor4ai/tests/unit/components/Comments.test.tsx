/**
 * 文件名: Comments.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: Comments 组件的单元测试 | Unit tests for Comments component
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { render, screen, fireEvent, waitFor } from "@testing-library/react";
import Comments from "../../../src/taskpane/components/tools/Comments";
import * as wordTools from "../../../src/word-tools";

// Mock word-tools 模块 / Mock word-tools module
vi.mock("../../../src/word-tools", () => ({
  getComments: vi.fn(),
}));

describe("Comments Component", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("应该正确渲染组件 | Should render component correctly", () => {
    render(<Comments />);

    expect(screen.getByText("获取选项")).toBeInTheDocument();
    expect(screen.getByText("获取批注内容")).toBeInTheDocument();
    expect(screen.getByText("包含已解决的批注")).toBeInTheDocument();
    expect(screen.getByText("包含批注回复")).toBeInTheDocument();
    expect(screen.getByText("包含关联文本")).toBeInTheDocument();
    expect(screen.getByText("详细元数据")).toBeInTheDocument();
  });

  it("应该在点击按钮时调用 getComments | Should call getComments when button is clicked", async () => {
    const mockComments = [
      {
        id: "comment-1",
        content: "Test comment",
        resolved: false,
      },
    ];

    vi.mocked(wordTools.getComments).mockResolvedValue(mockComments);

    render(<Comments />);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(wordTools.getComments).toHaveBeenCalled();
    });
  });

  it("应该显示获取到的批注 | Should display fetched comments", async () => {
    const mockComments = [
      {
        id: "comment-1",
        content: "Test comment 1",
        resolved: false,
      },
      {
        id: "comment-2",
        content: "Test comment 2",
        resolved: true,
      },
    ];

    vi.mocked(wordTools.getComments).mockResolvedValue(mockComments);

    render(<Comments />);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("找到 2 条批注:")).toBeInTheDocument();
      expect(screen.getByText("Test comment 1")).toBeInTheDocument();
      expect(screen.getByText("Test comment 2")).toBeInTheDocument();
    });
  });

  it("应该显示空状态 | Should display empty state", async () => {
    vi.mocked(wordTools.getComments).mockResolvedValue([]);

    render(<Comments />);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("未找到批注")).toBeInTheDocument();
    });
  });

  it("应该显示错误状态 | Should display error state", async () => {
    vi.mocked(wordTools.getComments).mockRejectedValue(new Error("Test error"));

    render(<Comments />);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/错误: Test error/)).toBeInTheDocument();
    });
  });

  it("应该正确传递选项参数 | Should pass options correctly", async () => {
    vi.mocked(wordTools.getComments).mockResolvedValue([]);

    render(<Comments />);

    // 切换选项 / Toggle options
    const includeResolvedSwitch = screen.getByLabelText("包含已解决的批注");
    fireEvent.click(includeResolvedSwitch);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(wordTools.getComments).toHaveBeenCalledWith(
        expect.objectContaining({
          includeResolved: false,
        })
      );
    });
  });

  it("应该显示批注的详细元数据 | Should display detailed metadata", async () => {
    const mockComments = [
      {
        id: "comment-1",
        content: "Test comment",
        resolved: false,
        authorName: "Test Author",
        authorEmail: "test@example.com",
        creationDate: new Date("2025-12-03"),
      },
    ];

    vi.mocked(wordTools.getComments).mockResolvedValue(mockComments);

    render(<Comments />);

    // 启用详细元数据 / Enable detailed metadata
    const detailedMetadataSwitch = screen.getByLabelText("详细元数据");
    fireEvent.click(detailedMetadataSwitch);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("Test Author")).toBeInTheDocument();
      expect(screen.getByText("test@example.com")).toBeInTheDocument();
    });
  });

  it("应该显示批注回复 | Should display comment replies", async () => {
    const mockComments = [
      {
        id: "comment-1",
        content: "Test comment",
        resolved: false,
        replies: [
          {
            id: "reply-1",
            content: "Test reply",
            authorName: "Reply Author",
          },
        ],
      },
    ];

    vi.mocked(wordTools.getComments).mockResolvedValue(mockComments);

    render(<Comments />);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("回复 (1 条):")).toBeInTheDocument();
      expect(screen.getByText(/💬 Test reply/)).toBeInTheDocument();
    });
  });

  it("应该显示关联文本 | Should display associated text", async () => {
    const mockComments = [
      {
        id: "comment-1",
        content: "Test comment",
        resolved: false,
        associatedText: "Associated text content",
      },
    ];

    vi.mocked(wordTools.getComments).mockResolvedValue(mockComments);

    render(<Comments />);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("关联文本:")).toBeInTheDocument();
      expect(screen.getByText("Associated text content")).toBeInTheDocument();
    });
  });

  it("应该显示已解决/未解决的徽章 | Should display resolved/unresolved badge", async () => {
    const mockComments = [
      {
        id: "comment-1",
        content: "Resolved comment",
        resolved: true,
      },
      {
        id: "comment-2",
        content: "Unresolved comment",
        resolved: false,
      },
    ];

    vi.mocked(wordTools.getComments).mockResolvedValue(mockComments);

    render(<Comments />);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("已解决")).toBeInTheDocument();
      expect(screen.getByText("未解决")).toBeInTheDocument();
    });
  });

  it("应该显示 JSON 输出 | Should display JSON output", async () => {
    const mockComments = [
      {
        id: "comment-1",
        content: "Test comment",
        resolved: false,
      },
    ];

    vi.mocked(wordTools.getComments).mockResolvedValue(mockComments);

    render(<Comments />);

    const button = screen.getByText("获取批注内容");
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText("JSON 输出")).toBeInTheDocument();
    });
  });
});
