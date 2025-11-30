/**
 * 文件名: SlideDeletion.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片删除调试组件
 */

import React, { useState, useEffect } from "react";
import { deleteCurrentSlide, deleteSlidesByNumbers } from "../../../ppt-tools";

export const SlideDeletion: React.FC = () => {
  const [slideNumbers, setSlideNumbers] = useState<string>("");
  const [totalSlides, setTotalSlides] = useState<number>(0);
  const [currentSlideNumber, setCurrentSlideNumber] = useState<number>(0);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<string>("");
  const [messageType, setMessageType] = useState<"success" | "error" | "info">("info");
  const [deleteDetails, setDeleteDetails] = useState<{
    deleted: number[];
    notFound: number[];
    errors: Array<{ slideNumber: number; error: string }>;
  } | null>(null);

  // 获取幻灯片总数和当前页码
  const fetchSlideInfo = async () => {
    try {
      /* global PowerPoint */
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");

        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");

        await context.sync();

        setTotalSlides(slides.items.length);

        if (selectedSlides.items.length > 0) {
          const currentSlide = selectedSlides.items[0];
          currentSlide.load("id");
          await context.sync();

          // 找到当前幻灯片的索引
          for (let i = 0; i < slides.items.length; i++) {
            slides.items[i].load("id");
          }
          await context.sync();

          for (let i = 0; i < slides.items.length; i++) {
            if (slides.items[i].id === currentSlide.id) {
              setCurrentSlideNumber(i + 1);
              break;
            }
          }
        }
      });
    } catch {
      // 获取幻灯片信息失败
    }
  };

  // 组件加载时获取幻灯片信息
  useEffect(() => {
    fetchSlideInfo();
  }, []);

  // 删除当前幻灯片
  const handleDeleteCurrentSlide = async () => {
    setLoading(true);
    setMessage("");
    setDeleteDetails(null);

    try {
      const result = await deleteCurrentSlide();

      if (result.success) {
        setMessage(result.message);
        setMessageType("success");
        setDeleteDetails(result.details || null);
        // 刷新幻灯片信息
        await fetchSlideInfo();
      } else {
        setMessage(result.message);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`删除失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 删除指定页码的幻灯片
  const handleDeleteByNumbers = async () => {
    if (!slideNumbers.trim()) {
      setMessage("请输入要删除的页码");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    setDeleteDetails(null);

    try {
      // 解析页码列表（支持逗号、空格、换行符分隔）
      const numbers = slideNumbers
        .split(/[,\s\n]+/)
        .map((num) => num.trim())
        .filter((num) => num.length > 0)
        .map((num) => parseInt(num, 10))
        .filter((num) => !isNaN(num));

      if (numbers.length === 0) {
        setMessage("请输入有效的页码");
        setMessageType("error");
        setLoading(false);
        return;
      }

      const result = await deleteSlidesByNumbers(numbers);

      if (result.success) {
        setMessage(result.message);
        setMessageType("success");
        setDeleteDetails(result.details || null);
        setSlideNumbers("");
        // 刷新幻灯片信息
        await fetchSlideInfo();
      } else {
        setMessage(result.message);
        setMessageType(result.deletedCount > 0 ? "info" : "error");
        setDeleteDetails(result.details || null);
      }
    } catch (error) {
      setMessage(`删除失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 快速选择页码
  const handleQuickSelect = (pageNumber: number) => {
    const currentNumbers = slideNumbers
      .split(/[,\s\n]+/)
      .map((num) => num.trim())
      .filter((num) => num.length > 0);

    if (currentNumbers.includes(pageNumber.toString())) {
      // 如果已存在，则移除
      const newNumbers = currentNumbers.filter((num) => num !== pageNumber.toString());
      setSlideNumbers(newNumbers.join(", "));
    } else {
      // 如果不存在，则添加
      const newNumbers = [...currentNumbers, pageNumber.toString()];
      setSlideNumbers(newNumbers.join(", "));
    }
  };

  return (
    <div style={{ padding: "16px" }}>
      <h3 style={{ marginTop: 0, marginBottom: "16px", fontSize: "16px", fontWeight: 600 }}>
        幻灯片删除调试工具
      </h3>

      {/* 幻灯片信息 */}
      <div
        style={{
          padding: "12px",
          marginBottom: "16px",
          backgroundColor: "#f5f5f5",
          borderRadius: "4px",
          fontSize: "14px",
        }}
      >
        <div style={{ marginBottom: "4px" }}>
          <strong>总页数:</strong> {totalSlides} 页
        </div>
        <div>
          <strong>当前页:</strong>{" "}
          {currentSlideNumber > 0 ? `第 ${currentSlideNumber} 页` : "未选中"}
        </div>
        <button
          onClick={fetchSlideInfo}
          disabled={loading}
          style={{
            marginTop: "8px",
            padding: "4px 12px",
            backgroundColor: "#0078d4",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading ? "not-allowed" : "pointer",
            fontSize: "12px",
          }}
        >
          刷新信息
        </button>
      </div>

      {/* 删除当前页按钮 */}
      <div style={{ marginBottom: "16px" }}>
        <button
          onClick={handleDeleteCurrentSlide}
          disabled={loading || currentSlideNumber === 0}
          style={{
            width: "100%",
            padding: "10px 16px",
            backgroundColor: "#d13438",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading || currentSlideNumber === 0 ? "not-allowed" : "pointer",
            fontSize: "14px",
            fontWeight: 600,
          }}
        >
          {loading ? "删除中..." : `删除当前页 (第 ${currentSlideNumber} 页)`}
        </button>
      </div>

      {/* 页码输入区域 */}
      <div style={{ marginBottom: "16px" }}>
        <label
          style={{
            display: "block",
            marginBottom: "8px",
            fontSize: "14px",
            fontWeight: 500,
          }}
        >
          指定页码删除（支持多个，用逗号分隔）:
        </label>
        <textarea
          value={slideNumbers}
          onChange={(e) => setSlideNumbers(e.target.value)}
          placeholder="输入页码，例如: 1, 3, 5"
          rows={3}
          style={{
            width: "100%",
            padding: "8px",
            border: "1px solid #ccc",
            borderRadius: "4px",
            fontSize: "14px",
            boxSizing: "border-box",
            fontFamily: "monospace",
            resize: "vertical",
          }}
        />
        <button
          onClick={handleDeleteByNumbers}
          disabled={loading || !slideNumbers.trim()}
          style={{
            width: "100%",
            padding: "8px 16px",
            backgroundColor: "#d13438",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading || !slideNumbers.trim() ? "not-allowed" : "pointer",
            fontSize: "14px",
            marginTop: "8px",
          }}
        >
          删除指定页码
        </button>
      </div>

      {/* 快速选择页码 */}
      {totalSlides > 0 && (
        <div style={{ marginBottom: "16px" }}>
          <label
            style={{
              display: "block",
              marginBottom: "8px",
              fontSize: "14px",
              fontWeight: 500,
            }}
          >
            快速选择页码:
          </label>
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: "6px",
            }}
          >
            {Array.from({ length: Math.min(totalSlides, 20) }, (_, i) => i + 1).map((pageNum) => {
              const isSelected = slideNumbers
                .split(/[,\s\n]+/)
                .map((num) => num.trim())
                .includes(pageNum.toString());

              return (
                <button
                  key={pageNum}
                  onClick={() => handleQuickSelect(pageNum)}
                  disabled={loading}
                  style={{
                    padding: "6px 12px",
                    backgroundColor: isSelected ? "#0078d4" : "#f5f5f5",
                    color: isSelected ? "white" : "#333",
                    border: `1px solid ${isSelected ? "#0078d4" : "#ccc"}`,
                    borderRadius: "4px",
                    cursor: loading ? "not-allowed" : "pointer",
                    fontSize: "12px",
                    minWidth: "40px",
                  }}
                >
                  {pageNum}
                </button>
              );
            })}
            {totalSlides > 20 && (
              <span style={{ padding: "6px 12px", fontSize: "12px", color: "#666" }}>
                ... 共 {totalSlides} 页
              </span>
            )}
          </div>
        </div>
      )}

      {/* 消息提示 */}
      {message && (
        <div
          style={{
            padding: "12px",
            marginBottom: "16px",
            borderRadius: "4px",
            fontSize: "14px",
            backgroundColor:
              messageType === "success"
                ? "#dff6dd"
                : messageType === "error"
                  ? "#fde7e9"
                  : "#e1f5fe",
            color:
              messageType === "success"
                ? "#107c10"
                : messageType === "error"
                  ? "#a80000"
                  : "#014361",
            border: `1px solid ${
              messageType === "success"
                ? "#107c10"
                : messageType === "error"
                  ? "#a80000"
                  : "#014361"
            }`,
          }}
        >
          {message}
        </div>
      )}

      {/* 删除详情 */}
      {deleteDetails && (
        <div
          style={{
            padding: "12px",
            marginBottom: "16px",
            backgroundColor: "#f5f5f5",
            borderRadius: "4px",
            fontSize: "13px",
          }}
        >
          <strong>删除详情:</strong>
          {deleteDetails.deleted.length > 0 && (
            <div style={{ marginTop: "8px", color: "#107c10" }}>
              ✓ 成功删除: {deleteDetails.deleted.join(", ")}
            </div>
          )}
          {deleteDetails.notFound.length > 0 && (
            <div style={{ marginTop: "8px", color: "#f59e0b" }}>
              ⚠ 页码不存在: {deleteDetails.notFound.join(", ")}
            </div>
          )}
          {deleteDetails.errors.length > 0 && (
            <div style={{ marginTop: "8px", color: "#a80000" }}>
              ✗ 删除失败:
              <ul style={{ margin: "4px 0 0 20px", paddingLeft: 0 }}>
                {deleteDetails.errors.map((err, idx) => (
                  <li key={idx}>
                    第 {err.slideNumber} 页: {err.error}
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}

      {/* 使用说明 */}
      <div
        style={{
          marginTop: "16px",
          padding: "12px",
          backgroundColor: "#f5f5f5",
          borderRadius: "4px",
          fontSize: "12px",
          color: "#666",
        }}
      >
        <strong>使用说明:</strong>
        <ol style={{ margin: "8px 0 0 0", paddingLeft: "20px" }}>
          <li>方式1: 在PPT中选中要删除的页面，点击&ldquo;删除当前页&rdquo;按钮</li>
          <li>
            方式2: 在输入框中输入页码（多个页码用逗号分隔），点击&ldquo;删除指定页码&rdquo;按钮
          </li>
          <li>
            方式3: 使用快速选择按钮选择页码（支持多选），然后点击&ldquo;删除指定页码&rdquo;按钮
          </li>
        </ol>
        <div style={{ marginTop: "8px", fontSize: "11px", color: "#999" }}>
          💡 提示: 如果页码不存在，不会抛出异常，只会在日志中记录。支持批量删除多个页面。
        </div>
      </div>
    </div>
  );
};
