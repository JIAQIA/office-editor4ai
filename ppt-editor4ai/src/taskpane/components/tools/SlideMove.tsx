/**
 * 文件名: SlideMove.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片移动调试组件，支持修改幻灯片页码/排序
 */

import React, { useState, useEffect } from "react";
import {
  moveSlide,
  moveCurrentSlide,
  swapSlides,
  getAllSlidesInfo,
  type SlideInfo,
} from "../../../ppt-tools";

export const SlideMove: React.FC = () => {
  const [fromIndex, setFromIndex] = useState<string>("");
  const [toIndex, setToIndex] = useState<string>("");
  const [swapIndex1, setSwapIndex1] = useState<string>("");
  const [swapIndex2, setSwapIndex2] = useState<string>("");
  const [moveCurrentTo, setMoveCurrentTo] = useState<string>("");

  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<string>("");
  const [messageType, setMessageType] = useState<"success" | "error" | "info">("info");

  const [slidesInfo, setSlidesInfo] = useState<SlideInfo[]>([]);
  const [totalSlides, setTotalSlides] = useState<number>(0);
  const [currentSlideIndex, setCurrentSlideIndex] = useState<number>(0);

  // 组件加载时获取幻灯片信息
  useEffect(() => {
    handleRefreshSlides();
  }, []);

  // 刷新幻灯片列表
  const handleRefreshSlides = async () => {
    setLoading(true);
    try {
      const info = await getAllSlidesInfo();
      setSlidesInfo(info);
      setTotalSlides(info.length);

      // 获取当前选中的幻灯片
      /* global PowerPoint */
      await PowerPoint.run(async (context) => {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          const selectedSlide = selectedSlides.items[0];
          selectedSlide.load("id");
          await context.sync();

          // 查找当前幻灯片的索引
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();

          for (let i = 0; i < slides.items.length; i++) {
            slides.items[i].load("id");
          }
          await context.sync();

          for (let i = 0; i < slides.items.length; i++) {
            if (slides.items[i].id === selectedSlide.id) {
              setCurrentSlideIndex(i + 1);
              break;
            }
          }
        }
      });

      setMessage(`已加载 ${info.length} 张幻灯片信息`);
      setMessageType("success");
    } catch (error) {
      setMessage(`加载幻灯片信息失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 移动幻灯片
  const handleMove = async () => {
    const from = parseInt(fromIndex);
    const to = parseInt(toIndex);

    if (isNaN(from) || from < 1) {
      setMessage("请输入有效的源位置（大于0的整数）");
      setMessageType("error");
      return;
    }

    if (isNaN(to) || to < 1) {
      setMessage("请输入有效的目标位置（大于0的整数）");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const result = await moveSlide({ fromIndex: from, toIndex: to });

      if (result.success) {
        setMessage(`移动成功: ${result.message}`);
        setMessageType("success");
        // 刷新幻灯片列表
        await handleRefreshSlides();
      } else {
        setMessage(`移动失败: ${result.message}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`移动失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 移动当前幻灯片
  const handleMoveCurrent = async () => {
    const to = parseInt(moveCurrentTo);

    if (isNaN(to) || to < 1) {
      setMessage("请输入有效的目标位置（大于0的整数）");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const result = await moveCurrentSlide(to);

      if (result.success) {
        setMessage(`移动成功: ${result.message}`);
        setMessageType("success");
        // 刷新幻灯片列表
        await handleRefreshSlides();
      } else {
        setMessage(`移动失败: ${result.message}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`移动失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 交换幻灯片
  const handleSwap = async () => {
    const index1 = parseInt(swapIndex1);
    const index2 = parseInt(swapIndex2);

    if (isNaN(index1) || index1 < 1) {
      setMessage("请输入有效的第一张幻灯片位置（大于0的整数）");
      setMessageType("error");
      return;
    }

    if (isNaN(index2) || index2 < 1) {
      setMessage("请输入有效的第二张幻灯片位置（大于0的整数）");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const result = await swapSlides(index1, index2);

      if (result.success) {
        setMessage(`交换成功: ${result.message}`);
        setMessageType("success");
        // 刷新幻灯片列表
        await handleRefreshSlides();
      } else {
        setMessage(`交换失败: ${result.message}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`交换失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 快速设置：将当前幻灯片移动到指定位置
  const handleQuickMove = (targetIndex: number) => {
    setMoveCurrentTo(targetIndex.toString());
  };

  return (
    <div style={{ padding: "16px" }}>
      <h3 style={{ marginTop: 0, marginBottom: "16px", fontSize: "16px", fontWeight: 600 }}>
        幻灯片移动工具
      </h3>

      {/* 幻灯片信息概览 */}
      <div
        style={{
          marginBottom: "16px",
          padding: "12px",
          backgroundColor: "#f5f5f5",
          borderRadius: "4px",
        }}
      >
        <div style={{ fontSize: "14px", marginBottom: "8px" }}>
          <strong>总幻灯片数:</strong> {totalSlides}
        </div>
        {currentSlideIndex > 0 && (
          <div style={{ fontSize: "14px", color: "#0078d4" }}>
            <strong>当前选中:</strong> 第 {currentSlideIndex} 张
          </div>
        )}
        <button
          onClick={handleRefreshSlides}
          disabled={loading}
          style={{
            marginTop: "8px",
            padding: "6px 12px",
            backgroundColor: "#0078d4",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading ? "not-allowed" : "pointer",
            fontSize: "12px",
          }}
        >
          {loading ? "刷新中..." : "刷新列表"}
        </button>
      </div>

      {/* 幻灯片列表 */}
      {slidesInfo.length > 0 && (
        <div
          style={{
            marginBottom: "16px",
            maxHeight: "200px",
            overflowY: "auto",
            border: "1px solid #ccc",
            borderRadius: "4px",
            padding: "8px",
            backgroundColor: "#fafafa",
          }}
        >
          <div style={{ fontSize: "12px", fontWeight: 600, marginBottom: "8px" }}>
            幻灯片列表:
          </div>
          {slidesInfo.map((slide) => (
            <div
              key={slide.id}
              style={{
                padding: "6px 8px",
                marginBottom: "4px",
                backgroundColor: slide.index === currentSlideIndex ? "#e1f5fe" : "white",
                borderRadius: "4px",
                fontSize: "12px",
                border:
                  slide.index === currentSlideIndex ? "1px solid #0078d4" : "1px solid #e0e0e0",
              }}
            >
              <strong>#{slide.index}</strong>
              {slide.title && <span style={{ marginLeft: "8px" }}>{slide.title}</span>}
            </div>
          ))}
        </div>
      )}

      {/* 方法1: 移动指定幻灯片 */}
      <div style={{ marginBottom: "16px", padding: "12px", border: "1px solid #ccc", borderRadius: "4px" }}>
        <h4 style={{ marginTop: 0, marginBottom: "12px", fontSize: "14px", fontWeight: 600 }}>
          方法1: 移动指定幻灯片
        </h4>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px", marginBottom: "8px" }}>
          <div>
            <label
              htmlFor="fromIndex"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              源位置:
            </label>
            <input
              id="fromIndex"
              type="number"
              min="1"
              value={fromIndex}
              onChange={(e) => setFromIndex(e.target.value)}
              placeholder="如: 1"
              style={{
                width: "100%",
                padding: "6px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                fontSize: "14px",
                boxSizing: "border-box",
              }}
            />
          </div>
          <div>
            <label
              htmlFor="toIndex"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              目标位置:
            </label>
            <input
              id="toIndex"
              type="number"
              min="1"
              value={toIndex}
              onChange={(e) => setToIndex(e.target.value)}
              placeholder="如: 3"
              style={{
                width: "100%",
                padding: "6px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                fontSize: "14px",
                boxSizing: "border-box",
              }}
            />
          </div>
        </div>
        <button
          onClick={handleMove}
          disabled={loading || !fromIndex || !toIndex}
          style={{
            width: "100%",
            padding: "8px 16px",
            backgroundColor: loading || !fromIndex || !toIndex ? "#ccc" : "#107c10",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading || !fromIndex || !toIndex ? "not-allowed" : "pointer",
            fontSize: "14px",
            fontWeight: 600,
          }}
        >
          {loading ? "移动中..." : "移动幻灯片"}
        </button>
      </div>

      {/* 方法2: 移动当前幻灯片 */}
      <div style={{ marginBottom: "16px", padding: "12px", border: "1px solid #ccc", borderRadius: "4px" }}>
        <h4 style={{ marginTop: 0, marginBottom: "12px", fontSize: "14px", fontWeight: 600 }}>
          方法2: 移动当前选中的幻灯片
        </h4>
        <div style={{ marginBottom: "8px" }}>
          <label
            htmlFor="moveCurrentTo"
            style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
          >
            移动到位置:
          </label>
          <input
            id="moveCurrentTo"
            type="number"
            min="1"
            value={moveCurrentTo}
            onChange={(e) => setMoveCurrentTo(e.target.value)}
            placeholder="如: 5"
            style={{
              width: "100%",
              padding: "6px",
              border: "1px solid #ccc",
              borderRadius: "4px",
              fontSize: "14px",
              boxSizing: "border-box",
            }}
          />
        </div>
        {currentSlideIndex > 0 && (
          <div style={{ marginBottom: "8px", fontSize: "12px", color: "#666" }}>
            快速移动到:
            <div style={{ display: "flex", gap: "4px", marginTop: "4px", flexWrap: "wrap" }}>
              <button
                onClick={() => handleQuickMove(1)}
                disabled={loading || currentSlideIndex === 1}
                style={{
                  padding: "4px 8px",
                  fontSize: "11px",
                  backgroundColor: "#f0f0f0",
                  border: "1px solid #ccc",
                  borderRadius: "4px",
                  cursor: loading || currentSlideIndex === 1 ? "not-allowed" : "pointer",
                }}
              >
                开头
              </button>
              <button
                onClick={() => handleQuickMove(totalSlides)}
                disabled={loading || currentSlideIndex === totalSlides}
                style={{
                  padding: "4px 8px",
                  fontSize: "11px",
                  backgroundColor: "#f0f0f0",
                  border: "1px solid #ccc",
                  borderRadius: "4px",
                  cursor: loading || currentSlideIndex === totalSlides ? "not-allowed" : "pointer",
                }}
              >
                末尾
              </button>
              {currentSlideIndex > 1 && (
                <button
                  onClick={() => handleQuickMove(currentSlideIndex - 1)}
                  disabled={loading}
                  style={{
                    padding: "4px 8px",
                    fontSize: "11px",
                    backgroundColor: "#f0f0f0",
                    border: "1px solid #ccc",
                    borderRadius: "4px",
                    cursor: loading ? "not-allowed" : "pointer",
                  }}
                >
                  前移一位
                </button>
              )}
              {currentSlideIndex < totalSlides && (
                <button
                  onClick={() => handleQuickMove(currentSlideIndex + 1)}
                  disabled={loading}
                  style={{
                    padding: "4px 8px",
                    fontSize: "11px",
                    backgroundColor: "#f0f0f0",
                    border: "1px solid #ccc",
                    borderRadius: "4px",
                    cursor: loading ? "not-allowed" : "pointer",
                  }}
                >
                  后移一位
                </button>
              )}
            </div>
          </div>
        )}
        <button
          onClick={handleMoveCurrent}
          disabled={loading || !moveCurrentTo || currentSlideIndex === 0}
          style={{
            width: "100%",
            padding: "8px 16px",
            backgroundColor: loading || !moveCurrentTo || currentSlideIndex === 0 ? "#ccc" : "#0078d4",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading || !moveCurrentTo || currentSlideIndex === 0 ? "not-allowed" : "pointer",
            fontSize: "14px",
            fontWeight: 600,
          }}
        >
          {loading ? "移动中..." : "移动当前幻灯片"}
        </button>
      </div>

      {/* 方法3: 交换两张幻灯片 */}
      <div style={{ marginBottom: "16px", padding: "12px", border: "1px solid #ccc", borderRadius: "4px" }}>
        <h4 style={{ marginTop: 0, marginBottom: "12px", fontSize: "14px", fontWeight: 600 }}>
          方法3: 交换两张幻灯片位置
        </h4>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px", marginBottom: "8px" }}>
          <div>
            <label
              htmlFor="swapIndex1"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              第一张位置:
            </label>
            <input
              id="swapIndex1"
              type="number"
              min="1"
              value={swapIndex1}
              onChange={(e) => setSwapIndex1(e.target.value)}
              placeholder="如: 2"
              style={{
                width: "100%",
                padding: "6px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                fontSize: "14px",
                boxSizing: "border-box",
              }}
            />
          </div>
          <div>
            <label
              htmlFor="swapIndex2"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              第二张位置:
            </label>
            <input
              id="swapIndex2"
              type="number"
              min="1"
              value={swapIndex2}
              onChange={(e) => setSwapIndex2(e.target.value)}
              placeholder="如: 4"
              style={{
                width: "100%",
                padding: "6px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                fontSize: "14px",
                boxSizing: "border-box",
              }}
            />
          </div>
        </div>
        <button
          onClick={handleSwap}
          disabled={loading || !swapIndex1 || !swapIndex2}
          style={{
            width: "100%",
            padding: "8px 16px",
            backgroundColor: loading || !swapIndex1 || !swapIndex2 ? "#ccc" : "#d83b01",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading || !swapIndex1 || !swapIndex2 ? "not-allowed" : "pointer",
            fontSize: "14px",
            fontWeight: 600,
          }}
        >
          {loading ? "交换中..." : "交换幻灯片"}
        </button>
      </div>

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
          <li>
            <strong>方法1</strong>: 输入源位置和目标位置，移动指定幻灯片
          </li>
          <li>
            <strong>方法2</strong>: 在PPT中选中一张幻灯片，输入目标位置，移动当前幻灯片
          </li>
          <li>
            <strong>方法3</strong>: 输入两张幻灯片的位置，交换它们的顺序
          </li>
        </ol>
        <div style={{ marginTop: "8px", fontSize: "11px", color: "#999" }}>
          💡 提示: 位置索引从1开始，操作后会自动刷新幻灯片列表
        </div>
      </div>
    </div>
  );
};

export default SlideMove;
