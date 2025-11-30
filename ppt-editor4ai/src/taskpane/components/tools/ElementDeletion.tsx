/**
 * 文件名: ElementDeletion.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 元素删除调试组件
 */

import React, { useState } from "react";
import { getCurrentSlideElements, type SlideElement } from "../../../ppt-tools";
import { deleteElementById } from "../../../ppt-tools";

export const ElementDeletion: React.FC = () => {
  const [elements, setElements] = useState<SlideElement[]>([]);
  const [selectedElementId, setSelectedElementId] = useState<string>("");
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<string>("");
  const [messageType, setMessageType] = useState<"success" | "error" | "info">("info");

  // 获取当前页面元素列表
  const handleGetElements = async () => {
    setLoading(true);
    setMessage("");
    try {
      const elementsList = await getCurrentSlideElements();
      setElements(elementsList);
      setMessage(`成功获取 ${elementsList.length} 个元素`);
      setMessageType("success");
    } catch (error) {
      setMessage(`获取元素列表失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
      console.error("获取元素列表失败:", error);
    } finally {
      setLoading(false);
    }
  };

  // 删除指定ID的元素
  const handleDeleteElement = async () => {
    if (!selectedElementId.trim()) {
      setMessage("请先输入或选择元素ID");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const result = await deleteElementById(selectedElementId);
      if (result.success) {
        setMessage(`删除成功: ${result.message}`);
        setMessageType("success");
        // 刷新元素列表
        await handleGetElements();
        setSelectedElementId("");
      } else {
        setMessage(`删除失败: ${result.message}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`删除元素失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
      console.error("删除元素失败:", error);
    } finally {
      setLoading(false);
    }
  };

  // 选中元素
  const handleSelectElement = (elementId: string) => {
    setSelectedElementId(elementId);
    setMessage(`已选中元素: ${elementId}`);
    setMessageType("info");
  };

  return (
    <div style={{ padding: "16px" }}>
      <h3 style={{ marginTop: 0, marginBottom: "16px", fontSize: "16px", fontWeight: 600 }}>
        元素删除调试工具
      </h3>

      {/* 操作按钮区域 */}
      <div style={{ marginBottom: "16px" }}>
        <button
          onClick={handleGetElements}
          disabled={loading}
          style={{
            width: "100%",
            padding: "8px 16px",
            backgroundColor: "#0078d4",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading ? "not-allowed" : "pointer",
            fontSize: "14px",
            marginBottom: "8px",
          }}
        >
          {loading ? "加载中..." : "获取当前页面元素"}
        </button>
      </div>

      {/* 元素ID输入区域 */}
      <div style={{ marginBottom: "16px" }}>
        <label
          style={{
            display: "block",
            marginBottom: "8px",
            fontSize: "14px",
            fontWeight: 500,
          }}
        >
          元素ID:
        </label>
        <input
          type="text"
          value={selectedElementId}
          onChange={(e) => setSelectedElementId(e.target.value)}
          placeholder="输入或从列表中选择元素ID"
          style={{
            width: "100%",
            padding: "8px",
            border: "1px solid #ccc",
            borderRadius: "4px",
            fontSize: "14px",
            boxSizing: "border-box",
          }}
        />
        <button
          onClick={handleDeleteElement}
          disabled={loading || !selectedElementId.trim()}
          style={{
            width: "100%",
            padding: "8px 16px",
            backgroundColor: "#d13438",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading || !selectedElementId.trim() ? "not-allowed" : "pointer",
            fontSize: "14px",
            marginTop: "8px",
          }}
        >
          删除选中元素
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

      {/* 元素列表 */}
      {elements.length > 0 && (
        <div>
          <h4 style={{ marginTop: 0, marginBottom: "12px", fontSize: "14px", fontWeight: 600 }}>
            元素列表 ({elements.length}):
          </h4>
          <div
            style={{
              maxHeight: "400px",
              overflowY: "auto",
              border: "1px solid #ccc",
              borderRadius: "4px",
            }}
          >
            {elements.map((element, index) => (
              <div
                key={element.id}
                onClick={() => handleSelectElement(element.id)}
                style={{
                  padding: "12px",
                  borderBottom: index < elements.length - 1 ? "1px solid #eee" : "none",
                  cursor: "pointer",
                  backgroundColor: selectedElementId === element.id ? "#e1f5fe" : "white",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) => {
                  if (selectedElementId !== element.id) {
                    e.currentTarget.style.backgroundColor = "#f5f5f5";
                  }
                }}
                onMouseLeave={(e) => {
                  if (selectedElementId !== element.id) {
                    e.currentTarget.style.backgroundColor = "white";
                  }
                }}
              >
                <div
                  style={{
                    fontSize: "12px",
                    fontWeight: 600,
                    marginBottom: "4px",
                    color: "#333",
                  }}
                >
                  {element.type}
                  {element.name && ` - ${element.name}`}
                </div>
                <div
                  style={{
                    fontSize: "11px",
                    color: "#666",
                    fontFamily: "monospace",
                    marginBottom: "4px",
                  }}
                >
                  ID: {element.id}
                </div>
                <div style={{ fontSize: "11px", color: "#999" }}>
                  位置: ({Math.round(element.left)}, {Math.round(element.top)}) | 尺寸:{" "}
                  {Math.round(element.width)} × {Math.round(element.height)}
                </div>
                {element.text && (
                  <div
                    style={{
                      fontSize: "11px",
                      color: "#666",
                      marginTop: "4px",
                      fontStyle: "italic",
                      maxWidth: "100%",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    文本: {element.text}
                  </div>
                )}
                {element.placeholderType && (
                  <div style={{ fontSize: "11px", color: "#0078d4", marginTop: "4px" }}>
                    占位符类型: {element.placeholderType}
                  </div>
                )}
              </div>
            ))}
          </div>
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
          <li>点击"获取当前页面元素"按钮获取元素列表</li>
          <li>在列表中点击选择要删除的元素</li>
          <li>或者手动输入元素ID</li>
          <li>点击"删除选中元素"按钮执行删除</li>
        </ol>
      </div>
    </div>
  );
};
