/**
 * 文件名: TextUpdate.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 文本框更新调试组件
 */

import React, { useState } from "react";
import { updateTextBox, getTextBoxStyle, type TextUpdateOptions } from "../../../ppt-tools";

export const TextUpdate: React.FC = () => {
  const [elementId, setElementId] = useState<string>("");
  const [text, setText] = useState<string>("");
  const [fontSize, setFontSize] = useState<string>("");
  const [fontName, setFontName] = useState<string>("");
  const [fontColor, setFontColor] = useState<string>("#000000");
  const [bold, setBold] = useState<boolean>(false);
  const [italic, setItalic] = useState<boolean>(false);
  const [underline, setUnderline] = useState<boolean>(false);
  const [horizontalAlignment, setHorizontalAlignment] = useState<string>("Left");
  const [verticalAlignment, setVerticalAlignment] = useState<string>("Top");
  const [backgroundColor, setBackgroundColor] = useState<string>("#FFFFFF");
  const [left, setLeft] = useState<string>("");
  const [top, setTop] = useState<string>("");
  const [width, setWidth] = useState<string>("");
  const [height, setHeight] = useState<string>("");

  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<string>("");
  const [messageType, setMessageType] = useState<"success" | "error" | "info">("info");
  const [selectedShapeType, setSelectedShapeType] = useState<string>("");

  // 获取用户在PPT中选中的元素
  const handleGetSelectedShape = async () => {
    setLoading(true);
    setMessage("");
    try {
      /* global PowerPoint */
      await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        const shapeCount = shapes.getCount();
        await context.sync();

        if (shapeCount.value === 0) {
          setMessage("请先在幻灯片中选中一个文本框元素");
          setMessageType("error");
          setSelectedShapeType("");
          return;
        }

        if (shapeCount.value > 1) {
          setMessage("请只选中一个元素");
          setMessageType("error");
          setSelectedShapeType("");
          return;
        }

        // 获取选中的形状
        shapes.load("items");
        await context.sync();

        const shape = shapes.items[0];
        shape.load("id,type,name");
        await context.sync();

        setElementId(shape.id);
        setSelectedShapeType(shape.type);

        // 验证元素类型
        const supportedTypes = ["TextBox", "Placeholder", "GeometricShape"];
        if (!supportedTypes.includes(shape.type)) {
          setMessage(`警告: 选中的元素类型 "${shape.type}" 可能不支持文本编辑`);
          setMessageType("error");
          return;
        }

        setMessage(`已获取选中元素: ${shape.type}${shape.name ? ` (${shape.name})` : ""}`);
        setMessageType("success");

        // 自动加载当前样式
        await handleLoadStyle(shape.id);
      });
    } catch (error) {
      setMessage(`获取选中元素失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
      setSelectedShapeType("");
    } finally {
      setLoading(false);
    }
  };

  // 加载元素当前样式
  const handleLoadStyle = async (targetId?: string) => {
    const idToLoad = targetId || elementId;
    if (!idToLoad.trim()) {
      setMessage("请先输入或选择元素ID");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const style = await getTextBoxStyle(idToLoad);
      if (style) {
        setText(style.text || "");
        setFontSize(style.fontSize?.toString() || "");
        setFontName(style.fontName || "");
        setFontColor(style.fontColor || "#000000");
        setBold(style.bold || false);
        setItalic(style.italic || false);
        setUnderline(style.underline || false);
        setHorizontalAlignment(style.horizontalAlignment || "Left");
        setVerticalAlignment(style.verticalAlignment || "Top");
        setBackgroundColor(style.backgroundColor || "#FFFFFF");
        setLeft(style.left?.toString() || "");
        setTop(style.top?.toString() || "");
        setWidth(style.width?.toString() || "");
        setHeight(style.height?.toString() || "");

        if (!targetId) {
          setMessage("成功加载当前样式");
          setMessageType("success");
        }
      } else {
        setMessage("加载样式失败");
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`加载样式失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 更新文本框
  const handleUpdate = async () => {
    if (!elementId.trim()) {
      setMessage("请先输入或选择元素ID");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const options: TextUpdateOptions = {
        elementId: elementId.trim(),
      };

      // 只添加用户修改过的属性
      // 注意：text 可以是空字符串（用于清空文本框）
      options.text = text;
      if (fontSize !== "") options.fontSize = parseFloat(fontSize);
      if (fontName !== "") options.fontName = fontName;
      if (fontColor !== "") options.fontColor = fontColor;
      options.bold = bold;
      options.italic = italic;
      options.underline = underline;
      options.horizontalAlignment = horizontalAlignment as any;
      options.verticalAlignment = verticalAlignment as any;
      if (backgroundColor !== "") options.backgroundColor = backgroundColor;
      if (left !== "") options.left = parseFloat(left);
      if (top !== "") options.top = parseFloat(top);
      if (width !== "") options.width = parseFloat(width);
      if (height !== "") options.height = parseFloat(height);

      const result = await updateTextBox(options);

      if (result.success) {
        setMessage(`更新成功: ${result.message}`);
        setMessageType("success");
      } else {
        setMessage(`更新失败: ${result.message}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`更新失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 重置表单
  const handleReset = () => {
    setText("");
    setFontSize("");
    setFontName("");
    setFontColor("#000000");
    setBold(false);
    setItalic(false);
    setUnderline(false);
    setHorizontalAlignment("Left");
    setVerticalAlignment("Top");
    setBackgroundColor("#FFFFFF");
    setLeft("");
    setTop("");
    setWidth("");
    setHeight("");
    setMessage("已重置所有字段");
    setMessageType("info");
  };

  // 检查是否可以更新
  const canUpdate = elementId.trim() !== "" && selectedShapeType !== "";
  const isUnsupportedType =
    selectedShapeType !== "" &&
    !["TextBox", "Placeholder", "GeometricShape"].includes(selectedShapeType);

  return (
    <div style={{ padding: "16px" }}>
      <h3 style={{ marginTop: 0, marginBottom: "16px", fontSize: "16px", fontWeight: 600 }}>
        文本框更新工具
      </h3>

      {/* 元素选择区域 */}
      <div style={{ marginBottom: "16px" }}>
        <button
          onClick={handleGetSelectedShape}
          disabled={loading}
          style={{
            width: "100%",
            padding: "8px 16px",
            backgroundColor: "#106ebe",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading ? "not-allowed" : "pointer",
            fontSize: "14px",
            marginBottom: "8px",
          }}
        >
          {loading ? "加载中..." : "获取PPT中选中的元素"}
        </button>

        <label
          htmlFor="elementId"
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
          id="elementId"
          type="text"
          value={elementId}
          onChange={(e) => setElementId(e.target.value)}
          placeholder="输入元素ID或从PPT中选择"
          style={{
            width: "100%",
            padding: "8px",
            border: "1px solid #ccc",
            borderRadius: "4px",
            fontSize: "14px",
            boxSizing: "border-box",
            fontFamily: "monospace",
          }}
        />

        <button
          onClick={() => handleLoadStyle()}
          disabled={loading || !elementId.trim()}
          style={{
            width: "100%",
            padding: "8px 16px",
            backgroundColor: "#0078d4",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading || !elementId.trim() ? "not-allowed" : "pointer",
            fontSize: "14px",
            marginTop: "8px",
          }}
        >
          加载当前样式
        </button>
      </div>

      {/* 警告信息 */}
      {isUnsupportedType && (
        <div
          style={{
            padding: "12px",
            marginBottom: "16px",
            borderRadius: "4px",
            fontSize: "14px",
            backgroundColor: "#fde7e9",
            color: "#a80000",
            border: "1px solid #a80000",
          }}
        >
          ⚠️ 选中的元素类型 "{selectedShapeType}" 不支持文本编辑，请选择文本框、占位符或几何形状
        </div>
      )}

      {/* 文本内容 */}
      <div style={{ marginBottom: "16px" }}>
        <label
          htmlFor="text"
          style={{
            display: "block",
            marginBottom: "8px",
            fontSize: "14px",
            fontWeight: 500,
          }}
        >
          文本内容:
        </label>
        <textarea
          id="text"
          value={text}
          onChange={(e) => setText(e.target.value)}
          placeholder="输入新的文本内容（留空则不修改）"
          rows={3}
          style={{
            width: "100%",
            padding: "8px",
            border: "1px solid #ccc",
            borderRadius: "4px",
            fontSize: "14px",
            boxSizing: "border-box",
            resize: "vertical",
          }}
        />
      </div>

      {/* 字体设置 */}
      <div style={{ marginBottom: "16px" }}>
        <h4 style={{ marginTop: 0, marginBottom: "12px", fontSize: "14px", fontWeight: 600 }}>
          字体设置
        </h4>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px" }}>
          <div>
            <label
              htmlFor="fontSize"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              字号:
            </label>
            <input
              id="fontSize"
              type="number"
              value={fontSize}
              onChange={(e) => setFontSize(e.target.value)}
              placeholder="如: 18"
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
              htmlFor="fontName"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              字体:
            </label>
            <input
              id="fontName"
              type="text"
              value={fontName}
              onChange={(e) => setFontName(e.target.value)}
              placeholder="如: Arial"
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
              htmlFor="fontColor"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              字体颜色:
            </label>
            <input
              id="fontColor"
              type="color"
              value={fontColor}
              onChange={(e) => setFontColor(e.target.value)}
              style={{
                width: "100%",
                height: "32px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                cursor: "pointer",
              }}
            />
          </div>
          <div>
            <label
              htmlFor="backgroundColor"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              背景颜色:
            </label>
            <input
              id="backgroundColor"
              type="color"
              value={backgroundColor}
              onChange={(e) => setBackgroundColor(e.target.value)}
              style={{
                width: "100%",
                height: "32px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                cursor: "pointer",
              }}
            />
          </div>
        </div>

        {/* 字体样式 */}
        <div style={{ marginTop: "12px", display: "flex", gap: "12px", flexWrap: "wrap" }}>
          <label
            style={{ display: "flex", alignItems: "center", fontSize: "14px", cursor: "pointer" }}
          >
            <input
              type="checkbox"
              checked={bold}
              onChange={(e) => setBold(e.target.checked)}
              style={{ marginRight: "6px" }}
            />
            <strong>加粗</strong>
          </label>
          <label
            style={{ display: "flex", alignItems: "center", fontSize: "14px", cursor: "pointer" }}
          >
            <input
              type="checkbox"
              checked={italic}
              onChange={(e) => setItalic(e.target.checked)}
              style={{ marginRight: "6px" }}
            />
            <em>斜体</em>
          </label>
          <label
            style={{ display: "flex", alignItems: "center", fontSize: "14px", cursor: "pointer" }}
          >
            <input
              type="checkbox"
              checked={underline}
              onChange={(e) => setUnderline(e.target.checked)}
              style={{ marginRight: "6px" }}
            />
            <u>下划线</u>
          </label>
        </div>
      </div>

      {/* 对齐方式 */}
      <div style={{ marginBottom: "16px" }}>
        <h4 style={{ marginTop: 0, marginBottom: "12px", fontSize: "14px", fontWeight: 600 }}>
          对齐方式
        </h4>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px" }}>
          <div>
            <label
              htmlFor="horizontalAlignment"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              水平对齐:
            </label>
            <select
              id="horizontalAlignment"
              value={horizontalAlignment}
              onChange={(e) => setHorizontalAlignment(e.target.value)}
              style={{
                width: "100%",
                padding: "6px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                fontSize: "14px",
                boxSizing: "border-box",
              }}
            >
              <option value="Left">左对齐</option>
              <option value="Center">居中</option>
              <option value="Right">右对齐</option>
              <option value="Justify">两端对齐</option>
              <option value="Distributed">分散对齐</option>
            </select>
          </div>
          <div>
            <label
              htmlFor="verticalAlignment"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              垂直对齐:
            </label>
            <select
              id="verticalAlignment"
              value={verticalAlignment}
              onChange={(e) => setVerticalAlignment(e.target.value)}
              style={{
                width: "100%",
                padding: "6px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                fontSize: "14px",
                boxSizing: "border-box",
              }}
            >
              <option value="Top">顶部</option>
              <option value="Middle">居中</option>
              <option value="Bottom">底部</option>
            </select>
          </div>
        </div>
      </div>

      {/* 位置和尺寸 */}
      <div style={{ marginBottom: "16px" }}>
        <h4 style={{ marginTop: 0, marginBottom: "12px", fontSize: "14px", fontWeight: 600 }}>
          位置和尺寸
        </h4>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px" }}>
          <div>
            <label
              htmlFor="left"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              X坐标:
            </label>
            <input
              id="left"
              type="number"
              value={left}
              onChange={(e) => setLeft(e.target.value)}
              placeholder="如: 100"
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
              htmlFor="top"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              Y坐标:
            </label>
            <input
              id="top"
              type="number"
              value={top}
              onChange={(e) => setTop(e.target.value)}
              placeholder="如: 100"
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
              htmlFor="width"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              宽度:
            </label>
            <input
              id="width"
              type="number"
              value={width}
              onChange={(e) => setWidth(e.target.value)}
              placeholder="如: 300"
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
              htmlFor="height"
              style={{ display: "block", marginBottom: "4px", fontSize: "12px" }}
            >
              高度:
            </label>
            <input
              id="height"
              type="number"
              value={height}
              onChange={(e) => setHeight(e.target.value)}
              placeholder="如: 100"
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
      </div>

      {/* 操作按钮 */}
      <div style={{ display: "flex", gap: "8px", marginBottom: "16px" }}>
        <button
          onClick={handleUpdate}
          disabled={loading || !canUpdate || isUnsupportedType}
          style={{
            flex: 1,
            padding: "10px 16px",
            backgroundColor: loading || !canUpdate || isUnsupportedType ? "#ccc" : "#107c10",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading || !canUpdate || isUnsupportedType ? "not-allowed" : "pointer",
            fontSize: "14px",
            fontWeight: 600,
          }}
        >
          {loading ? "更新中..." : "更新文本框"}
        </button>
        <button
          onClick={handleReset}
          disabled={loading}
          style={{
            padding: "10px 16px",
            backgroundColor: "#605e5c",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading ? "not-allowed" : "pointer",
            fontSize: "14px",
          }}
        >
          重置
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
          <li>在PPT中选中一个文本框元素，点击"获取PPT中选中的元素"</li>
          <li>点击"加载当前样式"查看元素的当前属性</li>
          <li>修改需要更新的属性（留空的字段不会被修改）</li>
          <li>点击"更新文本框"应用更改</li>
        </ol>
        <div style={{ marginTop: "8px", fontSize: "11px", color: "#999" }}>
          💡 提示: 支持的元素类型包括文本框、占位符和几何形状
        </div>
      </div>
    </div>
  );
};

export default TextUpdate;
