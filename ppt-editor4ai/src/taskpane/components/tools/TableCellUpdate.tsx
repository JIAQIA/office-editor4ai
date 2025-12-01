/**
 * 文件名: TableCellUpdate.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 最后修改日期: 2025/12/1
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 表格单元格更新调试组件 - 通过坐标修改单元格内容
 * Description: Table cell update debug component - modify cell content by coordinates
 */

import React, { useState } from "react";
import {
  Button,
  Field,
  Textarea,
  Input,
  tokens,
  makeStyles,
} from "@fluentui/react-components";
import {
  updateTableCell,
  updateTableCellsBatch,
  getTableCellContent,
  type CellUpdateOptions,
} from "../../../ppt-tools";

/* global PowerPoint */

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    marginBottom: "8px",
    textAlign: "center",
  },
  description: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginBottom: "16px",
    textAlign: "center",
    lineHeight: "1.4",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    marginLeft: "8px",
    marginRight: "8px",
    marginBottom: "16px",
    width: "calc(100% - 16px)",
    maxWidth: "100%",
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
    marginBottom: "8px",
  },
  field: {
    width: "100%",
  },
  row: {
    display: "flex",
    gap: "12px",
    width: "100%",
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    width: "100%",
    justifyContent: "center",
  },
  message: {
    padding: "12px",
    borderRadius: "4px",
    marginTop: "12px",
    marginLeft: "8px",
    marginRight: "8px",
    width: "calc(100% - 16px)",
    fontSize: tokens.fontSizeBase200,
  },
  messageSuccess: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground1,
  },
  messageError: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground1,
  },
  messageInfo: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
  },
  usageTips: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginLeft: "8px",
    marginRight: "8px",
    marginTop: "16px",
    width: "calc(100% - 16px)",
    lineHeight: "1.6",
  },
  tipsList: {
    marginTop: "8px",
    paddingLeft: "20px",
  },
});

export const TableCellUpdate: React.FC = () => {
  const styles = useStyles();
  // 表格定位
  const [shapeId, setShapeId] = useState<string>("");
  const [tableIndex, setTableIndex] = useState<string>("0");

  // 单个单元格更新
  const [rowIndex, setRowIndex] = useState<string>("1");
  const [columnIndex, setColumnIndex] = useState<string>("1");
  const [cellText, setCellText] = useState<string>("");

  // 批量更新
  const [batchData, setBatchData] = useState<string>("");

  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<string>("");
  const [messageType, setMessageType] = useState<"success" | "error" | "info">("info");
  const [selectedShapeType, setSelectedShapeType] = useState<string>("");

  // 获取用户在PPT中选中的表格
  const handleGetSelectedTable = async () => {
    setLoading(true);
    setMessage("");
    try {
      await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        const shapeCount = shapes.getCount();
        await context.sync();

        if (shapeCount.value === 0) {
          setMessage("请先在幻灯片中选中一个表格");
          setMessageType("error");
          setSelectedShapeType("");
          return;
        }

        if (shapeCount.value > 1) {
          setMessage("请只选中一个表格");
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

        setShapeId(shape.id);
        setSelectedShapeType(shape.type);

        // 验证元素类型
        if (shape.type !== "Table") {
          setMessage(`警告: 选中的元素类型是 "${shape.type}"，不是表格。请选择表格元素。`);
          setMessageType("error");
          return;
        }

        // 获取表格信息
        const table = shape.getTable();
        table.load("rowCount, columnCount");
        await context.sync();

        setMessage(
          `已获取选中表格: ${table.rowCount} 行 × ${table.columnCount} 列${shape.name ? ` (${shape.name})` : ""}`
        );
        setMessageType("success");
      });
    } catch (error) {
      setMessage(`获取选中表格失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
      setSelectedShapeType("");
    } finally {
      setLoading(false);
    }
  };

  // 获取单元格内容
  const handleGetCellContent = async () => {
    const row = parseInt(rowIndex);
    const col = parseInt(columnIndex);

    if (isNaN(row) || isNaN(col)) {
      setMessage("请输入有效的行列索引");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const tableLocation = shapeId.trim() ? { shapeId } : { tableIndex: parseInt(tableIndex) || 0 };
      // 转换为0开始的索引
      const content = await getTableCellContent(row - 1, col - 1, tableLocation);
      setCellText(content);
      setMessage(`已获取单元格 (${row}, ${col}) 的内容`);
      setMessageType("success");
    } catch (error) {
      setMessage(`获取单元格内容失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 更新单个单元格
  const handleUpdateCell = async () => {
    const row = parseInt(rowIndex);
    const col = parseInt(columnIndex);

    if (isNaN(row) || isNaN(col)) {
      setMessage("请输入有效的行列索引");
      setMessageType("error");
      return;
    }

    if (!cellText.trim()) {
      setMessage("请输入单元格内容");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const tableLocation = shapeId.trim() ? { shapeId } : { tableIndex: parseInt(tableIndex) || 0 };
      // 转换为0开始的索引
      const result = await updateTableCell(
        { rowIndex: row - 1, columnIndex: col - 1, text: cellText },
        tableLocation
      );

      if (result.success) {
        setMessage(
          `成功更新单元格 (${row}, ${col})。表格大小: ${result.rowCount} 行 × ${result.columnCount} 列`
        );
        setMessageType("success");
      } else {
        setMessage(`更新失败: ${result.error || "未知错误"}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`更新单元格失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 批量更新单元格
  const handleBatchUpdate = async () => {
    if (!batchData.trim()) {
      setMessage("请输入批量更新数据");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      // 解析批量数据 (格式: row,col,text 每行一个)
      const lines = batchData.trim().split("\n");
      const cells: CellUpdateOptions[] = [];

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;

        const parts = line.split(",");
        if (parts.length < 3) {
          setMessage(`第 ${i + 1} 行格式错误，应为: 行索引,列索引,文本内容`);
          setMessageType("error");
          return;
        }

        const row = parseInt(parts[0].trim());
        const col = parseInt(parts[1].trim());
        const text = parts.slice(2).join(",").trim();

        if (isNaN(row) || isNaN(col)) {
          setMessage(`第 ${i + 1} 行的行列索引无效`);
          setMessageType("error");
          return;
        }

        // 转换为0开始的索引
        cells.push({ rowIndex: row - 1, columnIndex: col - 1, text });
      }

      if (cells.length === 0) {
        setMessage("没有有效的更新数据");
        setMessageType("error");
        return;
      }

      const tableLocation = shapeId.trim() ? { shapeId } : { tableIndex: parseInt(tableIndex) || 0 };
      const result = await updateTableCellsBatch({ cells }, tableLocation);

      if (result.success) {
        setMessage(
          `成功批量更新 ${result.cellsUpdated} 个单元格。表格大小: ${result.rowCount} 行 × ${result.columnCount} 列`
        );
        setMessageType("success");
      } else {
        setMessage(`批量更新失败: ${result.error || "未知错误"}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`批量更新失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 判断按钮是否可用
  const isUpdateDisabled = selectedShapeType !== "" && selectedShapeType !== "Table";

  return (
    <div className={styles.container}>
      <div className={styles.title}>表格单元格更新</div>
      <div className={styles.description}>通过行列坐标修改表格单元格内容（行列编号从 1 开始）</div>

      {/* 表格定位 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>1. 表格定位</div>
        <Button
          appearance="primary"
          onClick={handleGetSelectedTable}
          disabled={loading}
          style={{ width: "100%" }}
        >
          {loading ? "获取中..." : "获取选中的表格"}
        </Button>

        <Field className={styles.field} label="表格形状 ID（可选）">
          <Input
            type="text"
            value={shapeId}
            onChange={(e) => setShapeId(e.target.value)}
            placeholder="留空则使用表格索引"
            disabled={loading}
          />
        </Field>

        <Field className={styles.field} label="表格索引（默认 0）">
          <Input
            type="number"
            value={tableIndex}
            onChange={(e) => setTableIndex(e.target.value)}
            disabled={loading || shapeId.trim() !== ""}
          />
        </Field>
      </div>

      {/* 单个单元格更新 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>2. 单个单元格更新</div>
        <div className={styles.row}>
          <Field className={styles.field} label="行索引">
            <Input
              type="number"
              value={rowIndex}
              onChange={(e) => setRowIndex(e.target.value)}
              disabled={loading}
            />
          </Field>
          <Field className={styles.field} label="列索引">
            <Input
              type="number"
              value={columnIndex}
              onChange={(e) => setColumnIndex(e.target.value)}
              disabled={loading}
            />
          </Field>
        </div>

        <Field className={styles.field} label="单元格内容">
          <Textarea
            value={cellText}
            onChange={(e) => setCellText(e.target.value)}
            placeholder="输入单元格文本内容"
            rows={3}
            disabled={loading}
          />
        </Field>

        <div className={styles.buttonGroup}>
          <Button onClick={handleGetCellContent} disabled={loading || isUpdateDisabled}>
            获取单元格内容
          </Button>
          <Button appearance="primary" onClick={handleUpdateCell} disabled={loading || isUpdateDisabled}>
            {loading ? "更新中..." : "更新单元格"}
          </Button>
        </div>
      </div>

      {/* 批量更新 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>3. 批量更新单元格</div>
        <Field className={styles.field} label="批量数据（每行格式: 行编号,列编号,文本内容）">
          <Textarea
            value={batchData}
            onChange={(e) => setBatchData(e.target.value)}
            placeholder={"示例:\n1,1,标题1\n1,2,标题2\n2,1,数据1\n2,2,数据2"}
            rows={6}
            disabled={loading}
          />
        </Field>

        <Button
          appearance="primary"
          onClick={handleBatchUpdate}
          disabled={loading || isUpdateDisabled}
          style={{ width: "100%" }}
        >
          {loading ? "批量更新中..." : "批量更新"}
        </Button>
      </div>

      {/* 消息提示 */}
      {message && (
        <div
          className={`${styles.message} ${
            messageType === "success"
              ? styles.messageSuccess
              : messageType === "error"
                ? styles.messageError
                : styles.messageInfo
          }`}
        >
          {messageType === "error" && "❌ "}
          {messageType === "success" && "✅ "}
          {messageType === "info" && "ℹ️ "}
          {message}
        </div>
      )}

      {isUpdateDisabled && (
        <div className={`${styles.message} ${styles.messageError}`}>
          ⚠️ 当前选中的元素不是表格，更新功能已禁用。请选择表格元素。
        </div>
      )}

      {/* 使用说明 */}
      <div className={styles.usageTips}>
        <div className={styles.sectionTitle}>使用说明:</div>
        <ul className={styles.tipsList}>
          <li>行列编号从 1 开始计数（第1行第1列）</li>
          <li>可以通过&quot;获取选中的表格&quot;按钮自动填充表格 ID</li>
          <li>也可以手动输入表格形状 ID 或使用表格索引</li>
          <li>批量更新时，每行一个单元格，格式: 行编号,列编号,文本内容</li>
          <li>文本内容中如果包含逗号，会被正确处理</li>
        </ul>
      </div>
    </div>
  );
};
