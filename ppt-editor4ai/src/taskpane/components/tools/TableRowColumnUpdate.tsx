/**
 * 文件名: TableRowColumnUpdate.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 最后修改日期: 2025/12/1
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 表格行/列批量更新调试组件 - 批量修改整行或整列
 * Description: Table row/column batch update debug component - batch modify entire row or column
 */

import React, { useState } from "react";
import {
  Button,
  Field,
  Textarea,
  Input,
  tokens,
  makeStyles,
  Checkbox,
  Radio,
  RadioGroup,
} from "@fluentui/react-components";
import {
  updateTableRow,
  updateTableColumn,
  updateTableRowsBatch,
  updateTableColumnsBatch,
  getTableRowContent,
  getTableColumnContent,
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
  radioGroup: {
    display: "flex",
    gap: "20px",
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

export const TableRowColumnUpdate: React.FC = () => {
  const styles = useStyles();
  // 表格定位
  const [shapeId, setShapeId] = useState<string>("");
  const [tableIndex, setTableIndex] = useState<string>("0");

  // 操作类型
  const [operationType, setOperationType] = useState<"row" | "column">("row");

  // 单行/列更新
  const [index, setIndex] = useState<string>("1");
  const [values, setValues] = useState<string>("");
  const [skipEmpty, setSkipEmpty] = useState<boolean>(false);

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

  // 获取行/列内容
  const handleGetContent = async () => {
    const idx = parseInt(index);

    if (isNaN(idx)) {
      setMessage(`请输入有效的${operationType === "row" ? "行" : "列"}索引`);
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const tableLocation = shapeId.trim() ? { shapeId } : { tableIndex: parseInt(tableIndex) || 0 };

      let content: string[];
      if (operationType === "row") {
        // 转换为0开始的索引
        content = await getTableRowContent(idx - 1, tableLocation);
      } else {
        // 转换为0开始的索引
        content = await getTableColumnContent(idx - 1, tableLocation);
      }

      setValues(content.join(","));
      setMessage(
        `已获取${operationType === "row" ? "行" : "列"} ${idx} 的内容，共 ${content.length} 个单元格`
      );
      setMessageType("success");
    } catch (error) {
      setMessage(
        `获取${operationType === "row" ? "行" : "列"}内容失败: ${error instanceof Error ? error.message : "未知错误"}`
      );
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 更新单行/列
  const handleUpdate = async () => {
    const idx = parseInt(index);

    if (isNaN(idx)) {
      setMessage(`请输入有效的${operationType === "row" ? "行" : "列"}索引`);
      setMessageType("error");
      return;
    }

    if (!values.trim()) {
      setMessage("请输入数据（用逗号分隔）");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const tableLocation = shapeId.trim() ? { shapeId } : { tableIndex: parseInt(tableIndex) || 0 };
      const valueArray = values.split(",").map((v) => v.trim());

      let result;
      if (operationType === "row") {
        // 转换为0开始的索引
        result = await updateTableRow(
          { rowIndex: idx - 1, values: valueArray, skipEmpty },
          tableLocation
        );
      } else {
        // 转换为0开始的索引
        result = await updateTableColumn(
          { columnIndex: idx - 1, values: valueArray, skipEmpty },
          tableLocation
        );
      }

      if (result.success) {
        setMessage(
          `成功更新${operationType === "row" ? "行" : "列"} ${idx}，更新了 ${result.cellsUpdated} 个单元格。表格大小: ${result.rowCount} 行 × ${result.columnCount} 列`
        );
        setMessageType("success");
      } else {
        setMessage(`更新失败: ${result.error || "未知错误"}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(
        `更新${operationType === "row" ? "行" : "列"}失败: ${error instanceof Error ? error.message : "未知错误"}`
      );
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 批量更新
  const handleBatchUpdate = async () => {
    if (!batchData.trim()) {
      setMessage("请输入批量更新数据");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const tableLocation = shapeId.trim() ? { shapeId } : { tableIndex: parseInt(tableIndex) || 0 };
      const lines = batchData.trim().split("\n");

      if (operationType === "row") {
        // 批量更新行
        const rows = [];
        for (let i = 0; i < lines.length; i++) {
          const line = lines[i].trim();
          if (!line) continue;

          const parts = line.split(":");
          if (parts.length < 2) {
            setMessage(`第 ${i + 1} 行格式错误，应为: 行编号:值1,值2,值3`);
            setMessageType("error");
            return;
          }

          const rowIndex = parseInt(parts[0].trim());
          if (isNaN(rowIndex)) {
            setMessage(`第 ${i + 1} 行的行编号无效`);
            setMessageType("error");
            return;
          }

          const values = parts[1].split(",").map((v) => v.trim());
          // 转换为0开始的索引
          rows.push({ rowIndex: rowIndex - 1, values, skipEmpty });
        }

        if (rows.length === 0) {
          setMessage("没有有效的更新数据");
          setMessageType("error");
          return;
        }

        const result = await updateTableRowsBatch({ rows }, tableLocation);

        if (result.success) {
          setMessage(
            `成功批量更新 ${rows.length} 行，共更新 ${result.cellsUpdated} 个单元格。表格大小: ${result.rowCount} 行 × ${result.columnCount} 列`
          );
          setMessageType("success");
        } else {
          setMessage(`批量更新失败: ${result.error || "未知错误"}`);
          setMessageType("error");
        }
      } else {
        // 批量更新列
        const columns = [];
        for (let i = 0; i < lines.length; i++) {
          const line = lines[i].trim();
          if (!line) continue;

          const parts = line.split(":");
          if (parts.length < 2) {
            setMessage(`第 ${i + 1} 行格式错误，应为: 列编号:值1,值2,值3`);
            setMessageType("error");
            return;
          }

          const columnIndex = parseInt(parts[0].trim());
          if (isNaN(columnIndex)) {
            setMessage(`第 ${i + 1} 行的列编号无效`);
            setMessageType("error");
            return;
          }

          const values = parts[1].split(",").map((v) => v.trim());
          // 转换为0开始的索引
          columns.push({ columnIndex: columnIndex - 1, values, skipEmpty });
        }

        if (columns.length === 0) {
          setMessage("没有有效的更新数据");
          setMessageType("error");
          return;
        }

        const result = await updateTableColumnsBatch({ columns }, tableLocation);

        if (result.success) {
          setMessage(
            `成功批量更新 ${columns.length} 列，共更新 ${result.cellsUpdated} 个单元格。表格大小: ${result.rowCount} 行 × ${result.columnCount} 列`
          );
          setMessageType("success");
        } else {
          setMessage(`批量更新失败: ${result.error || "未知错误"}`);
          setMessageType("error");
        }
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
      <div className={styles.title}>表格行/列批量更新</div>
      <div className={styles.description}>批量修改表格的整行或整列内容（行列编号从 1 开始）</div>

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

      {/* 操作类型选择 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>2. 操作类型</div>
        <RadioGroup
          value={operationType}
          onChange={(_, data) => setOperationType(data.value as "row" | "column")}
          disabled={loading}
          className={styles.radioGroup}
        >
          <Radio value="row" label="行操作" />
          <Radio value="column" label="列操作" />
        </RadioGroup>
      </div>

      {/* 单行/列更新 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>3. 单{operationType === "row" ? "行" : "列"}更新</div>
        <Field className={styles.field} label={`${operationType === "row" ? "行" : "列"}索引`}>
          <Input
            type="number"
            value={index}
            onChange={(e) => setIndex(e.target.value)}
            disabled={loading}
          />
        </Field>

        <Field className={styles.field} label="数据（用逗号分隔）">
          <Textarea
            value={values}
            onChange={(e) => setValues(e.target.value)}
            placeholder={`输入${operationType === "row" ? "行" : "列"}数据，例如: 值1,值2,值3`}
            rows={3}
            disabled={loading}
          />
        </Field>

        <Checkbox
          checked={skipEmpty}
          onChange={(_, data) => setSkipEmpty(data.checked as boolean)}
          disabled={loading}
          label="跳过空值（不更新空单元格）"
        />

        <div className={styles.buttonGroup}>
          <Button onClick={handleGetContent} disabled={loading || isUpdateDisabled}>
            获取{operationType === "row" ? "行" : "列"}内容
          </Button>
          <Button appearance="primary" onClick={handleUpdate} disabled={loading || isUpdateDisabled}>
            {loading ? "更新中..." : `更新${operationType === "row" ? "行" : "列"}`}
          </Button>
        </div>
      </div>

      {/* 批量更新 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>4. 批量更新多{operationType === "row" ? "行" : "列"}</div>
        <Field
          className={styles.field}
          label={`批量数据（每行格式: ${operationType === "row" ? "行" : "列"}编号:值1,值2,值3）`}
        >
          <Textarea
            value={batchData}
            onChange={(e) => setBatchData(e.target.value)}
            placeholder={
              operationType === "row"
                ? "示例:\n1:标题1,标题2,标题3\n2:数据1,数据2,数据3\n3:数据4,数据5,数据6"
                : "示例:\n1:姓名,张三,李四\n2:年龄,25,30\n3:城市,北京,上海"
            }
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
          <li>可以通过"获取选中的表格"按钮自动填充表格 ID</li>
          <li>单行/列更新时，数据用逗号分隔，例如: 值1,值2,值3</li>
          <li>批量更新时，每行一个行/列，格式: 编号:值1,值2,值3</li>
          <li>启用"跳过空值"后，空字符串不会更新到单元格</li>
          <li>如果数据长度超过表格大小，多余数据会被截断</li>
        </ul>
      </div>
    </div>
  );
};
