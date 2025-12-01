/**
 * 文件名: TableFormatUpdate.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 最后修改日期: 2025/12/1
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 表格格式更新组件 - 修改表格/单元格的格式属性
 * Description: Table format update component - modify table/cell format properties
 */

import React, { useState } from "react";
import {
  Button,
  Field,
  Input,
  tokens,
  makeStyles,
  Checkbox,
  Dropdown,
  Option,
} from "@fluentui/react-components";
import {
  updateCellFormat,
  updateRowFormat,
  updateColumnFormat,
  type CellFormatOptions,
  type RowFormatOptions,
  type ColumnFormatOptions,
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
  colorRow: {
    display: "flex",
    gap: "12px",
    alignItems: "flex-end",
    width: "100%",
  },
  colorInput: {
    flex: 1,
  },
  colorPreview: {
    width: "40px",
    height: "32px",
    borderRadius: "4px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
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

export const TableFormatUpdate: React.FC = () => {
  const styles = useStyles();
  
  // 表格定位
  const [shapeId, setShapeId] = useState<string>("");
  const [tableIndex, setTableIndex] = useState<string>("0");

  // 单元格/行/列索引
  const [rowIndex, setRowIndex] = useState<string>("1");
  const [columnIndex, setColumnIndex] = useState<string>("1");

  // 格式选项
  const [backgroundColor, setBackgroundColor] = useState<string>("");
  const [fontName, setFontName] = useState<string>("");
  const [fontSize, setFontSize] = useState<string>("");
  const [fontColor, setFontColor] = useState<string>("");
  const [fontBold, setFontBold] = useState<boolean>(false);
  const [fontItalic, setFontItalic] = useState<boolean>(false);
  const [fontUnderline, setFontUnderline] = useState<boolean>(false);
  const [borderWidth, setBorderWidth] = useState<string>("");
  const [borderColor, setBorderColor] = useState<string>("");
  const [horizontalAlignment, setHorizontalAlignment] = useState<"" | "Left" | "Center" | "Right" | "Justify">("");
  const [verticalAlignment, setVerticalAlignment] = useState<"" | "Top" | "Middle" | "Bottom">("");
  const [rowHeight, setRowHeight] = useState<string>("");
  const [columnWidth, setColumnWidth] = useState<string>("");

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

        shapes.load("items");
        await context.sync();

        const shape = shapes.items[0];
        shape.load("id,type,name");
        await context.sync();

        setShapeId(shape.id);
        setSelectedShapeType(shape.type);

        if (shape.type !== "Table") {
          setMessage(`警告: 选中的元素类型是 "${shape.type}"，不是表格。请选择表格元素。`);
          setMessageType("error");
          return;
        }

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

  // 构建格式选项对象
  const buildFormatOptions = () => {
    const options: Partial<CellFormatOptions> = {};
    
    if (backgroundColor.trim()) options.backgroundColor = backgroundColor.trim();
    if (fontName.trim()) options.fontName = fontName.trim();
    if (fontSize.trim()) options.fontSize = parseFloat(fontSize);
    if (fontColor.trim()) options.fontColor = fontColor.trim();
    if (fontBold) options.fontBold = true;
    if (fontItalic) options.fontItalic = true;
    if (fontUnderline) options.fontUnderline = true;
    if (borderWidth.trim()) options.borderWidth = parseFloat(borderWidth);
    if (borderColor.trim()) options.borderColor = borderColor.trim();
    if (horizontalAlignment) options.horizontalAlignment = horizontalAlignment;
    if (verticalAlignment) options.verticalAlignment = verticalAlignment;

    return options;
  };

  // 更新单元格格式
  const handleUpdateCellFormat = async () => {
    const row = parseInt(rowIndex);
    const col = parseInt(columnIndex);

    if (isNaN(row) || isNaN(col)) {
      setMessage("请输入有效的行列索引");
      setMessageType("error");
      return;
    }

    const formatOptions = buildFormatOptions();
    if (Object.keys(formatOptions).length === 0) {
      setMessage("请至少设置一个格式选项");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const tableLocation = shapeId.trim() ? { shapeId } : { tableIndex: parseInt(tableIndex) || 0 };
      // 转换为0开始的索引
      const result = await updateCellFormat(
        { rowIndex: row - 1, columnIndex: col - 1, ...formatOptions },
        tableLocation
      );

      if (result.success) {
        setMessage(`成功更新单元格 (${row}, ${col}) 的格式`);
        setMessageType("success");
      } else {
        setMessage(`更新失败: ${result.error || "未知错误"}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`更新单元格格式失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 更新整行格式
  const handleUpdateRowFormat = async () => {
    const row = parseInt(rowIndex);

    if (isNaN(row)) {
      setMessage("请输入有效的行索引");
      setMessageType("error");
      return;
    }

    const formatOptions = buildFormatOptions();
    if (Object.keys(formatOptions).length === 0 && !rowHeight.trim()) {
      setMessage("请至少设置一个格式选项");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const tableLocation = shapeId.trim() ? { shapeId } : { tableIndex: parseInt(tableIndex) || 0 };
      // 转换为0开始的索引
      const rowFormatOptions: RowFormatOptions = {
        rowIndex: row - 1,
        ...formatOptions,
      };
      
      if (rowHeight.trim()) {
        rowFormatOptions.height = parseFloat(rowHeight);
      }

      const result = await updateRowFormat(rowFormatOptions, tableLocation);

      if (result.success) {
        setMessage(`成功更新第 ${row} 行的格式，共更新 ${result.cellsUpdated} 个单元格`);
        setMessageType("success");
      } else {
        setMessage(`更新失败: ${result.error || "未知错误"}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`更新行格式失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 更新整列格式
  const handleUpdateColumnFormat = async () => {
    const col = parseInt(columnIndex);

    if (isNaN(col)) {
      setMessage("请输入有效的列索引");
      setMessageType("error");
      return;
    }

    const formatOptions = buildFormatOptions();
    if (Object.keys(formatOptions).length === 0 && !columnWidth.trim()) {
      setMessage("请至少设置一个格式选项");
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");
    try {
      const tableLocation = shapeId.trim() ? { shapeId } : { tableIndex: parseInt(tableIndex) || 0 };
      // 转换为0开始的索引
      const columnFormatOptions: ColumnFormatOptions = {
        columnIndex: col - 1,
        ...formatOptions,
      };
      
      if (columnWidth.trim()) {
        columnFormatOptions.width = parseFloat(columnWidth);
      }

      const result = await updateColumnFormat(columnFormatOptions, tableLocation);

      if (result.success) {
        setMessage(`成功更新第 ${col} 列的格式，共更新 ${result.cellsUpdated} 个单元格`);
        setMessageType("success");
      } else {
        setMessage(`更新失败: ${result.error || "未知错误"}`);
        setMessageType("error");
      }
    } catch (error) {
      setMessage(`更新列格式失败: ${error instanceof Error ? error.message : "未知错误"}`);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  };

  // 清空所有格式选项
  const handleClearFormats = () => {
    setBackgroundColor("");
    setFontName("");
    setFontSize("");
    setFontColor("");
    setFontBold(false);
    setFontItalic(false);
    setFontUnderline(false);
    setBorderWidth("");
    setBorderColor("");
    setHorizontalAlignment("");
    setVerticalAlignment("");
    setRowHeight("");
    setColumnWidth("");
    setMessage("已清空所有格式选项");
    setMessageType("info");
  };

  const isUpdateDisabled = selectedShapeType !== "" && selectedShapeType !== "Table";

  return (
    <div className={styles.container}>
      <div className={styles.title}>表格格式更新</div>
      <div className={styles.description}>修改表格单元格、行或列的格式属性（行列编号从 1 开始，留空则保持原样）</div>

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

      {/* 位置选择 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>2. 位置选择</div>
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
      </div>

      {/* 格式选项 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>3. 格式选项（留空保持原样）</div>
        
        {/* 背景色 */}
        <Field className={styles.field} label="背景色">
          <div className={styles.colorRow}>
            <Input
              className={styles.colorInput}
              type="text"
              value={backgroundColor}
              onChange={(e) => setBackgroundColor(e.target.value)}
              placeholder="#RRGGBB 例如: #FF0000"
              disabled={loading}
            />
            {backgroundColor && (
              <div
                className={styles.colorPreview}
                style={{ backgroundColor: backgroundColor }}
              />
            )}
          </div>
        </Field>

        {/* 字体 */}
        <div className={styles.row}>
          <Field className={styles.field} label="字体名称">
            <Input
              type="text"
              value={fontName}
              onChange={(e) => setFontName(e.target.value)}
              placeholder="例如: Arial, 微软雅黑"
              disabled={loading}
            />
          </Field>
          <Field className={styles.field} label="字体大小">
            <Input
              type="number"
              value={fontSize}
              onChange={(e) => setFontSize(e.target.value)}
              placeholder="磅"
              disabled={loading}
            />
          </Field>
        </div>

        {/* 字体颜色 */}
        <Field className={styles.field} label="字体颜色">
          <div className={styles.colorRow}>
            <Input
              className={styles.colorInput}
              type="text"
              value={fontColor}
              onChange={(e) => setFontColor(e.target.value)}
              placeholder="#RRGGBB 例如: #000000"
              disabled={loading}
            />
            {fontColor && (
              <div
                className={styles.colorPreview}
                style={{ backgroundColor: fontColor }}
              />
            )}
          </div>
        </Field>

        {/* 字体样式 */}
        <div className={styles.row}>
          <Checkbox
            checked={fontBold}
            onChange={(_, data) => setFontBold(data.checked as boolean)}
            disabled={loading}
            label="加粗"
          />
          <Checkbox
            checked={fontItalic}
            onChange={(_, data) => setFontItalic(data.checked as boolean)}
            disabled={loading}
            label="斜体"
          />
          <Checkbox
            checked={fontUnderline}
            onChange={(_, data) => setFontUnderline(data.checked as boolean)}
            disabled={loading}
            label="下划线"
          />
        </div>

        {/* 对齐方式 */}
        <div className={styles.row}>
          <Field className={styles.field} label="水平对齐">
            <Dropdown
              placeholder="选择对齐方式"
              value={horizontalAlignment}
              selectedOptions={horizontalAlignment ? [horizontalAlignment] : []}
              onOptionSelect={(_, data) => setHorizontalAlignment((data.optionValue || "") as "" | "Left" | "Center" | "Right" | "Justify")}
              disabled={loading}
            >
              <Option value="">不设置</Option>
              <Option value="Left">左对齐</Option>
              <Option value="Center">居中</Option>
              <Option value="Right">右对齐</Option>
              <Option value="Justify">两端对齐</Option>
            </Dropdown>
          </Field>
          <Field className={styles.field} label="垂直对齐">
            <Dropdown
              placeholder="选择对齐方式"
              value={verticalAlignment}
              selectedOptions={verticalAlignment ? [verticalAlignment] : []}
              onOptionSelect={(_, data) => setVerticalAlignment((data.optionValue || "") as "" | "Top" | "Middle" | "Bottom")}
              disabled={loading}
            >
              <Option value="">不设置</Option>
              <Option value="Top">顶部对齐</Option>
              <Option value="Middle">居中</Option>
              <Option value="Bottom">底部对齐</Option>
            </Dropdown>
          </Field>
        </div>

        {/* 行高/列宽 */}
        <div className={styles.row}>
          <Field className={styles.field} label="行高（磅）">
            <Input
              type="number"
              value={rowHeight}
              onChange={(e) => setRowHeight(e.target.value)}
              placeholder="仅用于整行更新"
              disabled={loading}
            />
          </Field>
          <Field className={styles.field} label="列宽（磅）">
            <Input
              type="number"
              value={columnWidth}
              onChange={(e) => setColumnWidth(e.target.value)}
              placeholder="仅用于整列更新"
              disabled={loading}
            />
          </Field>
        </div>

        <Button onClick={handleClearFormats} disabled={loading} style={{ width: "100%" }}>
          清空所有格式选项
        </Button>
      </div>

      {/* 操作按钮 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>4. 执行更新</div>
        <div className={styles.buttonGroup}>
          <Button
            appearance="primary"
            onClick={handleUpdateCellFormat}
            disabled={loading || isUpdateDisabled}
          >
            更新单元格
          </Button>
        </div>
        <div className={styles.buttonGroup}>
          <Button
            appearance="primary"
            onClick={handleUpdateRowFormat}
            disabled={loading || isUpdateDisabled}
          >
            更新整行
          </Button>
          <Button
            appearance="primary"
            onClick={handleUpdateColumnFormat}
            disabled={loading || isUpdateDisabled}
          >
            更新整列
          </Button>
        </div>
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
          <li>格式选项留空则保持原有格式不变</li>
          <li>颜色格式为 #RRGGBB，例如 #FF0000 表示红色</li>
          <li>更新单元格：只更新指定的单个单元格</li>
          <li>更新整行：更新指定行的所有单元格</li>
          <li>更新整列：更新指定列的所有单元格</li>
          <li>行高和列宽单位为磅（points）</li>
        </ul>
      </div>
    </div>
  );
};
