/**
 * 文件名: CreateTableDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 表格创建工具的调试组件 / Debug component for table creation tool
 */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  Input,
  Label,
  makeStyles,
  Spinner,
  Switch,
  Dropdown,
  Option,
  tokens,
  Field,
  Textarea,
} from "@fluentui/react-components";
import { Table24Regular } from "@fluentui/react-icons";
import {
  insertTable,
  type InsertTableOptions,
  type InsertLocation,
  type TableAlignment,
  type TableStyleType,
  type BorderStyle,
} from "../../../word-tools";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  formRow: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  formGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "12px",
  },
  formGrid3: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr 1fr",
    gap: "12px",
  },
  optionsContainer: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "8px",
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
  },
  resultMessage: {
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "12px",
  },
  success: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
  },
  error: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
  },
});

export const CreateTableDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);

  // 插入表格选项 / Insert table options
  const [rows, setRows] = useState<string>("3");
  const [cols, setCols] = useState<string>("3");
  const [insertLocation, setInsertLocation] = useState<InsertLocation>("End");
  const [tableData, setTableData] = useState<string>("");
  const [headerRow, setHeaderRow] = useState<string>("");
  const [columnWidths, setColumnWidths] = useState<string>("");
  const [alignment, setAlignment] = useState<TableAlignment>("Left");
  const [styleType, setStyleType] = useState<TableStyleType>(Word.BuiltInStyleName.tableGrid);
  const [firstRow, setFirstRow] = useState(true);
  const [bandedRows, setBandedRows] = useState(false);
  const [borderStyle, setBorderStyle] = useState<BorderStyle>("Single");
  const [borderWidth, setBorderWidth] = useState<string>("1");
  const [borderColor, setBorderColor] = useState<string>("#000000");

  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

  const parseTableData = (dataStr: string): string[][] | undefined => {
    if (!dataStr.trim()) return undefined;
    try {
      return dataStr.split("\n").map((row) => row.split(",").map((cell) => cell.trim()));
    } catch {
      return undefined;
    }
  };

  const parseHeaderRow = (headerStr: string): string[] | undefined => {
    if (!headerStr.trim()) return undefined;
    return headerStr.split(",").map((cell) => cell.trim());
  };

  const parseColumnWidths = (widthsStr: string): number | number[] | undefined => {
    if (!widthsStr.trim()) return undefined;
    const widths = widthsStr.split(",").map((w) => parseFloat(w.trim()));
    return widths.length === 1 ? widths[0] : widths;
  };

  const handleInsertTable = async () => {
    setLoading(true);
    setResult(null);

    try {
      const options: InsertTableOptions = {
        rows: parseInt(rows),
        cols: parseInt(cols),
        insertLocation,
        data: parseTableData(tableData),
        headerRow: parseHeaderRow(headerRow),
        columnWidths: parseColumnWidths(columnWidths),
        alignment,
        styleOptions: {
          styleType,
          firstRow,
          bandedRows,
        },
        borderOptions: {
          style: borderStyle,
          width: parseFloat(borderWidth),
          color: borderColor,
        },
      };

      const insertResult = await insertTable(options);

      if (insertResult.success) {
        setResult({
          success: true,
          message: `表格插入成功！索引: ${insertResult.tableIndex}`,
        });
      } else {
        setResult({
          success: false,
          message: `表格插入失败: ${insertResult.error}`,
        });
      }
    } catch (error) {
      console.error("插入表格失败:", error);
      setResult({
        success: false,
        message: `插入表格失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.formRow}>
        <Label weight="semibold">基本设置</Label>
        <div className={styles.optionsContainer}>
          <div className={styles.formGrid}>
            <Field label="行数" required>
              <Input
                type="number"
                value={rows}
                onChange={(e) => setRows(e.target.value)}
                placeholder="3"
              />
            </Field>

            <Field label="列数" required>
              <Input
                type="number"
                value={cols}
                onChange={(e) => setCols(e.target.value)}
                placeholder="3"
              />
            </Field>
          </div>

          <div className={styles.formGrid}>
            <Field label="插入位置">
              <Dropdown
                value={insertLocation}
                selectedOptions={[insertLocation]}
                onOptionSelect={(_, data) => setInsertLocation(data.optionValue as InsertLocation)}
              >
                <Option value="Start">文档开头</Option>
                <Option value="End">文档末尾</Option>
                <Option value="Before">选区之前</Option>
                <Option value="After">选区之后</Option>
                <Option value="Replace">替换选区</Option>
              </Dropdown>
            </Field>

            <Field label="对齐方式">
              <Dropdown
                value={alignment}
                selectedOptions={[alignment]}
                onOptionSelect={(_, data) => setAlignment(data.optionValue as TableAlignment)}
              >
                <Option value="Left">左对齐</Option>
                <Option value="Centered">居中</Option>
                <Option value="Right">右对齐</Option>
                <Option value="Justified">两端对齐</Option>
              </Dropdown>
            </Field>
          </div>

          <Field label="表头行（可选，用逗号分隔）">
            <Input
              value={headerRow}
              onChange={(e) => setHeaderRow(e.target.value)}
              placeholder="姓名,年龄,城市"
            />
          </Field>

          <Field label="表格数据（可选，每行用换行分隔，单元格用逗号分隔）">
            <Textarea
              value={tableData}
              onChange={(e) => setTableData(e.target.value)}
              placeholder="张三,25,北京&#10;李四,30,上海"
              rows={4}
            />
          </Field>

          <Field label="列宽（可选，单位：磅，用逗号分隔或单个值）">
            <Input
              value={columnWidths}
              onChange={(e) => setColumnWidths(e.target.value)}
              placeholder="100,150,200 或 120"
            />
          </Field>
        </div>
      </div>

      <div className={styles.formRow}>
        <Label weight="semibold">样式设置</Label>
        <div className={styles.optionsContainer}>
          <Field label="表格样式">
            <Dropdown
              value={styleType}
              selectedOptions={[styleType]}
              onOptionSelect={(_, data) => setStyleType(data.optionValue as TableStyleType)}
            >
              <Option value={Word.BuiltInStyleName.tableGrid}>网格</Option>
              <Option value={Word.BuiltInStyleName.plainTable1}>简单样式1</Option>
              <Option value={Word.BuiltInStyleName.plainTable2}>简单样式2</Option>
              <Option value={Word.BuiltInStyleName.gridTable1Light}>浅色网格1</Option>
              <Option value={Word.BuiltInStyleName.gridTable2}>网格2</Option>
              <Option value={Word.BuiltInStyleName.gridTable3}>网格3</Option>
              <Option value={Word.BuiltInStyleName.gridTable4}>网格4</Option>
              <Option value={Word.BuiltInStyleName.gridTable5Dark}>深色网格5</Option>
              <Option value={Word.BuiltInStyleName.gridTable6Colorful}>彩色网格6</Option>
              <Option value={Word.BuiltInStyleName.gridTable7Colorful}>彩色网格7</Option>
              <Option value={Word.BuiltInStyleName.listTable1Light}>浅色列表1</Option>
              <Option value={Word.BuiltInStyleName.listTable2}>列表2</Option>
              <Option value={Word.BuiltInStyleName.listTable3}>列表3</Option>
              <Option value={Word.BuiltInStyleName.listTable4}>列表4</Option>
              <Option value={Word.BuiltInStyleName.listTable5Dark}>深色列表5</Option>
              <Option value={Word.BuiltInStyleName.listTable6Colorful}>彩色列表6</Option>
              <Option value={Word.BuiltInStyleName.listTable7Colorful}>彩色列表7</Option>
            </Dropdown>
          </Field>

          <Switch
            label="首行特殊格式"
            checked={firstRow}
            onChange={(_, data) => setFirstRow(data.checked)}
          />

          <Switch
            label="条纹行"
            checked={bandedRows}
            onChange={(_, data) => setBandedRows(data.checked)}
          />
        </div>
      </div>

      <div className={styles.formRow}>
        <Label weight="semibold">边框设置</Label>
        <div className={styles.optionsContainer}>
          <div className={styles.formGrid3}>
            <Field label="边框样式">
              <Dropdown
                value={borderStyle}
                selectedOptions={[borderStyle]}
                onOptionSelect={(_, data) => setBorderStyle(data.optionValue as BorderStyle)}
              >
                <Option value="Single">单线</Option>
                <Option value="Dotted">点线</Option>
                <Option value="Dashed">虚线</Option>
                <Option value="Double">双线</Option>
                <Option value="None">无边框</Option>
              </Dropdown>
            </Field>

            <Field label="边框宽度（磅）">
              <Input
                type="number"
                value={borderWidth}
                onChange={(e) => setBorderWidth(e.target.value)}
                placeholder="1"
              />
            </Field>

            <Field label="边框颜色">
              <input
                type="color"
                value={borderColor}
                onChange={(e) => setBorderColor(e.target.value)}
                style={{
                  width: "100%",
                  height: "32px",
                  border: `1px solid ${tokens.colorNeutralStroke1}`,
                  borderRadius: tokens.borderRadiusMedium,
                  cursor: "pointer",
                }}
              />
            </Field>
          </div>
        </div>
      </div>

      <div className={styles.buttonGroup}>
        <Button
          appearance="primary"
          onClick={handleInsertTable}
          disabled={loading}
          icon={loading ? <Spinner size="tiny" /> : <Table24Regular />}
        >
          {loading ? "正在插入..." : "插入表格"}
        </Button>
      </div>

      {/* 结果显示 / Result display */}
      {result && (
        <div
          className={`${styles.resultMessage} ${result.success ? styles.success : styles.error}`}
        >
          {result.message}
        </div>
      )}
    </div>
  );
};
