/**
 * 文件名: InsertTableDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: insertTable工具的调试组件
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
  Card,
  CardHeader,
  Text,
} from "@fluentui/react-components";
import {
  Table24Regular,
  TableAdd24Regular,
  TableDeleteRow24Regular,
  TableEdit24Regular,
  Info24Regular,
} from "@fluentui/react-icons";
import {
  insertTable,
  updateTable,
  updateCell,
  getTableInfo,
  getAllTablesInfo,
  deleteTable,
  addTableRows,
  addTableColumns,
  deleteTableRows,
  deleteTableColumns,
  mergeCells,
  type InsertTableOptions,
  type UpdateTableOptions,
  type UpdateCellOptions,
  type InsertLocation,
  type TableAlignment,
  type TableStyleType,
  type BorderStyle,
  type CellAlignment,
  type CellVerticalAlignment,
  type TableInfo,
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
  infoBox: {
    padding: "12px",
    backgroundColor: tokens.colorPaletteYellowBackground1,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: "12px",
    marginTop: "8px",
  },
  tableCard: {
    marginTop: "8px",
  },
  tableDataContainer: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
    maxHeight: "200px",
    overflowY: "auto",
  },
  sectionTitle: {
    marginTop: "16px",
    marginBottom: "8px",
    fontSize: "16px",
    fontWeight: 600,
  },
});

export const InsertTableDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState<"insert" | "update" | "query" | "delete">("insert");

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

  // 更新表格选项 / Update table options
  const [updateTableIndex, setUpdateTableIndex] = useState<string>("0");
  const [updateData, setUpdateData] = useState<string>("");

  // 更新单元格选项 / Update cell options
  const [cellTableIndex, setCellTableIndex] = useState<string>("0");
  const [cellRowIndex, setCellRowIndex] = useState<string>("0");
  const [cellColumnIndex, setCellColumnIndex] = useState<string>("0");
  const [cellContent, setCellContent] = useState<string>("");
  const [cellAlignment, setCellAlignment] = useState<CellAlignment>("Left");
  const [cellVerticalAlignment, setCellVerticalAlignment] = useState<CellVerticalAlignment>("Top");
  const [cellBackgroundColor, setCellBackgroundColor] = useState<string>("");
  const [cellFontSize, setCellFontSize] = useState<string>("");
  const [cellBold, setCellBold] = useState(false);

  // 查询结果 / Query results
  const [tableInfo, setTableInfo] = useState<TableInfo | null>(null);
  const [allTablesInfo, setAllTablesInfo] = useState<TableInfo[]>([]);
  const [queryTableIndex, setQueryTableIndex] = useState<string>("0");

  // 删除选项 / Delete options
  const [deleteTableIndex, setDeleteTableIndex] = useState<string>("0");

  // 行列操作选项 / Row/Column operation options
  const [rowColTableIndex, setRowColTableIndex] = useState<string>("0");
  const [rowColCount, setRowColCount] = useState<string>("1");
  const [rowColStartIndex, setRowColStartIndex] = useState<string>("0");

  // 合并单元格选项 / Merge cells options
  const [mergeTableIndex, setMergeTableIndex] = useState<string>("0");
  const [mergeStartRow, setMergeStartRow] = useState<string>("0");
  const [mergeStartCol, setMergeStartCol] = useState<string>("0");
  const [mergeEndRow, setMergeEndRow] = useState<string>("0");
  const [mergeEndCol, setMergeEndCol] = useState<string>("0");

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

  const handleUpdateTable = async () => {
    setLoading(true);
    setResult(null);

    try {
      // 如果索引为空，不传递 tableIndex 以使用选中的表格 / If index is empty, don't pass tableIndex to use selected table
      const index = updateTableIndex.trim() === "" ? undefined : parseInt(updateTableIndex);
      const options: UpdateTableOptions = {
        tableIndex: index,
        data: parseTableData(updateData),
        alignment,
        styleOptions: {
          styleType,
          firstRow,
          bandedRows,
        },
      };

      const updateResult = await updateTable(options);

      if (updateResult.success) {
        setResult({
          success: true,
          message: "表格更新成功！",
        });
      } else {
        setResult({
          success: false,
          message: `表格更新失败: ${updateResult.error}`,
        });
      }
    } catch (error) {
      console.error("更新表格失败:", error);
      setResult({
        success: false,
        message: `更新表格失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleUpdateCell = async () => {
    setLoading(true);
    setResult(null);

    try {
      const options: UpdateCellOptions = {
        tableIndex: parseInt(cellTableIndex),
        rowIndex: parseInt(cellRowIndex),
        columnIndex: parseInt(cellColumnIndex),
        content: cellContent || undefined,
        format: {
          alignment: cellAlignment,
          verticalAlignment: cellVerticalAlignment,
          backgroundColor: cellBackgroundColor || undefined,
          fontSize: cellFontSize ? parseFloat(cellFontSize) : undefined,
          bold: cellBold,
        },
      };

      const updateResult = await updateCell(options);

      if (updateResult.success) {
        setResult({
          success: true,
          message: "单元格更新成功！",
        });
      } else {
        setResult({
          success: false,
          message: `单元格更新失败: ${updateResult.error}`,
        });
      }
    } catch (error) {
      console.error("更新单元格失败:", error);
      setResult({
        success: false,
        message: `更新单元格失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleGetTableInfo = async () => {
    setLoading(true);
    setResult(null);

    try {
      // 如果索引为空，传递 undefined 以使用选中的表格 / If index is empty, pass undefined to use selected table
      const index = queryTableIndex.trim() === "" ? undefined : parseInt(queryTableIndex);
      const info = await getTableInfo(index);

      if (info) {
        setTableInfo(info);
        setResult({
          success: true,
          message: "表格信息获取成功！",
        });
      } else {
        setResult({
          success: false,
          message: "表格不存在或获取失败",
        });
      }
    } catch (error) {
      console.error("获取表格信息失败:", error);
      setResult({
        success: false,
        message: `获取表格信息失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleGetAllTablesInfo = async () => {
    setLoading(true);
    setResult(null);

    try {
      const info = await getAllTablesInfo();
      setAllTablesInfo(info);
      setResult({
        success: true,
        message: `获取成功！共 ${info.length} 个表格`,
      });
    } catch (error) {
      console.error("获取所有表格信息失败:", error);
      setResult({
        success: false,
        message: `获取所有表格信息失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleDeleteTable = async () => {
    setLoading(true);
    setResult(null);

    try {
      // 如果索引为空，传递 undefined 以使用选中的表格 / If index is empty, pass undefined to use selected table
      const index = deleteTableIndex.trim() === "" ? undefined : parseInt(deleteTableIndex);
      const deleteResult = await deleteTable(index);

      if (deleteResult.success) {
        setResult({
          success: true,
          message: "表格删除成功！",
        });
      } else {
        setResult({
          success: false,
          message: `表格删除失败: ${deleteResult.error}`,
        });
      }
    } catch (error) {
      console.error("删除表格失败:", error);
      setResult({
        success: false,
        message: `删除表格失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleAddRows = async () => {
    setLoading(true);
    setResult(null);

    try {
      const addResult = await addTableRows(
        parseInt(rowColTableIndex),
        parseInt(rowColCount),
        "End"
      );

      if (addResult.success) {
        setResult({
          success: true,
          message: `成功添加 ${rowColCount} 行！`,
        });
      } else {
        setResult({
          success: false,
          message: `添加行失败: ${addResult.error}`,
        });
      }
    } catch (error) {
      console.error("添加行失败:", error);
      setResult({
        success: false,
        message: `添加行失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleAddColumns = async () => {
    setLoading(true);
    setResult(null);

    try {
      const addResult = await addTableColumns(
        parseInt(rowColTableIndex),
        parseInt(rowColCount),
        "End"
      );

      if (addResult.success) {
        setResult({
          success: true,
          message: `成功添加 ${rowColCount} 列！`,
        });
      } else {
        setResult({
          success: false,
          message: `添加列失败: ${addResult.error}`,
        });
      }
    } catch (error) {
      console.error("添加列失败:", error);
      setResult({
        success: false,
        message: `添加列失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleDeleteRows = async () => {
    setLoading(true);
    setResult(null);

    try {
      const deleteResult = await deleteTableRows(
        parseInt(rowColTableIndex),
        parseInt(rowColStartIndex),
        parseInt(rowColCount)
      );

      if (deleteResult.success) {
        setResult({
          success: true,
          message: `成功删除 ${rowColCount} 行！`,
        });
      } else {
        setResult({
          success: false,
          message: `删除行失败: ${deleteResult.error}`,
        });
      }
    } catch (error) {
      console.error("删除行失败:", error);
      setResult({
        success: false,
        message: `删除行失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleDeleteColumns = async () => {
    setLoading(true);
    setResult(null);

    try {
      const deleteResult = await deleteTableColumns(
        parseInt(rowColTableIndex),
        parseInt(rowColStartIndex),
        parseInt(rowColCount)
      );

      if (deleteResult.success) {
        setResult({
          success: true,
          message: `成功删除 ${rowColCount} 列！`,
        });
      } else {
        setResult({
          success: false,
          message: `删除列失败: ${deleteResult.error}`,
        });
      }
    } catch (error) {
      console.error("删除列失败:", error);
      setResult({
        success: false,
        message: `删除列失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleMergeCells = async () => {
    setLoading(true);
    setResult(null);

    try {
      const mergeResult = await mergeCells(
        parseInt(mergeTableIndex),
        parseInt(mergeStartRow),
        parseInt(mergeStartCol),
        parseInt(mergeEndRow),
        parseInt(mergeEndCol)
      );

      if (mergeResult.success) {
        setResult({
          success: true,
          message: "单元格合并成功！",
        });
      } else {
        setResult({
          success: false,
          message: `单元格合并失败: ${mergeResult.error}`,
        });
      }
    } catch (error) {
      console.error("合并单元格失败:", error);
      setResult({
        success: false,
        message: `合并单元格失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      {/* 标签切换 / Tab switching */}
      <div className={styles.buttonGroup}>
        <Button
          appearance={activeTab === "insert" ? "primary" : "secondary"}
          onClick={() => setActiveTab("insert")}
          icon={<TableAdd24Regular />}
        >
          插入表格
        </Button>
        <Button
          appearance={activeTab === "update" ? "primary" : "secondary"}
          onClick={() => setActiveTab("update")}
          icon={<TableEdit24Regular />}
        >
          更新表格
        </Button>
        <Button
          appearance={activeTab === "query" ? "primary" : "secondary"}
          onClick={() => setActiveTab("query")}
          icon={<Info24Regular />}
        >
          查询表格
        </Button>
        <Button
          appearance={activeTab === "delete" ? "primary" : "secondary"}
          onClick={() => setActiveTab("delete")}
          icon={<TableDeleteRow24Regular />}
        >
          删除操作
        </Button>
      </div>

      {/* 插入表格 / Insert table */}
      {activeTab === "insert" && (
        <>
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
        </>
      )}

      {/* 更新表格 / Update table */}
      {activeTab === "update" && (
        <>
          <div className={styles.formRow}>
            <Label weight="semibold">更新整个表格</Label>
            <div className={styles.optionsContainer}>
              <Field label="表格索引" required>
                <Input
                  type="number"
                  value={updateTableIndex}
                  onChange={(e) => setUpdateTableIndex(e.target.value)}
                  placeholder="0"
                />
              </Field>

              <Field label="新数据（可选，每行用换行分隔，单元格用逗号分隔）">
                <Textarea
                  value={updateData}
                  onChange={(e) => setUpdateData(e.target.value)}
                  placeholder="新数据1,新数据2&#10;新数据3,新数据4"
                  rows={4}
                />
              </Field>

              <div className={styles.buttonGroup}>
                <Button
                  appearance="primary"
                  onClick={handleUpdateTable}
                  disabled={loading}
                  icon={loading ? <Spinner size="tiny" /> : undefined}
                >
                  {loading ? "正在更新..." : "更新表格"}
                </Button>
              </div>
            </div>
          </div>

          <div className={styles.formRow}>
            <Label weight="semibold">更新单个单元格</Label>
            <div className={styles.optionsContainer}>
              <div className={styles.formGrid3}>
                <Field label="表格索引" required>
                  <Input
                    type="number"
                    value={cellTableIndex}
                    onChange={(e) => setCellTableIndex(e.target.value)}
                    placeholder="0"
                  />
                </Field>

                <Field label="行索引" required>
                  <Input
                    type="number"
                    value={cellRowIndex}
                    onChange={(e) => setCellRowIndex(e.target.value)}
                    placeholder="0"
                  />
                </Field>

                <Field label="列索引" required>
                  <Input
                    type="number"
                    value={cellColumnIndex}
                    onChange={(e) => setCellColumnIndex(e.target.value)}
                    placeholder="0"
                  />
                </Field>
              </div>

              <Field label="单元格内容">
                <Input
                  value={cellContent}
                  onChange={(e) => setCellContent(e.target.value)}
                  placeholder="新内容"
                />
              </Field>

              <div className={styles.formGrid}>
                <Field label="水平对齐">
                  <Dropdown
                    value={cellAlignment}
                    selectedOptions={[cellAlignment]}
                    onOptionSelect={(_, data) => setCellAlignment(data.optionValue as CellAlignment)}
                  >
                    <Option value="Left">左对齐</Option>
                    <Option value="Centered">居中</Option>
                    <Option value="Right">右对齐</Option>
                  </Dropdown>
                </Field>

                <Field label="垂直对齐">
                  <Dropdown
                    value={cellVerticalAlignment}
                    selectedOptions={[cellVerticalAlignment]}
                    onOptionSelect={(_, data) =>
                      setCellVerticalAlignment(data.optionValue as CellVerticalAlignment)
                    }
                  >
                    <Option value="Top">顶部</Option>
                    <Option value="Center">居中</Option>
                    <Option value="Bottom">底部</Option>
                  </Dropdown>
                </Field>
              </div>

              <div className={styles.formGrid}>
                <Field label="背景色（可选）">
                  <input
                    type="color"
                    value={cellBackgroundColor}
                    onChange={(e) => setCellBackgroundColor(e.target.value)}
                    style={{
                      width: "100%",
                      height: "32px",
                      border: `1px solid ${tokens.colorNeutralStroke1}`,
                      borderRadius: tokens.borderRadiusMedium,
                      cursor: "pointer",
                    }}
                  />
                </Field>

                <Field label="字体大小（可选）">
                  <Input
                    type="number"
                    value={cellFontSize}
                    onChange={(e) => setCellFontSize(e.target.value)}
                    placeholder="12"
                  />
                </Field>
              </div>

              <Switch
                label="加粗"
                checked={cellBold}
                onChange={(_, data) => setCellBold(data.checked)}
              />

              <div className={styles.buttonGroup}>
                <Button
                  appearance="primary"
                  onClick={handleUpdateCell}
                  disabled={loading}
                  icon={loading ? <Spinner size="tiny" /> : undefined}
                >
                  {loading ? "正在更新..." : "更新单元格"}
                </Button>
              </div>
            </div>
          </div>

          <div className={styles.formRow}>
            <Label weight="semibold">行列操作</Label>
            <div className={styles.optionsContainer}>
              <div className={styles.formGrid3}>
                <Field label="表格索引" required>
                  <Input
                    type="number"
                    value={rowColTableIndex}
                    onChange={(e) => setRowColTableIndex(e.target.value)}
                    placeholder="0"
                  />
                </Field>

                <Field label="数量">
                  <Input
                    type="number"
                    value={rowColCount}
                    onChange={(e) => setRowColCount(e.target.value)}
                    placeholder="1"
                  />
                </Field>

                <Field label="起始索引（删除时）">
                  <Input
                    type="number"
                    value={rowColStartIndex}
                    onChange={(e) => setRowColStartIndex(e.target.value)}
                    placeholder="0"
                  />
                </Field>
              </div>

              <div className={styles.buttonGroup}>
                <Button appearance="secondary" onClick={handleAddRows} disabled={loading}>
                  添加行
                </Button>
                <Button appearance="secondary" onClick={handleAddColumns} disabled={loading}>
                  添加列
                </Button>
                <Button appearance="secondary" onClick={handleDeleteRows} disabled={loading}>
                  删除行
                </Button>
                <Button appearance="secondary" onClick={handleDeleteColumns} disabled={loading}>
                  删除列
                </Button>
              </div>
            </div>
          </div>

          <div className={styles.formRow}>
            <Label weight="semibold">合并单元格</Label>
            <div className={styles.optionsContainer}>
              <Field label="表格索引" required>
                <Input
                  type="number"
                  value={mergeTableIndex}
                  onChange={(e) => setMergeTableIndex(e.target.value)}
                  placeholder="0"
                />
              </Field>

              <div className={styles.formGrid}>
                <Field label="起始行">
                  <Input
                    type="number"
                    value={mergeStartRow}
                    onChange={(e) => setMergeStartRow(e.target.value)}
                    placeholder="0"
                  />
                </Field>

                <Field label="起始列">
                  <Input
                    type="number"
                    value={mergeStartCol}
                    onChange={(e) => setMergeStartCol(e.target.value)}
                    placeholder="0"
                  />
                </Field>

                <Field label="结束行">
                  <Input
                    type="number"
                    value={mergeEndRow}
                    onChange={(e) => setMergeEndRow(e.target.value)}
                    placeholder="1"
                  />
                </Field>

                <Field label="结束列">
                  <Input
                    type="number"
                    value={mergeEndCol}
                    onChange={(e) => setMergeEndCol(e.target.value)}
                    placeholder="1"
                  />
                </Field>
              </div>

              <div className={styles.buttonGroup}>
                <Button
                  appearance="primary"
                  onClick={handleMergeCells}
                  disabled={loading}
                  icon={loading ? <Spinner size="tiny" /> : undefined}
                >
                  {loading ? "正在合并..." : "合并单元格"}
                </Button>
              </div>
            </div>
          </div>
        </>
      )}

      {/* 查询表格 / Query table */}
      {activeTab === "query" && (
        <>
          <div className={styles.formRow}>
            <Label weight="semibold">查询单个表格</Label>
            <div className={styles.optionsContainer}>
              <Field label="表格索引">
                <Input
                  type="number"
                  value={queryTableIndex}
                  onChange={(e) => setQueryTableIndex(e.target.value)}
                  placeholder="0"
                />
              </Field>

              <div className={styles.buttonGroup}>
                <Button
                  appearance="primary"
                  onClick={handleGetTableInfo}
                  disabled={loading}
                  icon={loading ? <Spinner size="tiny" /> : undefined}
                >
                  {loading ? "正在查询..." : "查询表格"}
                </Button>
              </div>

              {tableInfo && (
                <Card className={styles.tableCard}>
                  <CardHeader
                    header={<Text weight="semibold">表格信息</Text>}
                  />
                  <div className={styles.tableDataContainer}>
                    <Text>索引: {tableInfo.index}</Text>
                    <br />
                    <Text>行数: {tableInfo.rowCount}</Text>
                    <br />
                    <Text>列数: {tableInfo.columnCount}</Text>
                    <br />
                    <Text>样式: {tableInfo.style}</Text>
                    <br />
                    <Text>对齐: {tableInfo.alignment}</Text>
                    <br />
                    {tableInfo.data && (
                      <>
                        <Text weight="semibold">数据:</Text>
                        <pre style={{ fontSize: "11px", marginTop: "8px" }}>
                          {JSON.stringify(tableInfo.data, null, 2)}
                        </pre>
                      </>
                    )}
                  </div>
                </Card>
              )}
            </div>
          </div>

          <div className={styles.formRow}>
            <Label weight="semibold">查询所有表格</Label>
            <div className={styles.optionsContainer}>
              <div className={styles.buttonGroup}>
                <Button
                  appearance="primary"
                  onClick={handleGetAllTablesInfo}
                  disabled={loading}
                  icon={loading ? <Spinner size="tiny" /> : undefined}
                >
                  {loading ? "正在查询..." : "查询所有表格"}
                </Button>
              </div>

              {allTablesInfo.length > 0 && (
                <div style={{ marginTop: "12px" }}>
                  {allTablesInfo.map((info) => (
                    <Card key={info.index} className={styles.tableCard}>
                      <CardHeader
                        header={<Text weight="semibold">表格 {info.index}</Text>}
                      />
                      <div className={styles.tableDataContainer}>
                        <Text>行数: {info.rowCount} | 列数: {info.columnCount}</Text>
                        <br />
                        <Text>样式: {info.style}</Text>
                      </div>
                    </Card>
                  ))}
                </div>
              )}
            </div>
          </div>
        </>
      )}

      {/* 删除操作 / Delete operations */}
      {activeTab === "delete" && (
        <>
          <div className={styles.formRow}>
            <Label weight="semibold">删除表格</Label>
            <div className={styles.optionsContainer}>
              <Field label="表格索引" required>
                <Input
                  type="number"
                  value={deleteTableIndex}
                  onChange={(e) => setDeleteTableIndex(e.target.value)}
                  placeholder="0"
                />
              </Field>

              <div className={styles.infoBox}>
                <Text>⚠️ 删除操作不可撤销，请谨慎操作！</Text>
              </div>

              <div className={styles.buttonGroup}>
                <Button
                  appearance="primary"
                  onClick={handleDeleteTable}
                  disabled={loading}
                  icon={loading ? <Spinner size="tiny" /> : <TableDeleteRow24Regular />}
                >
                  {loading ? "正在删除..." : "删除表格"}
                </Button>
              </div>
            </div>
          </div>
        </>
      )}

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
