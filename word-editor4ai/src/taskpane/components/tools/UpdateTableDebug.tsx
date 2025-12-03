/**
 * 文件名: UpdateTableDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 表格更新工具的调试组件 / Debug component for table update tool
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
import {
  updateTable,
  updateCell,
  addTableRows,
  addTableColumns,
  deleteTableRows,
  deleteTableColumns,
  mergeCells,
  type UpdateTableOptions,
  type UpdateCellOptions,
  type TableAlignment,
  type TableStyleType,
  type CellAlignment,
  type CellVerticalAlignment,
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
  sectionTitle: {
    marginTop: "16px",
    marginBottom: "8px",
    fontSize: "16px",
    fontWeight: 600,
  },
});

export const UpdateTableDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);

  // 更新表格选项 / Update table options
  const [updateTableIndex, setUpdateTableIndex] = useState<string>("0");
  const [updateData, setUpdateData] = useState<string>("");
  const [alignment, setAlignment] = useState<TableAlignment>("Left");
  const [styleType, setStyleType] = useState<TableStyleType>(Word.BuiltInStyleName.tableGrid);
  const [firstRow, setFirstRow] = useState(true);
  const [bandedRows, setBandedRows] = useState(false);

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

  const handleAddRows = async () => {
    setLoading(true);
    setResult(null);

    try {
      // 如果索引为空，传递 undefined 以使用选中的表格 / If index is empty, pass undefined to use selected table
      const index = rowColTableIndex.trim() === "" ? undefined : parseInt(rowColTableIndex);
      const addResult = await addTableRows(
        index,
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
      // 如果索引为空，传递 undefined 以使用选中的表格 / If index is empty, pass undefined to use selected table
      const index = rowColTableIndex.trim() === "" ? undefined : parseInt(rowColTableIndex);
      const addResult = await addTableColumns(
        index,
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
      // 如果索引为空，传递 undefined 以使用选中的表格 / If index is empty, pass undefined to use selected table
      const index = rowColTableIndex.trim() === "" ? undefined : parseInt(rowColTableIndex);
      const deleteResult = await deleteTableRows(
        index,
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
      // 如果索引为空，传递 undefined 以使用选中的表格 / If index is empty, pass undefined to use selected table
      const index = rowColTableIndex.trim() === "" ? undefined : parseInt(rowColTableIndex);
      const deleteResult = await deleteTableColumns(
        index,
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
      // 如果索引为空，传递 undefined 以使用选中的表格 / If index is empty, pass undefined to use selected table
      const index = mergeTableIndex.trim() === "" ? undefined : parseInt(mergeTableIndex);
      const mergeResult = await mergeCells(
        index,
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
            <Field label="表格索引（留空使用选中表格）">
              <Input
                type="number"
                value={rowColTableIndex}
                onChange={(e) => setRowColTableIndex(e.target.value)}
                placeholder="留空使用选中表格"
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
          <Field label="表格索引（留空使用选中表格）">
            <Input
              type="number"
              value={mergeTableIndex}
              onChange={(e) => setMergeTableIndex(e.target.value)}
              placeholder="留空使用选中表格"
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
