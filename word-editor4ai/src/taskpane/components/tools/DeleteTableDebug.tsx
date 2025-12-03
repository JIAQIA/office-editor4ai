/**
 * 文件名: DeleteTableDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 表格删除工具的调试组件 / Debug component for table deletion tool
 */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  Input,
  Label,
  makeStyles,
  Spinner,
  tokens,
  Field,
  Text,
} from "@fluentui/react-components";
import { TableDeleteRow24Regular } from "@fluentui/react-icons";
import { deleteTable } from "../../../word-tools";

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
});

export const DeleteTableDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);

  // 删除选项 / Delete options
  const [deleteTableIndex, setDeleteTableIndex] = useState<string>("0");

  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

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

  return (
    <div className={styles.container}>
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
