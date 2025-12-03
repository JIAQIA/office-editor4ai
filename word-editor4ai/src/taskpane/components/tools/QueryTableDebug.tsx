/**
 * 文件名: QueryTableDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 表格查询工具的调试组件 / Debug component for table query tool
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
  Card,
  CardHeader,
  Text,
} from "@fluentui/react-components";
import { getTableInfo, getAllTablesInfo, type TableInfo } from "../../../word-tools";

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
});

export const QueryTableDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);

  // 查询结果 / Query results
  const [tableInfo, setTableInfo] = useState<TableInfo | null>(null);
  const [allTablesInfo, setAllTablesInfo] = useState<TableInfo[]>([]);
  const [queryTableIndex, setQueryTableIndex] = useState<string>("0");

  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

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

  return (
    <div className={styles.container}>
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
              <CardHeader header={<Text weight="semibold">表格信息</Text>} />
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
                  <CardHeader header={<Text weight="semibold">表格 {info.index}</Text>} />
                  <div className={styles.tableDataContainer}>
                    <Text>
                      行数: {info.rowCount} | 列数: {info.columnCount}
                    </Text>
                    <br />
                    <Text>样式: {info.style}</Text>
                  </div>
                </Card>
              ))}
            </div>
          )}
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
