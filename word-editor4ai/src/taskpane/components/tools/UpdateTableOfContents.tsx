/**
 * 文件名: UpdateTableOfContents.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 更新目录组件 / Update Table of Contents Component
 */

import * as React from "react";
import { useState } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Input,
  Label,
  Card,
  Spinner,
} from "@fluentui/react-components";
import { updateTableOfContents } from "../../../word-tools/tableOfContents";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  field: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    flex: 1,
  },
  result: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    fontSize: tokens.fontSizeBase200,
    fontFamily: "monospace",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    maxHeight: "400px",
    overflowY: "auto",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
  },
  success: {
    color: tokens.colorPaletteGreenForeground1,
  },
});

export const UpdateTableOfContents: React.FC = () => {
  const styles = useStyles();

  // 更新目录状态 / Update TOC state
  const [tocIndex, setTocIndex] = useState("");

  // 结果状态 / Result state
  const [result, setResult] = useState<string>("");
  const [isLoading, setIsLoading] = useState(false);

  // 更新目录 / Update TOC
  const handleUpdateTOC = async () => {
    setIsLoading(true);
    setResult("");

    try {
      const index = tocIndex ? parseInt(tocIndex) : undefined;
      const updateResult = await updateTableOfContents(index);

      if (updateResult.success) {
        setResult(
          `✓ 目录更新成功！\n\n` +
            `更新数量: ${updateResult.updatedCount}\n` +
            `${index !== undefined ? `目录索引: ${index}` : "更新了所有目录"}`
        );
      } else {
        setResult(`✗ 更新失败: ${updateResult.error}`);
      }
    } catch (error) {
      setResult(`✗ 更新失败: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <Card>
        <div className={styles.section}>
          <h3>更新目录</h3>

          <div className={styles.field}>
            <Label htmlFor="update-toc-index">目录索引（留空则更新所有）</Label>
            <Input
              id="update-toc-index"
              type="number"
              value={tocIndex}
              onChange={(e) => setTocIndex(e.target.value)}
              placeholder="0"
            />
          </div>

          <Button appearance="primary" onClick={handleUpdateTOC} disabled={isLoading}>
            {isLoading ? <Spinner size="tiny" /> : "更新目录"}
          </Button>
        </div>
      </Card>

      {/* 结果显示 / Result display */}
      {result && (
        <Card>
          <div className={styles.section}>
            <h3>执行结果</h3>
            <div
              className={`${styles.result} ${
                result.startsWith("✓") ? styles.success : result.startsWith("✗") ? styles.error : ""
              }`}
            >
              {result}
            </div>
          </div>
        </Card>
      )}
    </div>
  );
};
