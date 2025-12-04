/**
 * 文件名: QueryTableOfContents.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 查询目录组件 / Query Table of Contents Component
 */

import * as React from "react";
import { useState } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Card,
  Spinner,
} from "@fluentui/react-components";
import { getTableOfContentsList, type TOCInfo } from "../../../word-tools/tableOfContents";

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
  tocItem: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: "8px",
  },
  tocItemTitle: {
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "8px",
  },
  tocItemContent: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
});

export const QueryTableOfContents: React.FC = () => {
  const styles = useStyles();

  // 结果状态 / Result state
  const [result, setResult] = useState<string>("");
  const [isLoading, setIsLoading] = useState(false);
  const [tocList, setTocList] = useState<TOCInfo[]>([]);

  // 获取目录列表 / Get TOC list
  const handleGetTOCList = async () => {
    setIsLoading(true);
    setResult("");

    try {
      const listResult = await getTableOfContentsList();

      if (listResult.success && listResult.tocs) {
        setTocList(listResult.tocs);
        setResult(`✓ 找到 ${listResult.tocs.length} 个目录`);
      } else {
        setResult(`✗ 获取失败: ${listResult.error}`);
        setTocList([]);
      }
    } catch (error) {
      setResult(`✗ 获取失败: ${error instanceof Error ? error.message : String(error)}`);
      setTocList([]);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <Card>
        <div className={styles.section}>
          <h3>获取目录列表</h3>

          <Button appearance="primary" onClick={handleGetTOCList} disabled={isLoading}>
            {isLoading ? <Spinner size="tiny" /> : "获取目录列表"}
          </Button>

          {tocList.length > 0 && (
            <div>
              <h4>目录列表（共 {tocList.length} 个）</h4>
              {tocList.map((toc) => (
                <div key={toc.index} className={styles.tocItem}>
                  <div className={styles.tocItemTitle}>目录 #{toc.index}</div>
                  <div className={styles.tocItemContent}>
                    <div>条目数量: {toc.entryCount}</div>
                    <div>包含级别: {toc.levels.join(", ")}</div>
                    <div>预览: {toc.text.substring(0, 100)}...</div>
                  </div>
                </div>
              ))}
            </div>
          )}
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
