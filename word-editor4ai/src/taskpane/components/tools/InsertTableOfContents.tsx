/**
 * 文件名: InsertTableOfContents.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 插入目录组件 / Insert Table of Contents Component
 */

import * as React from "react";
import { useState } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Input,
  Label,
  Checkbox,
  Card,
  Spinner,
} from "@fluentui/react-components";
import {
  insertTableOfContents,
  type InsertLocation,
  type TOCOptions,
} from "../../../word-tools/tableOfContents";

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
  buttonGroup: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
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

export const InsertTableOfContents: React.FC = () => {
  const styles = useStyles();

  // 插入目录状态 / Insert TOC state
  const [insertLocation, setInsertLocation] = useState<InsertLocation>("Start");
  const [tocTitle, setTocTitle] = useState("目录");
  const [tocLevels, setTocLevels] = useState("1,2,3");
  const [showPageNumbers, setShowPageNumbers] = useState(true);
  const [rightAlignPageNumbers, setRightAlignPageNumbers] = useState(true);
  const [useHyperlinks, setUseHyperlinks] = useState(true);
  const [includeHidden, setIncludeHidden] = useState(false);

  // 结果状态 / Result state
  const [result, setResult] = useState<string>("");
  const [isLoading, setIsLoading] = useState(false);

  // 插入目录 / Insert TOC
  const handleInsertTOC = async () => {
    setIsLoading(true);
    setResult("");

    try {
      // 解析标题级别 / Parse heading levels
      const levels = tocLevels
        .split(",")
        .map((l) => parseInt(l.trim()))
        .filter((l) => !isNaN(l) && l >= 1 && l <= 9);

      if (levels.length === 0) {
        setResult("错误：请输入有效的标题级别（1-9）");
        setIsLoading(false);
        return;
      }

      const options: TOCOptions = {
        title: tocTitle || undefined,
        levels,
        showPageNumbers,
        rightAlignPageNumbers,
        useHyperlinks,
        includeHidden,
      };

      const insertResult = await insertTableOfContents(insertLocation, options);

      if (insertResult.success) {
        setResult(
          `✓ 目录插入成功！\n\n` +
            `位置: ${insertLocation}\n` +
            `标题: ${tocTitle || "无"}\n` +
            `包含级别: ${levels.join(", ")}\n` +
            `条目数量: ${insertResult.tocInfo?.entryCount || 0}\n` +
            `目录索引: ${insertResult.tocInfo?.index}`
        );
      } else {
        setResult(`✗ 插入失败: ${insertResult.error}`);
      }
    } catch (error) {
      setResult(`✗ 插入失败: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <Card>
        <div className={styles.section}>
          <h3>插入目录</h3>

          <div className={styles.field}>
            <Label>插入位置</Label>
            <div className={styles.buttonGroup}>
              {(["Start", "End", "Before", "After", "Replace"] as InsertLocation[]).map((loc) => (
                <Button
                  key={loc}
                  appearance={insertLocation === loc ? "primary" : "secondary"}
                  onClick={() => setInsertLocation(loc)}
                >
                  {loc}
                </Button>
              ))}
            </div>
          </div>

          <div className={styles.field}>
            <Label htmlFor="toc-title">目录标题（可选）</Label>
            <Input
              id="toc-title"
              value={tocTitle}
              onChange={(e) => setTocTitle(e.target.value)}
              placeholder="目录"
            />
          </div>

          <div className={styles.field}>
            <Label htmlFor="toc-levels">包含的标题级别（1-9，逗号分隔）</Label>
            <Input
              id="toc-levels"
              value={tocLevels}
              onChange={(e) => setTocLevels(e.target.value)}
              placeholder="1,2,3"
            />
          </div>

          <div className={styles.field}>
            <Checkbox
              checked={showPageNumbers}
              onChange={(_e, data) => setShowPageNumbers(data.checked === true)}
              label="显示页码"
            />
          </div>

          <div className={styles.field}>
            <Checkbox
              checked={rightAlignPageNumbers}
              onChange={(_e, data) => setRightAlignPageNumbers(data.checked === true)}
              label="页码右对齐"
            />
          </div>

          <div className={styles.field}>
            <Checkbox
              checked={useHyperlinks}
              onChange={(_e, data) => setUseHyperlinks(data.checked === true)}
              label="使用超链接"
            />
          </div>

          <div className={styles.field}>
            <Checkbox
              checked={includeHidden}
              onChange={(_e, data) => setIncludeHidden(data.checked === true)}
              label="包含隐藏文本"
            />
          </div>

          <Button appearance="primary" onClick={handleInsertTOC} disabled={isLoading}>
            {isLoading ? <Spinner size="tiny" /> : "插入目录"}
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
