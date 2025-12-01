/**
 * 文件名: DocumentStats.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 文档统计信息工具，用于获取并显示文档的字数、段落数、页数等统计信息
 */

/* global console */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  makeStyles,
  tokens,
  Spinner,
  Switch,
  Label,
  Card,
  Divider,
  Badge,
} from "@fluentui/react-components";
import {
  DocumentData24Regular,
  ArrowDownload24Regular,
  DocumentText24Regular,
  TextNumberFormat24Regular,
} from "@fluentui/react-icons";
import {
  getDocumentStats,
  getBasicDocumentStats,
  formatDocumentStats,
  type DocumentStats as IDocumentStats,
  type GetDocumentStatsOptions,
} from "../../../word-tools";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
    gap: "16px",
    padding: "8px",
  },
  optionsContainer: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    marginBottom: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  optionRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    width: "100%",
  },
  button: {
    flex: 1,
  },
  errorMessage: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase300,
    padding: "12px",
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorPaletteRedBorder1}`,
    width: "100%",
  },
  successMessage: {
    color: tokens.colorPaletteGreenForeground1,
    fontSize: tokens.fontSizeBase300,
    padding: "12px",
    backgroundColor: tokens.colorPaletteGreenBackground1,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorPaletteGreenBorder1}`,
    width: "100%",
  },
  statsCard: {
    width: "100%",
    padding: "16px",
  },
  statsGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "12px",
    width: "100%",
  },
  statItem: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  statLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  statValue: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    marginTop: "12px",
    marginBottom: "8px",
  },
  headingList: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  headingItem: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    fontSize: tokens.fontSizeBase200,
  },
});

/**
 * 文档统计信息组件
 * Document Statistics Component
 */
export const DocumentStats: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [stats, setStats] = useState<IDocumentStats | null>(null);

  // 选项状态 / Options state
  const [includeHeaderFooter, setIncludeHeaderFooter] = useState(false);
  const [includeNotes, setIncludeNotes] = useState(false);
  const [includeHeadingStats, setIncludeHeadingStats] = useState(true);

  /**
   * 获取文档统计信息
   * Get document statistics
   */
  const handleGetStats = async () => {
    setLoading(true);
    setError(null);
    setSuccess(null);
    setStats(null);

    try {
      const options: GetDocumentStatsOptions = {
        includeHeaderFooter,
        includeNotes,
        includeHeadingStats,
      };

      const result = await getDocumentStats(options);
      setStats(result);
      setSuccess("成功获取文档统计信息");
    } catch (err) {
      console.error("获取文档统计信息失败:", err);
      setError(err instanceof Error ? err.message : "获取文档统计信息失败");
    } finally {
      setLoading(false);
    }
  };

  /**
   * 获取基本统计信息
   * Get basic statistics
   */
  const handleGetBasicStats = async () => {
    setLoading(true);
    setError(null);
    setSuccess(null);
    setStats(null);

    try {
      const result = await getBasicDocumentStats();
      // 转换为完整的 IDocumentStats 格式 / Convert to full IDocumentStats format
      const fullStats: IDocumentStats = {
        ...result,
        characterCountNoSpaces: result.characterCount, // 基本统计中使用相同值 / Use same value in basic stats
        sectionCount: 0,
        tableCount: 0,
        imageCount: 0,
        inlinePictureCount: 0,
        contentControlCount: 0,
        listCount: 0,
        footnoteCount: 0,
        endnoteCount: 0,
        headingCounts: {},
        totalHeadingCount: 0,
      };
      setStats(fullStats);
      setSuccess("成功获取基本统计信息");
    } catch (err) {
      console.error("获取基本统计信息失败:", err);
      setError(err instanceof Error ? err.message : "获取基本统计信息失败");
    } finally {
      setLoading(false);
    }
  };

  /**
   * 复制统计信息到剪贴板
   * Copy statistics to clipboard
   */
  const handleCopyStats = async () => {
    if (!stats) return;

    try {
      const formatted = formatDocumentStats(stats);
      await navigator.clipboard.writeText(formatted);
      setSuccess("统计信息已复制到剪贴板");
      setTimeout(() => setSuccess(null), 2000);
    } catch (err) {
      console.error("复制失败:", err);
      setError("复制失败");
    }
  };

  /**
   * 格式化数字
   * Format number
   */
  const formatNumber = (num: number): string => {
    return num.toLocaleString("zh-CN");
  };

  return (
    <div className={styles.container}>
      {/* 选项配置 / Options Configuration */}
      <div className={styles.optionsContainer}>
        <div className={styles.optionRow}>
          <Label>包含页眉页脚</Label>
          <Switch
            checked={includeHeaderFooter}
            onChange={(_, data) => setIncludeHeaderFooter(data.checked)}
          />
        </div>
        <div className={styles.optionRow}>
          <Label>包含脚注尾注</Label>
          <Switch checked={includeNotes} onChange={(_, data) => setIncludeNotes(data.checked)} />
        </div>
        <div className={styles.optionRow}>
          <Label>包含标题统计</Label>
          <Switch
            checked={includeHeadingStats}
            onChange={(_, data) => setIncludeHeadingStats(data.checked)}
          />
        </div>
      </div>

      {/* 操作按钮 / Action Buttons */}
      <div className={styles.buttonGroup}>
        <Button
          className={styles.button}
          appearance="primary"
          icon={<DocumentData24Regular />}
          onClick={handleGetStats}
          disabled={loading}
        >
          获取完整统计
        </Button>
        <Button
          className={styles.button}
          appearance="secondary"
          icon={<TextNumberFormat24Regular />}
          onClick={handleGetBasicStats}
          disabled={loading}
        >
          获取基本统计
        </Button>
      </div>

      {/* 加载状态 / Loading State */}
      {loading && (
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          <Spinner size="small" />
          <span>正在获取统计信息...</span>
        </div>
      )}

      {/* 错误信息 / Error Message */}
      {error && <div className={styles.errorMessage}>{error}</div>}

      {/* 成功信息 / Success Message */}
      {success && <div className={styles.successMessage}>{success}</div>}

      {/* 统计信息显示 / Statistics Display */}
      {stats && (
        <>
          <Card className={styles.statsCard}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <h3 style={{ margin: 0, fontSize: tokens.fontSizeBase400 }}>
                <DocumentText24Regular style={{ verticalAlign: "middle", marginRight: "8px" }} />
                文档统计信息
              </h3>
              <Button
                appearance="subtle"
                icon={<ArrowDownload24Regular />}
                onClick={handleCopyStats}
                size="small"
              >
                复制
              </Button>
            </div>

            <Divider style={{ margin: "12px 0" }} />

            {/* 基本统计 / Basic Statistics */}
            <div className={styles.sectionTitle}>基本统计</div>
            <div className={styles.statsGrid}>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>字符数（含空格）</div>
                <div className={styles.statValue}>{formatNumber(stats.characterCount)}</div>
              </div>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>字符数（不含空格）</div>
                <div className={styles.statValue}>{formatNumber(stats.characterCountNoSpaces)}</div>
              </div>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>单词数</div>
                <div className={styles.statValue}>{formatNumber(stats.wordCount)}</div>
              </div>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>段落数</div>
                <div className={styles.statValue}>{formatNumber(stats.paragraphCount)}</div>
              </div>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>页数（估算）</div>
                <div className={styles.statValue}>
                  {stats.pageCount} <Badge appearance="tint">估算</Badge>
                </div>
              </div>
            </div>

            {/* 结构统计 / Structure Statistics */}
            <div className={styles.sectionTitle}>结构统计</div>
            <div className={styles.statsGrid}>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>节数</div>
                <div className={styles.statValue}>{stats.sectionCount}</div>
              </div>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>表格数</div>
                <div className={styles.statValue}>{stats.tableCount}</div>
              </div>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>图片数</div>
                <div className={styles.statValue}>{stats.imageCount}</div>
              </div>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>内容控件数</div>
                <div className={styles.statValue}>{stats.contentControlCount}</div>
              </div>
              <div className={styles.statItem}>
                <div className={styles.statLabel}>列表数</div>
                <div className={styles.statValue}>{stats.listCount}</div>
              </div>
            </div>

            {/* 注释统计 / Notes Statistics */}
            {(stats.footnoteCount > 0 || stats.endnoteCount > 0) && (
              <>
                <div className={styles.sectionTitle}>注释统计</div>
                <div className={styles.statsGrid}>
                  <div className={styles.statItem}>
                    <div className={styles.statLabel}>脚注数</div>
                    <div className={styles.statValue}>{stats.footnoteCount}</div>
                  </div>
                  <div className={styles.statItem}>
                    <div className={styles.statLabel}>尾注数</div>
                    <div className={styles.statValue}>{stats.endnoteCount}</div>
                  </div>
                </div>
              </>
            )}

            {/* 标题统计 / Heading Statistics */}
            {stats.totalHeadingCount > 0 && (
              <>
                <div className={styles.sectionTitle}>
                  标题统计 <Badge appearance="filled">{stats.totalHeadingCount}</Badge>
                </div>
                <div className={styles.headingList}>
                  {Object.entries(stats.headingCounts)
                    .sort(([a], [b]) => Number(a) - Number(b))
                    .map(([level, count]) => (
                      <div key={level} className={styles.headingItem}>
                        <span>标题 {level}</span>
                        <Badge appearance="tint">{count}</Badge>
                      </div>
                    ))}
                </div>
              </>
            )}
          </Card>
        </>
      )}
    </div>
  );
};
