/**
 * 文件名: PageContent.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 获取指定页面内容的工具组件
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
  CardHeader,
  Divider,
  Input,
} from "@fluentui/react-components";
import {
  getPageContent,
  getPageStats,
  type PageInfo,
  type AnyContentElement,
  type GetPageContentOptions,
} from "../../../word-tools";

/**
 * 获取元素类型的友好显示名称
 * Get friendly display name for element type
 */
const getElementTypeDisplay = (type: string): string => {
  const typeMap: Record<string, string> = {
    Paragraph: "段落",
    Table: "表格",
    Image: "图片",
    InlinePicture: "内联图片",
    ContentControl: "内容控件",
    Unknown: "未知",
  };
  return typeMap[type] || type;
};

/**
 * 获取元素类型的图标
 * Get icon for element type
 */
const getElementTypeIcon = (type: string): string => {
  const iconMap: Record<string, string> = {
    Paragraph: "📝",
    Table: "📋",
    Image: "🖼️",
    InlinePicture: "🖼️",
    ContentControl: "🎛️",
    Unknown: "❓",
  };
  return iconMap[type] || "⬜";
};

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
    gap: "16px",
    padding: "8px",
  },
  inputContainer: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginBottom: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  inputRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
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
    justifyContent: "space-between",
    alignItems: "center",
  },
  buttonContainer: {
    width: "100%",
    display: "flex",
    gap: "8px",
    justifyContent: "center",
    marginBottom: "8px",
  },
  statsContainer: {
    width: "100%",
    padding: "16px",
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: tokens.borderRadiusMedium,
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "12px",
  },
  statItem: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  statValue: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightBold,
    color: tokens.colorBrandForeground1,
  },
  statLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
  },
  emptyState: {
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
    padding: "32px 16px",
  },
  pageCard: {
    width: "100%",
  },
  pageHeader: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
    marginBottom: "12px",
  },
  elementsList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginTop: "8px",
  },
  elementCard: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusSmall,
    padding: "10px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  elementHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "6px",
  },
  elementType: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorBrandForeground1,
    display: "flex",
    alignItems: "center",
    gap: "6px",
  },
  typeIcon: {
    fontSize: "14px",
  },
  elementId: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    fontFamily: "monospace",
  },
  elementText: {
    marginTop: "6px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground1,
    wordBreak: "break-word",
    lineHeight: "1.4",
    maxHeight: "100px",
    overflowY: "auto",
  },
  elementMetadata: {
    marginTop: "6px",
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
  },
  metadataItem: {
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "2px 6px",
    borderRadius: tokens.borderRadiusSmall,
  },
  errorMessage: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase300,
    padding: "16px",
    textAlign: "center",
  },
  successMessage: {
    color: tokens.colorPaletteGreenForeground1,
    fontSize: tokens.fontSizeBase300,
    padding: "16px",
    textAlign: "center",
  },
});

const PageContent: React.FC = () => {
  const styles = useStyles();
  const [pageNumber, setPageNumber] = useState<string>("1");
  const [pageInfo, setPageInfo] = useState<PageInfo | null>(null);
  const [stats, setStats] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  // 选项状态 / Option states
  const [includeText, setIncludeText] = useState(true);
  const [includeImages, setIncludeImages] = useState(true);
  const [includeTables, setIncludeTables] = useState(true);
  const [includeContentControls, setIncludeContentControls] = useState(true);
  const [detailedMetadata, setDetailedMetadata] = useState(false);

  const fetchPageContent = async () => {
    setLoading(true);
    setError(null);
    setSuccessMessage(null);
    setStats(null);

    try {
      const pageNum = parseInt(pageNumber, 10);
      if (isNaN(pageNum) || pageNum < 1) {
        setError("请输入有效的页面编号（大于等于1）");
        return;
      }

      const options: GetPageContentOptions = {
        includeText,
        includeImages,
        includeTables,
        includeContentControls,
        detailedMetadata,
        maxTextLength: 500, // 限制文本长度 / Limit text length
      };

      const content = await getPageContent(pageNum, options);
      setPageInfo(content);
      setSuccessMessage(`成功获取第 ${pageNum} 页内容，包含 ${content.elements.length} 个元素`);
    } catch (err) {
      console.error("获取页面内容失败:", err);
      setError(err instanceof Error ? err.message : "获取页面内容失败");
      setPageInfo(null);
    } finally {
      setLoading(false);
    }
  };

  const fetchPageStats = async () => {
    setLoading(true);
    setError(null);
    setSuccessMessage(null);
    setPageInfo(null);

    try {
      const pageNum = parseInt(pageNumber, 10);
      if (isNaN(pageNum) || pageNum < 1) {
        setError("请输入有效的页面编号（大于等于1）");
        return;
      }

      const statistics = await getPageStats(pageNum);
      setStats(statistics);
      setSuccessMessage(`成功获取第 ${pageNum} 页统计信息`);
    } catch (err) {
      console.error("获取页面统计信息失败:", err);
      setError(err instanceof Error ? err.message : "获取页面统计信息失败");
      setStats(null);
    } finally {
      setLoading(false);
    }
  };

  const renderElementMetadata = (element: AnyContentElement) => {
    const metadata: string[] = [];

    if (element.type === "Paragraph") {
      const para = element as any;
      if (para.style) metadata.push(`样式: ${para.style}`);
      if (para.alignment) metadata.push(`对齐: ${para.alignment}`);
      if (para.isListItem) metadata.push("列表项");
    } else if (element.type === "Table") {
      const table = element as any;
      if (table.rowCount) metadata.push(`${table.rowCount} 行`);
      if (table.columnCount) metadata.push(`${table.columnCount} 列`);
    } else if (element.type === "Image" || element.type === "InlinePicture") {
      const img = element as any;
      if (img.width && img.height) metadata.push(`${img.width}×${img.height}`);
      if (img.altText) metadata.push(`描述: ${img.altText}`);
    } else if (element.type === "ContentControl") {
      const ctrl = element as any;
      if (ctrl.title) metadata.push(`标题: ${ctrl.title}`);
      if (ctrl.tag) metadata.push(`标签: ${ctrl.tag}`);
      if (ctrl.controlType) metadata.push(`类型: ${ctrl.controlType}`);
    }

    return metadata;
  };

  return (
    <div className={styles.container}>
      <div className={styles.inputContainer}>
        <Label weight="semibold">页面编号</Label>
        <div className={styles.inputRow}>
          <Input
            type="number"
            value={pageNumber}
            onChange={(_e, data) => setPageNumber(data.value)}
            placeholder="输入页面编号（从1开始）"
            min={1}
            style={{ flex: 1 }}
          />
        </div>
      </div>

      <div className={styles.optionsContainer}>
        <Label weight="semibold">获取选项</Label>
        <div className={styles.optionRow}>
          <Label>包含文本内容</Label>
          <Switch checked={includeText} onChange={(_e, data) => setIncludeText(data.checked)} />
        </div>
        <div className={styles.optionRow}>
          <Label>包含图片信息</Label>
          <Switch checked={includeImages} onChange={(_e, data) => setIncludeImages(data.checked)} />
        </div>
        <div className={styles.optionRow}>
          <Label>包含表格信息</Label>
          <Switch checked={includeTables} onChange={(_e, data) => setIncludeTables(data.checked)} />
        </div>
        <div className={styles.optionRow}>
          <Label>包含内容控件</Label>
          <Switch
            checked={includeContentControls}
            onChange={(_e, data) => setIncludeContentControls(data.checked)}
          />
        </div>
        <div className={styles.optionRow}>
          <Label>详细元数据</Label>
          <Switch
            checked={detailedMetadata}
            onChange={(_e, data) => setDetailedMetadata(data.checked)}
          />
        </div>
      </div>

      <div className={styles.buttonContainer}>
        <Button appearance="primary" size="large" onClick={fetchPageContent} disabled={loading}>
          {loading ? <Spinner size="tiny" /> : "获取页面内容"}
        </Button>
        <Button appearance="secondary" size="large" onClick={fetchPageStats} disabled={loading}>
          {loading ? <Spinner size="tiny" /> : "获取统计信息"}
        </Button>
      </div>

      {error && <div className={styles.errorMessage}>❌ {error}</div>}
      {successMessage && <div className={styles.successMessage}>✅ {successMessage}</div>}

      {stats && (
        <div className={styles.statsContainer}>
          <div className={styles.statItem}>
            <div className={styles.statValue}>{stats.pageIndex + 1}</div>
            <div className={styles.statLabel}>页面编号</div>
          </div>
          <div className={styles.statItem}>
            <div className={styles.statValue}>{stats.elementCount}</div>
            <div className={styles.statLabel}>元素总数</div>
          </div>
          <div className={styles.statItem}>
            <div className={styles.statValue}>{stats.paragraphCount}</div>
            <div className={styles.statLabel}>段落数</div>
          </div>
          <div className={styles.statItem}>
            <div className={styles.statValue}>{stats.tableCount}</div>
            <div className={styles.statLabel}>表格数</div>
          </div>
          <div className={styles.statItem}>
            <div className={styles.statValue}>{stats.imageCount}</div>
            <div className={styles.statLabel}>图片数</div>
          </div>
          <div className={styles.statItem}>
            <div className={styles.statValue}>{stats.contentControlCount}</div>
            <div className={styles.statLabel}>控件数</div>
          </div>
          <div className={styles.statItem}>
            <div className={styles.statValue}>{stats.characterCount}</div>
            <div className={styles.statLabel}>字符数</div>
          </div>
        </div>
      )}

      {!loading && !error && !pageInfo && !stats && (
        <div className={styles.emptyState}>输入页面编号并点击按钮获取页面内容或统计信息</div>
      )}

      {pageInfo && (
        <Card className={styles.pageCard}>
          <CardHeader
            header={
              <div className={styles.pageHeader}>
                📄 页面 {pageInfo.index + 1} ({pageInfo.elements.length} 个元素)
              </div>
            }
          />
          <Divider />
          <div className={styles.elementsList}>
            {pageInfo.elements.map((element, elementIndex) => (
              <div key={element.id} className={styles.elementCard}>
                <div className={styles.elementHeader}>
                  <div className={styles.elementType}>
                    <span className={styles.typeIcon}>{getElementTypeIcon(element.type)}</span>
                    <span>{getElementTypeDisplay(element.type)}</span>
                  </div>
                  <span className={styles.elementId}>#{elementIndex + 1}</span>
                </div>

                {element.text && (
                  <div className={styles.elementText}>
                    {element.text.length > 200 ? `${element.text.substring(0, 200)}...` : element.text}
                  </div>
                )}

                {detailedMetadata && renderElementMetadata(element).length > 0 && (
                  <div className={styles.elementMetadata}>
                    {renderElementMetadata(element).map((meta, idx) => (
                      <span key={idx} className={styles.metadataItem}>
                        {meta}
                      </span>
                    ))}
                  </div>
                )}
              </div>
            ))}
          </div>
        </Card>
      )}
    </div>
  );
};

export default PageContent;
