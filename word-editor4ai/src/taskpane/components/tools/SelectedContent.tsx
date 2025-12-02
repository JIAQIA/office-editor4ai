/**
 * 文件名: SelectedContent.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 获取选中内容的工具组件
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
} from "@fluentui/react-components";
import {
  getSelectedContent,
  type ContentInfo,
  type AnyContentElement,
  type GetSelectedContentOptions,
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
  contentCard: {
    width: "100%",
  },
  contentHeader: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
    marginBottom: "12px",
  },
  textPreview: {
    marginTop: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground1,
    wordBreak: "break-word",
    lineHeight: "1.5",
    maxHeight: "150px",
    overflowY: "auto",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  elementsList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginTop: "12px",
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
  infoMessage: {
    color: tokens.colorPaletteBlueForeground2,
    fontSize: tokens.fontSizeBase300,
    padding: "16px",
    textAlign: "center",
    backgroundColor: tokens.colorPaletteBlueBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
});

const SelectedContent: React.FC = () => {
  const styles = useStyles();
  const [contentInfo, setContentInfo] = useState<ContentInfo | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  // 选项状态 / Option states
  const [includeText, setIncludeText] = useState(true);
  const [includeImages, setIncludeImages] = useState(true);
  const [includeTables, setIncludeTables] = useState(true);
  const [includeContentControls, setIncludeContentControls] = useState(true);
  const [detailedMetadata, setDetailedMetadata] = useState(false);

  const fetchSelectedContent = async () => {
    setLoading(true);
    setError(null);
    setSuccessMessage(null);

    try {
      const options: GetSelectedContentOptions = {
        includeText,
        includeImages,
        includeTables,
        includeContentControls,
        detailedMetadata,
        maxTextLength: 500, // 限制文本长度 / Limit text length
      };

      const content = await getSelectedContent(options);
      setContentInfo(content);

      if (content.metadata?.isEmpty) {
        setSuccessMessage("当前没有选中任何内容");
      } else {
        setSuccessMessage(
          `成功获取选中内容，包含 ${content.elements.length} 个元素，${content.metadata?.characterCount || 0} 个字符`
        );
      }
    } catch (err) {
      console.error("获取选中内容失败:", err);
      setError(err instanceof Error ? err.message : "获取选中内容失败");
      setContentInfo(null);
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
      <div className={styles.infoMessage}>
        💡 请先在文档中选中要获取的内容，然后点击下方按钮
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
          <Switch checked={detailedMetadata} onChange={(_e, data) => setDetailedMetadata(data.checked)} />
        </div>
      </div>

      <div className={styles.buttonContainer}>
        <Button appearance="primary" size="large" onClick={fetchSelectedContent} disabled={loading}>
          {loading ? <Spinner size="tiny" /> : "获取选中内容"}
        </Button>
      </div>

      {error && <div className={styles.errorMessage}>❌ {error}</div>}
      {successMessage && <div className={styles.successMessage}>✅ {successMessage}</div>}

      {!loading && !error && !contentInfo && (
        <div className={styles.emptyState}>在文档中选中内容后，点击上方按钮获取选中内容信息</div>
      )}

      {contentInfo && !contentInfo.metadata?.isEmpty && (
        <>
          {/* 统计信息 / Statistics */}
          <div className={styles.statsContainer}>
            <div className={styles.statItem}>
              <div className={styles.statValue}>{contentInfo.metadata?.characterCount || 0}</div>
              <div className={styles.statLabel}>字符数</div>
            </div>
            <div className={styles.statItem}>
              <div className={styles.statValue}>{contentInfo.elements.length}</div>
              <div className={styles.statLabel}>元素总数</div>
            </div>
            <div className={styles.statItem}>
              <div className={styles.statValue}>{contentInfo.metadata?.paragraphCount || 0}</div>
              <div className={styles.statLabel}>段落数</div>
            </div>
            <div className={styles.statItem}>
              <div className={styles.statValue}>{contentInfo.metadata?.tableCount || 0}</div>
              <div className={styles.statLabel}>表格数</div>
            </div>
            <div className={styles.statItem}>
              <div className={styles.statValue}>{contentInfo.metadata?.imageCount || 0}</div>
              <div className={styles.statLabel}>图片数</div>
            </div>
          </div>

          {/* 选中文本预览 / Selected text preview */}
          {includeText && contentInfo.text && (
            <Card className={styles.contentCard}>
              <CardHeader
                header={<div className={styles.contentHeader}>📄 选中文本预览</div>}
              />
              <Divider />
              <div className={styles.textPreview}>{contentInfo.text}</div>
            </Card>
          )}

          {/* 元素列表 / Elements list */}
          {contentInfo.elements.length > 0 && (
            <Card className={styles.contentCard}>
              <CardHeader
                header={
                  <div className={styles.contentHeader}>
                    📦 内容元素 ({contentInfo.elements.length})
                  </div>
                }
              />
              <Divider />
              <div className={styles.elementsList}>
                {contentInfo.elements.map((element, elementIndex) => (
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
        </>
      )}

      {contentInfo && contentInfo.metadata?.isEmpty && (
        <div className={styles.emptyState}>
          当前没有选中任何内容
          <br />
          请在文档中选中文本、表格或其他内容后重试
        </div>
      )}
    </div>
  );
};

export default SelectedContent;
