/**
 * 文件名: TextBoxContent.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 获取文本框内容的工具组件
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
  getTextBoxes,
  type TextBoxInfo,
  type GetTextBoxOptions,
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
    alignItems: "center",
    gap: "8px",
  },
  button: {
    width: "100%",
    marginTop: "8px",
  },
  resultContainer: {
    width: "100%",
    marginTop: "16px",
  },
  resultCard: {
    marginBottom: "12px",
    width: "100%",
  },
  cardContent: {
    padding: "12px",
  },
  textBoxHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "8px",
  },
  textBoxIcon: {
    fontSize: "24px",
  },
  textBoxTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  metadataGrid: {
    display: "grid",
    gridTemplateColumns: "auto 1fr",
    gap: "8px",
    marginBottom: "12px",
  },
  metadataLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
  },
  metadataValue: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    wordBreak: "break-word",
  },
  textContent: {
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    marginBottom: "8px",
  },
  paragraphItem: {
    padding: "8px",
    marginBottom: "8px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
    borderLeft: `3px solid ${tokens.colorBrandBackground}`,
  },
  paragraphText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    marginBottom: "4px",
  },
  paragraphMeta: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
  },
  emptyState: {
    textAlign: "center",
    padding: "24px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
  },
  errorState: {
    textAlign: "center",
    padding: "24px",
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase300,
  },
  jsonOutput: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    fontFamily: "monospace",
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    overflowX: "auto",
    maxHeight: "400px",
    overflowY: "auto",
  },
});

const TextBoxContent: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [textBoxes, setTextBoxes] = useState<TextBoxInfo[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  // 选项状态 / Option states
  const [includeText, setIncludeText] = useState(true);
  const [includeParagraphs, setIncludeParagraphs] = useState(false);
  const [detailedMetadata, setDetailedMetadata] = useState(false);
  const [maxTextLength, setMaxTextLength] = useState<string>("");

  /**
   * 获取文本框内容
   * Get text box content
   */
  const handleGetTextBoxes = async () => {
    setLoading(true);
    setError(null);
    setTextBoxes(null);

    try {
      const options: GetTextBoxOptions = {
        includeText,
        includeParagraphs,
        detailedMetadata,
        maxTextLength: maxTextLength ? parseInt(maxTextLength, 10) : undefined,
      };

      console.log("获取文本框内容，选项:", options);
      const result = await getTextBoxes(options);
      console.log("获取到的文本框:", result);
      setTextBoxes(result);
    } catch (err) {
      console.error("获取文本框内容失败:", err);
      setError(err instanceof Error ? err.message : "未知错误");
    } finally {
      setLoading(false);
    }
  };

  /**
   * 渲染文本框卡片
   * Render text box card
   */
  const renderTextBoxCard = (textBox: TextBoxInfo, index: number) => {
    return (
      <Card key={textBox.id} className={styles.resultCard}>
        <CardHeader
          header={
            <div className={styles.textBoxHeader}>
              <span className={styles.textBoxIcon}>📦</span>
              <span className={styles.textBoxTitle}>
                文本框 {index + 1}
                {textBox.name && `: ${textBox.name}`}
              </span>
            </div>
          }
        />
        <div className={styles.cardContent}>
          {/* 元数据信息 / Metadata information */}
          {detailedMetadata && (
            <>
              <div className={styles.metadataGrid}>
                <span className={styles.metadataLabel}>ID:</span>
                <span className={styles.metadataValue}>{textBox.id}</span>

                {textBox.width !== undefined && (
                  <>
                    <span className={styles.metadataLabel}>宽度:</span>
                    <span className={styles.metadataValue}>{textBox.width.toFixed(2)} pt</span>
                  </>
                )}

                {textBox.height !== undefined && (
                  <>
                    <span className={styles.metadataLabel}>高度:</span>
                    <span className={styles.metadataValue}>{textBox.height.toFixed(2)} pt</span>
                  </>
                )}

                {textBox.left !== undefined && (
                  <>
                    <span className={styles.metadataLabel}>左边距:</span>
                    <span className={styles.metadataValue}>{textBox.left.toFixed(2)} pt</span>
                  </>
                )}

                {textBox.top !== undefined && (
                  <>
                    <span className={styles.metadataLabel}>上边距:</span>
                    <span className={styles.metadataValue}>{textBox.top.toFixed(2)} pt</span>
                  </>
                )}

                {textBox.rotation !== undefined && (
                  <>
                    <span className={styles.metadataLabel}>旋转角度:</span>
                    <span className={styles.metadataValue}>{textBox.rotation}°</span>
                  </>
                )}

                {textBox.visible !== undefined && (
                  <>
                    <span className={styles.metadataLabel}>可见性:</span>
                    <span className={styles.metadataValue}>{textBox.visible ? "可见" : "隐藏"}</span>
                  </>
                )}

                {textBox.lockAspectRatio !== undefined && (
                  <>
                    <span className={styles.metadataLabel}>锁定纵横比:</span>
                    <span className={styles.metadataValue}>
                      {textBox.lockAspectRatio ? "是" : "否"}
                    </span>
                  </>
                )}
              </div>
              <Divider />
            </>
          )}

          {/* 文本内容 / Text content */}
          {includeText && textBox.text && (
            <>
              <Label weight="semibold">文本内容:</Label>
              <div className={styles.textContent}>{textBox.text}</div>
            </>
          )}

          {/* 段落详情 / Paragraph details */}
          {includeParagraphs && textBox.paragraphs && textBox.paragraphs.length > 0 && (
            <>
              <Label weight="semibold">段落详情 ({textBox.paragraphs.length} 个段落):</Label>
              {textBox.paragraphs.map((para) => (
                <div key={para.id} className={styles.paragraphItem}>
                  <div className={styles.paragraphText}>
                    {getElementTypeIcon(para.type)} {para.text}
                  </div>
                  {detailedMetadata && (
                    <div className={styles.paragraphMeta}>
                      {para.style && `样式: ${para.style} | `}
                      {para.alignment && `对齐: ${para.alignment} | `}
                      {para.isListItem !== undefined && `列表项: ${para.isListItem ? "是" : "否"}`}
                    </div>
                  )}
                </div>
              ))}
            </>
          )}
        </div>
      </Card>
    );
  };

  return (
    <div className={styles.container}>
      {/* 选项配置 / Options configuration */}
      <div className={styles.optionsContainer}>
        <Label weight="semibold">获取选项</Label>

        <div className={styles.optionRow}>
          <Switch
            checked={includeText}
            onChange={(_, data) => setIncludeText(data.checked)}
            label="包含文本内容"
          />
        </div>

        <div className={styles.optionRow}>
          <Switch
            checked={includeParagraphs}
            onChange={(_, data) => setIncludeParagraphs(data.checked)}
            label="包含段落详情"
          />
        </div>

        <div className={styles.optionRow}>
          <Switch
            checked={detailedMetadata}
            onChange={(_, data) => setDetailedMetadata(data.checked)}
            label="详细元数据"
          />
        </div>

        <div className={styles.optionRow}>
          <Label>最大文本长度 (可选):</Label>
          <Input
            type="number"
            value={maxTextLength}
            onChange={(_, data) => setMaxTextLength(data.value)}
            placeholder="不限制"
          />
        </div>
      </div>

      {/* 获取按钮 / Get button */}
      <Button
        appearance="primary"
        className={styles.button}
        onClick={handleGetTextBoxes}
        disabled={loading}
      >
        {loading ? <Spinner size="tiny" /> : "获取文本框内容"}
      </Button>

      {/* 结果展示 / Result display */}
      {error && <div className={styles.errorState}>错误: {error}</div>}

      {!loading && !error && textBoxes !== null && (
        <div className={styles.resultContainer}>
          {textBoxes.length === 0 ? (
            <div className={styles.emptyState}>未找到文本框</div>
          ) : (
            <>
              <Label weight="semibold">找到 {textBoxes.length} 个文本框:</Label>
              {textBoxes.map((textBox, index) => renderTextBoxCard(textBox, index))}

              {/* JSON 输出 / JSON output */}
              <Card className={styles.resultCard}>
                <CardHeader header={<Label weight="semibold">JSON 输出</Label>} />
                <div className={styles.cardContent}>
                  <div className={styles.jsonOutput}>{JSON.stringify(textBoxes, null, 2)}</div>
                </div>
              </Card>
            </>
          )}
        </div>
      )}
    </div>
  );
};

export default TextBoxContent;
