/**
 * 文件名: RangeContent.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 获取指定范围内容的工具组件
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
  Dropdown,
  Option,
} from "@fluentui/react-components";
import {
  getRangeContent,
  type ContentInfo,
  type AnyContentElement,
  type GetRangeContentOptions,
  type RangeLocator,
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
  locatorContainer: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    marginBottom: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  locatorRow: {
    display: "flex",
    flexDirection: "column",
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
    marginBottom: "8px",
  },
  statsGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "12px",
    marginTop: "12px",
  },
  statItem: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  statLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  statValue: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
  },
  elementsContainer: {
    width: "100%",
    maxHeight: "400px",
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  elementCard: {
    width: "100%",
  },
  elementHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  elementIcon: {
    fontSize: "20px",
  },
  elementContent: {
    padding: "12px",
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    maxHeight: "200px",
    overflowY: "auto",
  },
  emptyState: {
    width: "100%",
    padding: "32px",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
  },
  errorState: {
    width: "100%",
    padding: "16px",
    backgroundColor: tokens.colorPaletteRedBackground2,
    borderRadius: tokens.borderRadiusMedium,
    color: tokens.colorPaletteRedForeground1,
  },
});

/**
 * 获取指定范围内容组件
 * Range Content Component
 */
export const RangeContent: React.FC = () => {
  const styles = useStyles();

  // 状态管理 / State management
  const [loading, setLoading] = useState(false);
  const [contentInfo, setContentInfo] = useState<ContentInfo | null>(null);
  const [error, setError] = useState<string | null>(null);

  // 定位器类型 / Locator type
  const [locatorType, setLocatorType] = useState<string>("bookmark");

  // 书签定位器参数 / Bookmark locator parameters
  const [bookmarkName, setBookmarkName] = useState<string>("");

  // 标题定位器参数 / Heading locator parameters
  const [headingText, setHeadingText] = useState<string>("");
  const [headingLevel, setHeadingLevel] = useState<string>("");
  const [headingIndex, setHeadingIndex] = useState<string>("0");

  // 段落定位器参数 / Paragraph locator parameters
  const [paragraphStartIndex, setParagraphStartIndex] = useState<string>("0");
  const [paragraphEndIndex, setParagraphEndIndex] = useState<string>("");

  // 节定位器参数 / Section locator parameters
  const [sectionIndex, setSectionIndex] = useState<string>("0");

  // 内容控件定位器参数 / Content control locator parameters
  const [controlTitle, setControlTitle] = useState<string>("");
  const [controlTag, setControlTag] = useState<string>("");
  const [controlIndex, setControlIndex] = useState<string>("0");

  // 选项 / Options
  const [includeText, setIncludeText] = useState(true);
  const [includeImages, setIncludeImages] = useState(true);
  const [includeTables, setIncludeTables] = useState(true);
  const [includeContentControls, setIncludeContentControls] = useState(true);
  const [detailedMetadata, setDetailedMetadata] = useState(false);

  /**
   * 构建范围定位器
   * Build range locator
   */
  const buildLocator = (): RangeLocator | null => {
    switch (locatorType) {
      case "bookmark":
        if (!bookmarkName.trim()) {
          setError("请输入书签名称");
          return null;
        }
        return { type: "bookmark", name: bookmarkName.trim() };

      case "heading":
        return {
          type: "heading",
          text: headingText.trim() || undefined,
          level: headingLevel ? parseInt(headingLevel) : undefined,
          index: parseInt(headingIndex),
        };

      case "paragraph":
        const startIdx = parseInt(paragraphStartIndex);
        return {
          type: "paragraph",
          startIndex: startIdx,
          endIndex: paragraphEndIndex ? parseInt(paragraphEndIndex) : undefined,
        };

      case "section":
        return {
          type: "section",
          index: parseInt(sectionIndex),
        };

      case "contentControl":
        if (!controlTitle.trim() && !controlTag.trim()) {
          setError("请至少输入控件标题或标签");
          return null;
        }
        return {
          type: "contentControl",
          title: controlTitle.trim() || undefined,
          tag: controlTag.trim() || undefined,
          index: parseInt(controlIndex),
        };

      default:
        setError(`不支持的定位器类型: ${locatorType}`);
        return null;
    }
  };

  /**
   * 获取范围内容
   * Get range content
   */
  const handleGetRangeContent = async () => {
    setLoading(true);
    setError(null);
    setContentInfo(null);

    try {
      const locator = buildLocator();
      if (!locator) {
        setLoading(false);
        return;
      }

      const options: GetRangeContentOptions = {
        includeText,
        includeImages,
        includeTables,
        includeContentControls,
        detailedMetadata,
      };

      const result = await getRangeContent(locator, options);
      setContentInfo(result);
      console.log("范围内容 / Range content:", result);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : String(err);
      setError(errorMessage);
      console.error("获取范围内容失败 / Failed to get range content:", err);
    } finally {
      setLoading(false);
    }
  };

  /**
   * 清空结果
   * Clear results
   */
  const handleClear = () => {
    setContentInfo(null);
    setError(null);
  };

  /**
   * 渲染元素详情
   * Render element details
   */
  const renderElementDetails = (element: AnyContentElement) => {
    const details: string[] = [];

    if (element.text) {
      details.push(`文本: ${element.text.substring(0, 100)}${element.text.length > 100 ? "..." : ""}`);
    }

    if (element.type === "Paragraph") {
      const para = element as any;
      if (para.style) details.push(`样式: ${para.style}`);
      if (para.alignment) details.push(`对齐: ${para.alignment}`);
      if (para.isListItem) details.push(`列表项: 是`);
    }

    if (element.type === "Table") {
      const table = element as any;
      details.push(`行数: ${table.rowCount || 0}`);
      details.push(`列数: ${table.columnCount || 0}`);
    }

    if (element.type === "Image" || element.type === "InlinePicture") {
      const img = element as any;
      if (img.width) details.push(`宽度: ${img.width.toFixed(1)} pt`);
      if (img.height) details.push(`高度: ${img.height.toFixed(1)} pt`);
      if (img.altText) details.push(`替代文本: ${img.altText}`);
    }

    if (element.type === "ContentControl") {
      const ctrl = element as any;
      if (ctrl.title) details.push(`标题: ${ctrl.title}`);
      if (ctrl.tag) details.push(`标签: ${ctrl.tag}`);
      if (ctrl.controlType) details.push(`类型: ${ctrl.controlType}`);
    }

    return details.join("\n");
  };

  /**
   * 渲染定位器输入区域
   * Render locator input area
   */
  const renderLocatorInputs = () => {
    switch (locatorType) {
      case "bookmark":
        return (
          <div className={styles.locatorRow}>
            <Label>书签名称</Label>
            <Input value={bookmarkName} onChange={(_, data) => setBookmarkName(data.value)} placeholder="输入书签名称" />
          </div>
        );

      case "heading":
        return (
          <>
            <div className={styles.locatorRow}>
              <Label>标题文本（可选）</Label>
              <Input value={headingText} onChange={(_, data) => setHeadingText(data.value)} placeholder="输入标题文本" />
            </div>
            <div className={styles.locatorRow}>
              <Label>标题级别（可选，1-9）</Label>
              <Input
                type="number"
                value={headingLevel}
                onChange={(_, data) => setHeadingLevel(data.value)}
                placeholder="输入标题级别"
              />
            </div>
            <div className={styles.locatorRow}>
              <Label>标题索引（从0开始）</Label>
              <Input
                type="number"
                value={headingIndex}
                onChange={(_, data) => setHeadingIndex(data.value)}
                placeholder="0"
              />
            </div>
          </>
        );

      case "paragraph":
        return (
          <>
            <div className={styles.locatorRow}>
              <Label>起始段落索引（从0开始）</Label>
              <Input
                type="number"
                value={paragraphStartIndex}
                onChange={(_, data) => setParagraphStartIndex(data.value)}
                placeholder="0"
              />
            </div>
            <div className={styles.locatorRow}>
              <Label>结束段落索引（可选）</Label>
              <Input
                type="number"
                value={paragraphEndIndex}
                onChange={(_, data) => setParagraphEndIndex(data.value)}
                placeholder="留空则只获取单个段落"
              />
            </div>
          </>
        );

      case "section":
        return (
          <div className={styles.locatorRow}>
            <Label>节索引（从0开始）</Label>
            <Input
              type="number"
              value={sectionIndex}
              onChange={(_, data) => setSectionIndex(data.value)}
              placeholder="0"
            />
          </div>
        );

      case "contentControl":
        return (
          <>
            <div className={styles.locatorRow}>
              <Label>控件标题（可选）</Label>
              <Input value={controlTitle} onChange={(_, data) => setControlTitle(data.value)} placeholder="输入控件标题" />
            </div>
            <div className={styles.locatorRow}>
              <Label>控件标签（可选）</Label>
              <Input value={controlTag} onChange={(_, data) => setControlTag(data.value)} placeholder="输入控件标签" />
            </div>
            <div className={styles.locatorRow}>
              <Label>控件索引（从0开始）</Label>
              <Input
                type="number"
                value={controlIndex}
                onChange={(_, data) => setControlIndex(data.value)}
                placeholder="0"
              />
            </div>
          </>
        );

      default:
        return null;
    }
  };

  return (
    <div className={styles.container}>
      {/* 定位器配置 / Locator Configuration */}
      <div className={styles.locatorContainer}>
        <Label weight="semibold">范围定位方式</Label>
        <Dropdown
          value={
            locatorType === "bookmark"
              ? "书签"
              : locatorType === "heading"
                ? "标题"
                : locatorType === "paragraph"
                  ? "段落索引"
                  : locatorType === "section"
                    ? "节"
                    : "内容控件"
          }
          onOptionSelect={(_, data) => {
            const typeMap: Record<string, string> = {
              书签: "bookmark",
              标题: "heading",
              段落索引: "paragraph",
              节: "section",
              内容控件: "contentControl",
            };
            setLocatorType(typeMap[data.optionValue as string] || "bookmark");
            setError(null);
          }}
        >
          <Option value="书签">书签</Option>
          <Option value="标题">标题</Option>
          <Option value="段落索引">段落索引</Option>
          <Option value="节">节</Option>
          <Option value="内容控件">内容控件</Option>
        </Dropdown>

        <Divider />

        {renderLocatorInputs()}
      </div>

      {/* 选项配置 / Options Configuration */}
      <div className={styles.optionsContainer}>
        <Label weight="semibold">获取选项</Label>
        <div className={styles.optionRow}>
          <Label>包含文本内容</Label>
          <Switch checked={includeText} onChange={(_, data) => setIncludeText(data.checked)} />
        </div>
        <div className={styles.optionRow}>
          <Label>包含图片信息</Label>
          <Switch checked={includeImages} onChange={(_, data) => setIncludeImages(data.checked)} />
        </div>
        <div className={styles.optionRow}>
          <Label>包含表格信息</Label>
          <Switch checked={includeTables} onChange={(_, data) => setIncludeTables(data.checked)} />
        </div>
        <div className={styles.optionRow}>
          <Label>包含内容控件</Label>
          <Switch checked={includeContentControls} onChange={(_, data) => setIncludeContentControls(data.checked)} />
        </div>
        <div className={styles.optionRow}>
          <Label>详细元数据</Label>
          <Switch checked={detailedMetadata} onChange={(_, data) => setDetailedMetadata(data.checked)} />
        </div>
      </div>

      {/* 操作按钮 / Action Buttons */}
      <div className={styles.buttonContainer}>
        <Button appearance="primary" onClick={handleGetRangeContent} disabled={loading}>
          {loading ? <Spinner size="tiny" /> : "获取范围内容"}
        </Button>
        <Button appearance="secondary" onClick={handleClear} disabled={loading || !contentInfo}>
          清空
        </Button>
      </div>

      {/* 错误信息 / Error Message */}
      {error && <div className={styles.errorState}>❌ {error}</div>}

      {/* 统计信息 / Statistics */}
      {contentInfo?.metadata && (
        <div className={styles.statsContainer}>
          <Label weight="semibold">范围统计信息</Label>
          <div className={styles.statsGrid}>
            <div className={styles.statItem}>
              <span className={styles.statLabel}>定位方式</span>
              <span className={styles.statValue}>{contentInfo.metadata.locatorType}</span>
            </div>
            <div className={styles.statItem}>
              <span className={styles.statLabel}>字符数</span>
              <span className={styles.statValue}>{contentInfo.metadata.characterCount}</span>
            </div>
            <div className={styles.statItem}>
              <span className={styles.statLabel}>段落数</span>
              <span className={styles.statValue}>{contentInfo.metadata.paragraphCount}</span>
            </div>
            <div className={styles.statItem}>
              <span className={styles.statLabel}>表格数</span>
              <span className={styles.statValue}>{contentInfo.metadata.tableCount}</span>
            </div>
            <div className={styles.statItem}>
              <span className={styles.statLabel}>图片数</span>
              <span className={styles.statValue}>{contentInfo.metadata.imageCount}</span>
            </div>
            <div className={styles.statItem}>
              <span className={styles.statLabel}>元素总数</span>
              <span className={styles.statValue}>{contentInfo.elements.length}</span>
            </div>
          </div>
        </div>
      )}

      {/* 元素列表 / Elements List */}
      {contentInfo && contentInfo.elements.length > 0 && (
        <div className={styles.elementsContainer}>
          <Label weight="semibold">范围内容元素 ({contentInfo.elements.length})</Label>
          {contentInfo.elements.map((element, index) => (
            <Card key={element.id || index} className={styles.elementCard}>
              <CardHeader
                header={
                  <div className={styles.elementHeader}>
                    <span className={styles.elementIcon}>{getElementTypeIcon(element.type)}</span>
                    <span>
                      {getElementTypeDisplay(element.type)} #{index + 1}
                    </span>
                  </div>
                }
              />
              <div className={styles.elementContent}>{renderElementDetails(element)}</div>
            </Card>
          ))}
        </div>
      )}

      {/* 空状态 / Empty State */}
      {contentInfo && contentInfo.elements.length === 0 && (
        <div className={styles.emptyState}>指定范围内没有内容元素</div>
      )}
    </div>
  );
};

export default RangeContent;
