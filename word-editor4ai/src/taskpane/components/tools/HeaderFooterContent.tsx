/**
 * 文件名: HeaderFooterContent.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/02
 * 最后修改日期: 2025/12/02
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 页眉页脚内容获取工具，用于获取并显示文档所有节的页眉页脚内容
 */

/* global console */

import * as React from "react";
import { useState, useEffect, useRef } from "react";
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
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Input,
} from "@fluentui/react-components";
import {
  DocumentHeader24Regular,
  DocumentFooter24Regular,
  ArrowDownload24Regular,
  Info24Regular,
  DocumentPageBreak24Regular,
} from "@fluentui/react-icons";
import {
  getHeaderFooterContent,
  type DocumentHeaderFooterInfo,
  type GetHeaderFooterContentOptions,
  type HeaderFooterContentItem,
  HeaderFooterType,
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
  inputRow: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
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
  emptyState: {
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
    padding: "32px 16px",
  },
  sectionsContainer: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  sectionCard: {
    width: "100%",
  },
  sectionHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "8px",
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
  },
  headerFooterSection: {
    marginTop: "12px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
  },
  headerFooterTitle: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "8px",
    fontSize: tokens.fontSizeBase300,
  },
  headerFooterItem: {
    marginLeft: "24px",
    marginBottom: "12px",
    fontSize: tokens.fontSizeBase200,
  },
  contentText: {
    marginTop: "8px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    maxHeight: "200px",
    overflowY: "auto",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  statsCard: {
    width: "100%",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  statsGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(2, 1fr)",
    gap: "12px",
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
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  elementsBadge: {
    marginTop: "8px",
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
  },
});

const HeaderFooterContentComponent: React.FC = () => {
  const styles = useStyles();
  const [result, setResult] = useState<DocumentHeaderFooterInfo | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [includeElements, setIncludeElements] = useState(false);
  const [includeMetadata, setIncludeMetadata] = useState(true);
  const [sectionIndexInput, setSectionIndexInput] = useState("");

  // 用于存储定时器 ID，以便在组件卸载时清理 / Store timer IDs for cleanup on unmount
  const timersRef = useRef<NodeJS.Timeout[]>([]);

  // 组件卸载时清理所有定时器 / Cleanup all timers on component unmount
  useEffect(() => {
    return () => {
      timersRef.current.forEach((timer) => clearTimeout(timer));
      timersRef.current = [];
    };
  }, []);

  const handleGetContent = async () => {
    setLoading(true);
    setError(null);
    setSuccessMessage(null);

    try {
      const options: GetHeaderFooterContentOptions = {
        includeElements,
        includeMetadata,
      };

      // 如果输入了节索引，则解析并使用 / Parse and use section index if provided
      if (sectionIndexInput.trim() !== "") {
        const index = parseInt(sectionIndexInput.trim(), 10);
        if (isNaN(index) || index < 0) {
          throw new Error("节索引必须是非负整数");
        }
        options.sectionIndex = index;
      }

      const data = await getHeaderFooterContent(options);
      setResult(data);

      const sectionText = options.sectionIndex !== undefined 
        ? `节 ${options.sectionIndex}` 
        : `${data.totalSections} 个节`;
      setSuccessMessage(`成功获取 ${sectionText} 的页眉页脚内容`);

      const timer = setTimeout(() => setSuccessMessage(null), 3000);
      timersRef.current.push(timer);
    } catch (err) {
      console.error("获取页眉页脚内容失败:", err);
      setError(err instanceof Error ? err.message : "获取页眉页脚内容失败");
    } finally {
      setLoading(false);
    }
  };

  const handleExportJSON = () => {
    if (!result) return;

    try {
      const json = JSON.stringify(result, null, 2);
      const blob = new Blob([json], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `header-footer-content-${Date.now()}.json`;
      a.click();
      URL.revokeObjectURL(url);

      setSuccessMessage("JSON 文件已下载");
      const timer = setTimeout(() => setSuccessMessage(null), 3000);
      timersRef.current.push(timer);
    } catch (err) {
      console.error("导出 JSON 失败:", err);
      setError("导出 JSON 失败");
    }
  };

  const getHeaderFooterTypeName = (type: HeaderFooterType): string => {
    const typeMap: Record<HeaderFooterType, string> = {
      [HeaderFooterType.FirstPage]: "首页",
      [HeaderFooterType.OddPages]: "奇数页",
      [HeaderFooterType.EvenPages]: "偶数页",
    };
    return typeMap[type] || type;
  };

  const renderHeaderFooterItem = (item: HeaderFooterContentItem) => {
    return (
      <div key={item.type} className={styles.headerFooterItem}>
        <Badge
          appearance={item.exists ? "filled" : "outline"}
          color={item.exists ? "success" : "subtle"}
        >
          {getHeaderFooterTypeName(item.type)}
        </Badge>
        {item.exists && item.text && (
          <div className={styles.contentText}>{item.text}</div>
        )}
        {item.exists && includeElements && item.elements && item.elements.length > 0 && (
          <div className={styles.elementsBadge}>
            {item.elements.map((element, idx) => (
              <Badge key={idx} size="small" appearance="outline">
                {element.type}
              </Badge>
            ))}
          </div>
        )}
      </div>
    );
  };

  return (
    <div className={styles.container}>
      {/* 选项区域 */}
      <div className={styles.optionsContainer}>
        <div className={styles.inputRow}>
          <Label>节索引（可选，留空获取所有节）</Label>
          <Input
            type="number"
            value={sectionIndexInput}
            onChange={(_e, data) => setSectionIndexInput(data.value)}
            disabled={loading}
            placeholder="例如: 0, 1, 2..."
          />
        </div>
        <div className={styles.optionRow}>
          <Label>包含详细内容元素</Label>
          <Switch
            checked={includeElements}
            onChange={(_e, data) => setIncludeElements(data.checked)}
            disabled={loading}
          />
        </div>
        <div className={styles.optionRow}>
          <Label>包含元数据统计</Label>
          <Switch
            checked={includeMetadata}
            onChange={(_e, data) => setIncludeMetadata(data.checked)}
            disabled={loading}
          />
        </div>
      </div>

      {/* 操作按钮 */}
      <div className={styles.buttonGroup}>
        <Button
          appearance="primary"
          size="large"
          icon={<DocumentHeader24Regular />}
          onClick={handleGetContent}
          disabled={loading}
          className={styles.button}
        >
          {loading ? <Spinner size="tiny" /> : "获取页眉页脚"}
        </Button>
        <Button
          appearance="secondary"
          size="large"
          icon={<ArrowDownload24Regular />}
          onClick={handleExportJSON}
          disabled={!result || loading}
          className={styles.button}
        >
          导出 JSON
        </Button>
      </div>

      {/* 错误信息 */}
      {error && <div className={styles.errorMessage}>❌ {error}</div>}

      {/* 成功信息 */}
      {successMessage && <div className={styles.successMessage}>✅ {successMessage}</div>}

      {/* 空状态 */}
      {!loading && !error && !result && (
        <div className={styles.emptyState}>
          <Info24Regular style={{ fontSize: "48px", marginBottom: "16px" }} />
          <div>点击"获取页眉页脚"按钮查看文档的页眉页脚内容</div>
        </div>
      )}

      {/* 统计信息 */}
      {result && includeMetadata && result.metadata && (
        <Card className={styles.statsCard}>
          <div className={styles.statsGrid}>
            <div className={styles.statItem}>
              <div className={styles.statLabel}>总节数</div>
              <div className={styles.statValue}>{result.totalSections}</div>
            </div>
            <div className={styles.statItem}>
              <div className={styles.statLabel}>页眉总数</div>
              <div className={styles.statValue}>{result.metadata.totalHeaders}</div>
            </div>
            <div className={styles.statItem}>
              <div className={styles.statLabel}>页脚总数</div>
              <div className={styles.statValue}>{result.metadata.totalFooters}</div>
            </div>
            <div className={styles.statItem}>
              <div className={styles.statLabel}>状态</div>
              <div className={styles.statValue}>
                {result.metadata.hasAnyHeader || result.metadata.hasAnyFooter ? "✓" : "—"}
              </div>
            </div>
          </div>
        </Card>
      )}

      {/* 节信息列表 */}
      {result && result.sections.length > 0 && (
        <div className={styles.sectionsContainer}>
          <Accordion multiple collapsible>
            {result.sections.map((section) => (
              <AccordionItem key={section.sectionIndex} value={`section-${section.sectionIndex}`}>
                <AccordionHeader>
                  <div className={styles.sectionHeader}>
                    <DocumentPageBreak24Regular />
                    <span className={styles.sectionTitle}>节 {section.sectionIndex + 1}</span>
                    {section.differentFirstPage && (
                      <Badge appearance="outline" size="small">首页不同</Badge>
                    )}
                    {section.differentOddAndEven && (
                      <Badge appearance="outline" size="small">奇偶页不同</Badge>
                    )}
                  </div>
                </AccordionHeader>
                <AccordionPanel>
                  {/* 页眉信息 */}
                  <div className={styles.headerFooterSection}>
                    <div className={styles.headerFooterTitle}>
                      <DocumentHeader24Regular />
                      页眉
                    </div>
                    {section.headers.map((header) => renderHeaderFooterItem(header))}
                  </div>

                  <Divider style={{ margin: "12px 0" }} />

                  {/* 页脚信息 */}
                  <div className={styles.headerFooterSection}>
                    <div className={styles.headerFooterTitle}>
                      <DocumentFooter24Regular />
                      页脚
                    </div>
                    {section.footers.map((footer) => renderHeaderFooterItem(footer))}
                  </div>
                </AccordionPanel>
              </AccordionItem>
            ))}
          </Accordion>
        </div>
      )}
    </div>
  );
};

export default HeaderFooterContentComponent;
