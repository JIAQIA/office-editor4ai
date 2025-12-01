/**
 * 文件名: DocumentSections.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 文档节信息获取工具，用于获取并显示文档的分节符、页眉页脚配置等信息
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
} from "@fluentui/react-components";
import {
  DocumentPageBreak24Regular,
  ArrowDownload24Regular,
  Info24Regular,
  DocumentHeader24Regular,
  DocumentFooter24Regular,
} from "@fluentui/react-icons";
import {
  getDocumentSections,
  type SectionInfo,
  type GetDocumentSectionsOptions,
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
  infoGrid: {
    display: "grid",
    gridTemplateColumns: "auto 1fr",
    gap: "8px",
    fontSize: tokens.fontSizeBase200,
  },
  infoLabel: {
    color: tokens.colorNeutralForeground3,
    fontWeight: tokens.fontWeightSemibold,
  },
  infoValue: {
    color: tokens.colorNeutralForeground1,
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
    marginBottom: "8px",
    fontSize: tokens.fontSizeBase200,
  },
  statsCard: {
    width: "100%",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    textAlign: "center",
  },
});

const DocumentSectionsComponent: React.FC = () => {
  const styles = useStyles();
  const [sections, setSections] = useState<SectionInfo[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [includeContent, setIncludeContent] = useState(false);
  const [includePageSetup, setIncludePageSetup] = useState(true);

  // 用于存储定时器 ID，以便在组件卸载时清理 / Store timer IDs for cleanup on unmount
  const timersRef = useRef<NodeJS.Timeout[]>([]);

  // 组件卸载时清理所有定时器 / Cleanup all timers on component unmount
  useEffect(() => {
    return () => {
      timersRef.current.forEach((timer) => clearTimeout(timer));
      timersRef.current = [];
    };
  }, []);

  const handleGetSections = async () => {
    setLoading(true);
    setError(null);
    setSuccessMessage(null);

    try {
      const options: GetDocumentSectionsOptions = {
        includeContent,
        includePageSetup,
      };

      const result = await getDocumentSections(options);
      setSections(result);
      setSuccessMessage(`成功获取 ${result.length} 个文档节信息`);

      const timer = setTimeout(() => setSuccessMessage(null), 3000);
      timersRef.current.push(timer);
    } catch (err) {
      console.error("获取文档节信息失败:", err);
      setError(err instanceof Error ? err.message : "获取文档节信息失败");
    } finally {
      setLoading(false);
    }
  };

  const handleExportJSON = () => {
    if (sections.length === 0) return;

    try {
      const json = JSON.stringify(sections, null, 2);
      const blob = new Blob([json], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `document-sections-${Date.now()}.json`;
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

  const getSectionTypeName = (type: string): string => {
    const typeMap: Record<string, string> = {
      continuous: "连续",
      nextPage: "下一页",
      oddPage: "奇数页",
      evenPage: "偶数页",
      nextColumn: "下一栏",
    };
    return typeMap[type] || type;
  };

  const getHeaderFooterTypeName = (type: HeaderFooterType): string => {
    const typeMap: Record<HeaderFooterType, string> = {
      [HeaderFooterType.FirstPage]: "首页",
      [HeaderFooterType.OddPages]: "奇数页",
      [HeaderFooterType.EvenPages]: "偶数页",
    };
    return typeMap[type] || type;
  };

  return (
    <div className={styles.container}>
      {/* 选项区域 */}
      <div className={styles.optionsContainer}>
        <div className={styles.optionRow}>
          <Label>包含页眉页脚内容</Label>
          <Switch
            checked={includeContent}
            onChange={(_e, data) => setIncludeContent(data.checked)}
            disabled={loading}
          />
        </div>
        <div className={styles.optionRow}>
          <Label>包含页面设置详情</Label>
          <Switch
            checked={includePageSetup}
            onChange={(_e, data) => setIncludePageSetup(data.checked)}
            disabled={loading}
          />
        </div>
      </div>

      {/* 操作按钮 */}
      <div className={styles.buttonGroup}>
        <Button
          appearance="primary"
          size="large"
          icon={<DocumentPageBreak24Regular />}
          onClick={handleGetSections}
          disabled={loading}
          className={styles.button}
        >
          {loading ? <Spinner size="tiny" /> : "获取节信息"}
        </Button>
        <Button
          appearance="secondary"
          size="large"
          icon={<ArrowDownload24Regular />}
          onClick={handleExportJSON}
          disabled={sections.length === 0 || loading}
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
      {!loading && !error && sections.length === 0 && (
        <div className={styles.emptyState}>
          <Info24Regular style={{ fontSize: "48px", marginBottom: "16px" }} />
          <div>点击"获取节信息"按钮查看文档的分节信息</div>
        </div>
      )}

      {/* 统计信息 */}
      {sections.length > 0 && (
        <Card className={styles.statsCard}>
          <div>
            文档共有 <strong>{sections.length}</strong> 个节
          </div>
        </Card>
      )}

      {/* 节信息列表 */}
      {sections.length > 0 && (
        <div className={styles.sectionsContainer}>
          <Accordion multiple collapsible>
            {sections.map((section) => (
              <AccordionItem key={section.index} value={`section-${section.index}`}>
                <AccordionHeader>
                  <div className={styles.sectionHeader}>
                    <DocumentPageBreak24Regular />
                    <span className={styles.sectionTitle}>节 {section.index + 1}</span>
                    <Badge appearance="outline">{getSectionTypeName(section.sectionType)}</Badge>
                  </div>
                </AccordionHeader>
                <AccordionPanel>
                  {/* 基本信息 */}
                  <div className={styles.infoGrid}>
                    <span className={styles.infoLabel}>分节类型:</span>
                    <span className={styles.infoValue}>
                      {getSectionTypeName(section.sectionType)}
                    </span>

                    <span className={styles.infoLabel}>首页不同:</span>
                    <span className={styles.infoValue}>
                      {section.differentFirstPage ? "是" : "否"}
                    </span>

                    <span className={styles.infoLabel}>奇偶页不同:</span>
                    <span className={styles.infoValue}>
                      {section.differentOddAndEven ? "是" : "否"}
                    </span>

                    <span className={styles.infoLabel}>列数:</span>
                    <span className={styles.infoValue}>{section.columnCount}</span>

                    {section.columnSpacing !== undefined && (
                      <>
                        <span className={styles.infoLabel}>列间距:</span>
                        <span className={styles.infoValue}>
                          {section.columnSpacing.toFixed(2)} 磅
                        </span>
                      </>
                    )}
                  </div>

                  {/* 页面设置 */}
                  {includePageSetup && (
                    <>
                      <Divider style={{ margin: "12px 0" }} />
                      <div className={styles.infoGrid}>
                        <span className={styles.infoLabel}>页面方向:</span>
                        <span className={styles.infoValue}>
                          {section.pageSetup.orientation === "portrait" ? "纵向" : "横向"}
                        </span>

                        <span className={styles.infoLabel}>页面尺寸:</span>
                        <span className={styles.infoValue}>
                          {section.pageSetup.pageWidth.toFixed(2)} ×{" "}
                          {section.pageSetup.pageHeight.toFixed(2)} 磅
                        </span>

                        <span className={styles.infoLabel}>边距:</span>
                        <span className={styles.infoValue}>
                          上 {section.pageSetup.topMargin.toFixed(2)} / 下{" "}
                          {section.pageSetup.bottomMargin.toFixed(2)} / 左{" "}
                          {section.pageSetup.leftMargin.toFixed(2)} / 右{" "}
                          {section.pageSetup.rightMargin.toFixed(2)} 磅
                        </span>
                      </div>
                    </>
                  )}

                  {/* 页眉信息 */}
                  <div className={styles.headerFooterSection}>
                    <div className={styles.headerFooterTitle}>
                      <DocumentHeader24Regular />
                      页眉
                    </div>
                    {section.headers.map((header, idx) => (
                      <div key={idx} className={styles.headerFooterItem}>
                        <Badge
                          appearance={header.exists ? "filled" : "outline"}
                          color={header.exists ? "success" : "subtle"}
                        >
                          {getHeaderFooterTypeName(header.type)}
                        </Badge>
                        {header.exists && includeContent && header.text && (
                          <div
                            style={{
                              marginTop: "4px",
                              fontSize: tokens.fontSizeBase100,
                              color: tokens.colorNeutralForeground3,
                            }}
                          >
                            {header.text}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>

                  {/* 页脚信息 */}
                  <div className={styles.headerFooterSection}>
                    <div className={styles.headerFooterTitle}>
                      <DocumentFooter24Regular />
                      页脚
                    </div>
                    {section.footers.map((footer, idx) => (
                      <div key={idx} className={styles.headerFooterItem}>
                        <Badge
                          appearance={footer.exists ? "filled" : "outline"}
                          color={footer.exists ? "success" : "subtle"}
                        >
                          {getHeaderFooterTypeName(footer.type)}
                        </Badge>
                        {footer.exists && includeContent && footer.text && (
                          <div
                            style={{
                              marginTop: "4px",
                              fontSize: tokens.fontSizeBase100,
                              color: tokens.colorNeutralForeground3,
                            }}
                          >
                            {footer.text}
                          </div>
                        )}
                      </div>
                    ))}
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

export default DocumentSectionsComponent;
