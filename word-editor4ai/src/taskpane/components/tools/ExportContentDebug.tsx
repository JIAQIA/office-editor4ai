/**
 * 文件名: ExportContentDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: exportContent 工具的调试组件
 */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  Label,
  makeStyles,
  Spinner,
  tokens,
  Select,
  Card,
  CardHeader,
  Divider,
  Text,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Textarea,
} from "@fluentui/react-components";
import {
  exportContent,
  type ExportContentOptions,
  type ExportScope,
  type ExportFormat,
} from "../../../word-tools";
import { DocumentArrowDown24Regular, Dismiss24Regular, Copy24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  formRow: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    marginTop: "12px",
  },
  resultCard: {
    width: "100%",
  },
  resultMessage: {
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
  },
  success: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground1,
  },
  error: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground1,
  },
  contentPreview: {
    marginTop: "12px",
    maxHeight: "400px",
    overflowY: "auto",
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    whiteSpace: "pre-wrap",
    wordBreak: "break-all",
  },
  infoBox: {
    marginTop: "8px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    fontSize: tokens.fontSizeBase200,
  },
  warningBox: {
    marginBottom: "12px",
  },
});

export const ExportContentDebug: React.FC = () => {
  const styles = useStyles();

  const [loading, setLoading] = useState(false);
  const [scope, setScope] = useState<ExportScope>("selection");
  const [format, setFormat] = useState<ExportFormat>("ooxml");

  const [result, setResult] = useState<{
    content: string;
    format: ExportFormat;
    scope: ExportScope;
    timestamp: number;
    size: number;
    mimeType: string;
  } | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleExport = async () => {
    setLoading(true);
    setError(null);
    setResult(null);

    try {
      const options: ExportContentOptions = {
        scope,
        format,
      };

      const exportResult = await exportContent(options);
      setResult(exportResult);
    } catch (err) {
      console.error("导出失败:", err);
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  };

  const handleCopyContent = async () => {
    if (!result) return;

    try {
      await navigator.clipboard.writeText(result.content);
      alert("内容已复制到剪贴板");
    } catch (error) {
      console.error("复制失败:", error);
      alert("复制失败，请手动复制");
    }
  };

  const handleDownload = () => {
    if (!result) return;

    const blob = new Blob([result.content], { type: result.mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const extension = result.format === "ooxml" ? "xml" : result.format;
    a.download = `export-${result.scope}-${result.timestamp}.${extension}`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleClear = () => {
    setResult(null);
    setError(null);
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) {
      return `${bytes} B`;
    } else if (bytes < 1024 * 1024) {
      return `${(bytes / 1024).toFixed(2)} KB`;
    } else {
      return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
    }
  };

  return (
    <div className={styles.container}>
      {/* 提示信息 */}
      <MessageBar intent="info" className={styles.warningBox}>
        <MessageBarBody>
          <MessageBarTitle>功能说明</MessageBarTitle>
          支持导出 OOXML 和 HTML 格式。OOXML 保留完整格式，HTML 适合网页显示。
          <br />
          PDF 格式暂不可用（Word API 限制）。
        </MessageBarBody>
      </MessageBar>

      {/* 导出选项 */}
      <Card className={styles.section}>
        <CardHeader header={<Text weight="semibold">导出选项</Text>} />
        <Divider />

        <div className={styles.formRow}>
          <Label htmlFor="scope">导出范围</Label>
          <Select
            id="scope"
            value={scope}
            onChange={(_, data) => setScope(data.value as ExportScope)}
          >
            <option value="selection">当前选中区域</option>
            <option value="document">整个文档</option>
            <option value="visible">当前可见区域</option>
          </Select>
        </div>

        <div className={styles.formRow}>
          <Label htmlFor="format">导出格式</Label>
          <Select
            id="format"
            value={format}
            onChange={(_, data) => setFormat(data.value as ExportFormat)}
          >
            <option value="ooxml">OOXML（保留完整格式）</option>
            <option value="html">HTML（适合网页显示）</option>
            <option value="pdf" disabled>
              PDF（暂不可用）
            </option>
          </Select>
        </div>

        <div className={styles.buttonGroup}>
          <Button
            appearance="primary"
            icon={<DocumentArrowDown24Regular />}
            onClick={handleExport}
            disabled={loading}
          >
            {loading ? <Spinner size="tiny" /> : "导出"}
          </Button>
          {(result || error) && (
            <Button appearance="subtle" icon={<Dismiss24Regular />} onClick={handleClear}>
              清除结果
            </Button>
          )}
        </div>
      </Card>

      {/* 错误信息 */}
      {error && (
        <Card className={styles.resultCard}>
          <div className={`${styles.resultMessage} ${styles.error}`}>
            <strong>导出失败</strong>
            <br />
            {error}
          </div>
        </Card>
      )}

      {/* 导出结果 */}
      {result && (
        <Card className={styles.section}>
          <CardHeader header={<Text weight="semibold">导出结果</Text>} />
          <Divider />

          <div className={`${styles.resultMessage} ${styles.success}`}>
            <strong>导出成功！</strong>
          </div>

          {/* 导出信息 */}
          <div className={styles.infoBox}>
            <div>
              <strong>范围:</strong> {getScopeLabel(result.scope)}
            </div>
            <div>
              <strong>格式:</strong> {result.format.toUpperCase()}
            </div>
            <div>
              <strong>大小:</strong> {formatFileSize(result.size)}
            </div>
            <div>
              <strong>MIME:</strong> {result.mimeType}
            </div>
            <div>
              <strong>时间:</strong> {new Date(result.timestamp).toLocaleString("zh-CN")}
            </div>
          </div>

          {/* 内容预览 */}
          <div>
            <Label>内容预览（前 1000 字符）</Label>
            <Textarea
              className={styles.contentPreview}
              value={
                result.content.substring(0, 1000) + (result.content.length > 1000 ? "..." : "")
              }
              readOnly
              rows={10}
            />
          </div>

          {/* 操作按钮 */}
          <div className={styles.buttonGroup}>
            <Button appearance="primary" icon={<Copy24Regular />} onClick={handleCopyContent}>
              复制内容
            </Button>
            <Button appearance="secondary" onClick={handleDownload}>
              下载文件
            </Button>
          </div>
        </Card>
      )}
    </div>
  );
};

function getScopeLabel(scope: ExportScope): string {
  const labels: Record<ExportScope, string> = {
    selection: "当前选中区域",
    document: "整个文档",
    visible: "当前可见区域",
  };
  return labels[scope];
}
