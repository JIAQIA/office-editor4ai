/**
 * 文件名: SlideLayoutInfo.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 页面布局信息工具，用于获取并显示完整的页面布局、尺寸和元素详细信息
 */

/* global console */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  makeStyles,
  tokens,
  Spinner,
  Input,
  Label,
  Switch,
  Divider,
  Card,
} from "@fluentui/react-components";
import { getSlideLayoutInfo, type SlideLayoutInfo, type EnhancedElement } from "../../../ppt-tools";
import { Copy24Regular, CheckmarkCircle24Regular } from "@fluentui/react-icons";

// Office.js 错误类型定义
interface OfficeError extends Error {
  debugInfo?: {
    errorCode?: string;
    errorLocation?: string;
    message?: string;
    [key: string]: any;
  };
  code?: string;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
    gap: "16px",
  },
  controlsSection: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    marginBottom: "8px",
  },
  inputContainer: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  inputField: {
    width: "100%",
  },
  switchContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  switchItem: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  buttonContainer: {
    width: "100%",
    display: "flex",
    gap: "8px",
    justifyContent: "center",
  },
  emptyState: {
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
    padding: "32px 16px",
  },
  infoSection: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  sectionCard: {
    padding: "16px",
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
    marginBottom: "12px",
  },
  infoGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "12px",
  },
  infoItem: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  infoLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  infoValue: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  elementsList: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  elementCard: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    padding: "12px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  elementHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
  },
  elementType: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorBrandForeground1,
  },
  elementIndex: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    fontFamily: "monospace",
  },
  elementDetails: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    fontSize: tokens.fontSizeBase200,
  },
  detailRow: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "8px",
  },
  detailItem: {
    display: "flex",
    flexDirection: "column",
  },
  detailLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    marginBottom: "2px",
  },
  detailValue: {
    color: tokens.colorNeutralForeground1,
    fontWeight: tokens.fontWeightSemibold,
    fontFamily: "monospace",
  },
  textContent: {
    marginTop: "8px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground1,
    wordBreak: "break-word",
    lineHeight: "1.4",
    maxHeight: "100px",
    overflow: "auto",
  },
  errorMessage: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase300,
    padding: "16px",
    textAlign: "left",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorPaletteRedBorder1}`,
    fontFamily: "monospace",
    maxHeight: "300px",
    overflow: "auto",
  },
  copyButton: {
    width: "100%",
  },
  successIcon: {
    color: tokens.colorPaletteGreenForeground1,
  },
});

const SlideLayoutInfo: React.FC = () => {
  const styles = useStyles();
  const [layoutInfo, setLayoutInfo] = useState<SlideLayoutInfo | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [slideNumber, setSlideNumber] = useState<string>("");
  const [includeImages, setIncludeImages] = useState(false);
  const [includeBackground, setIncludeBackground] = useState(false);
  const [includeTextDetails, setIncludeTextDetails] = useState(false);
  const [copied, setCopied] = useState(false);

  const fetchLayoutInfo = async () => {
    setLoading(true);
    setError(null);
    setCopied(false);

    try {
      const pageNum = slideNumber.trim() === "" ? undefined : parseInt(slideNumber, 10);

      if (pageNum !== undefined && (isNaN(pageNum) || pageNum < 1)) {
        setError("页码必须是大于0的整数");
        setLoading(false);
        return;
      }

      const info = await getSlideLayoutInfo({
        slideNumber: pageNum,
        includeImages,
        includeBackground,
        includeTextDetails,
      });

      setLayoutInfo(info);
    } catch (err) {
      console.error("获取布局信息失败 - 完整错误:", err);
      console.error("错误名称:", (err as Error)?.name);
      console.error("错误消息:", (err as Error)?.message);
      console.error("错误堆栈:", (err as Error)?.stack);
      
      // 打印 Office.js 特定的调试信息
      const officeErr = err as OfficeError;
      if (officeErr?.debugInfo) {
        console.error("Office.js 调试信息:", JSON.stringify(officeErr.debugInfo, null, 2));
      }
      
      // 打印完整的错误对象
      console.error("完整错误对象:", JSON.stringify(err, Object.getOwnPropertyNames(err), 2));
      
      // 构建更详细的错误消息
      let errorMessage = "获取布局信息失败";
      if (err instanceof Error) {
        errorMessage = err.message;
        
        // 如果有 Office.js 调试信息，添加到错误消息中
        if (officeErr.debugInfo) {
          errorMessage += `\n\n调试信息:\n${JSON.stringify(officeErr.debugInfo, null, 2)}`;
        }
      }
      
      setError(errorMessage);
    } finally {
      setLoading(false);
    }
  };

  const copyToClipboard = async () => {
    if (!layoutInfo) return;

    try {
      const jsonString = JSON.stringify(layoutInfo, null, 2);
      await navigator.clipboard.writeText(jsonString);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error("复制失败:", err);
      setError("复制到剪贴板失败");
    }
  };

  const renderElement = (element: EnhancedElement, index: number) => (
    <div key={element.id} className={styles.elementCard}>
      <div className={styles.elementHeader}>
        <span className={styles.elementType}>{element.type}</span>
        <span className={styles.elementIndex}>#{index + 1}</span>
      </div>

      <div className={styles.elementDetails}>
        {/* 绝对位置和尺寸 */}
        <div className={styles.detailRow}>
          <div className={styles.detailItem}>
            <span className={styles.detailLabel}>位置 (X, Y)</span>
            <span className={styles.detailValue}>
              {element.left}, {element.top}
            </span>
          </div>
          <div className={styles.detailItem}>
            <span className={styles.detailLabel}>尺寸 (W × H)</span>
            <span className={styles.detailValue}>
              {element.width} × {element.height}
            </span>
          </div>
        </div>

        {/* 相对位置（百分比） */}
        <div className={styles.detailRow}>
          <div className={styles.detailItem}>
            <span className={styles.detailLabel}>相对位置 (%)</span>
            <span className={styles.detailValue}>
              {element.relativePosition.leftPercent.toFixed(1)}%, {element.relativePosition.topPercent.toFixed(1)}%
            </span>
          </div>
          <div className={styles.detailItem}>
            <span className={styles.detailLabel}>相对尺寸 (%)</span>
            <span className={styles.detailValue}>
              {element.relativePosition.widthPercent.toFixed(1)}% × {element.relativePosition.heightPercent.toFixed(1)}
              %
            </span>
          </div>
        </div>

        {/* 其他属性 */}
        {(element.rotation !== undefined || element.zOrder !== undefined || element.name) && (
          <div className={styles.detailRow}>
            {element.rotation !== undefined && (
              <div className={styles.detailItem}>
                <span className={styles.detailLabel}>旋转角度</span>
                <span className={styles.detailValue}>{element.rotation}°</span>
              </div>
            )}
            {element.zOrder !== undefined && (
              <div className={styles.detailItem}>
                <span className={styles.detailLabel}>Z轴顺序</span>
                <span className={styles.detailValue}>{element.zOrder}</span>
              </div>
            )}
          </div>
        )}

        {element.name && (
          <div className={styles.detailItem}>
            <span className={styles.detailLabel}>元素名称</span>
            <span className={styles.detailValue}>{element.name}</span>
          </div>
        )}

        {/* 填充信息 */}
        {element.fill && (
          <div className={styles.detailItem}>
            <span className={styles.detailLabel}>填充类型</span>
            <span className={styles.detailValue}>
              {element.fill.type}
              {element.fill.color && ` (${element.fill.color})`}
            </span>
          </div>
        )}

        {/* 文本内容 */}
        {element.text && (
          <div className={styles.textContent}>
            <div className={styles.detailLabel}>文本内容:</div>
            {element.text.content}
            {element.text.fontSize && (
              <div style={{ marginTop: "4px", fontSize: "11px", color: tokens.colorNeutralForeground3 }}>
                字体: {element.text.fontFamily || "未知"} | 大小: {element.text.fontSize}
                {element.text.color && ` | 颜色: ${element.text.color}`}
              </div>
            )}
          </div>
        )}

        {/* 图片信息 */}
        {element.image && (
          <div className={styles.detailItem}>
            <span className={styles.detailLabel}>图片格式</span>
            <span className={styles.detailValue}>{element.image.format}</span>
          </div>
        )}
      </div>
    </div>
  );

  return (
    <div className={styles.container}>
      {/* 控制区域 */}
      <div className={styles.controlsSection}>
        <div className={styles.inputContainer}>
          <Label htmlFor="slideNumber">页码（可选，不填则使用当前页）</Label>
          <Input
            id="slideNumber"
            type="number"
            min="1"
            placeholder="请输入页码，从1开始"
            value={slideNumber}
            onChange={(e) => setSlideNumber(e.target.value)}
            className={styles.inputField}
            disabled={loading}
          />
        </div>

        <div className={styles.switchContainer}>
          <div className={styles.switchItem}>
            <Label>包含图片数据</Label>
            <Switch checked={includeImages} onChange={(_e, data) => setIncludeImages(data.checked)} disabled={loading} />
          </div>
          <div className={styles.switchItem}>
            <Label>包含背景信息</Label>
            <Switch
              checked={includeBackground}
              onChange={(_e, data) => setIncludeBackground(data.checked)}
              disabled={loading}
            />
          </div>
          <div className={styles.switchItem}>
            <Label>包含文本详细信息</Label>
            <Switch
              checked={includeTextDetails}
              onChange={(_e, data) => setIncludeTextDetails(data.checked)}
              disabled={loading}
            />
          </div>
        </div>

        <div className={styles.buttonContainer}>
          <Button appearance="primary" size="large" onClick={fetchLayoutInfo} disabled={loading}>
            {loading ? <Spinner size="tiny" /> : "获取布局信息"}
          </Button>
        </div>
      </div>

      {/* 错误信息 */}
      {error && <div className={styles.errorMessage}>❌ {error}</div>}

      {/* 空状态 */}
      {!loading && !error && !layoutInfo && (
        <div className={styles.emptyState}>配置选项并点击按钮获取页面布局信息</div>
      )}

      {/* 布局信息展示 */}
      {layoutInfo && (
        <div className={styles.infoSection}>
          {/* 基础信息卡片 */}
          <Card className={styles.sectionCard}>
            <div className={styles.sectionTitle}>📄 页面信息</div>
            <div className={styles.infoGrid}>
              <div className={styles.infoItem}>
                <span className={styles.infoLabel}>页码</span>
                <span className={styles.infoValue}>{layoutInfo.slideNumber}</span>
              </div>
              <div className={styles.infoItem}>
                <span className={styles.infoLabel}>布局类型</span>
                <span className={styles.infoValue}>{layoutInfo.layout.name}</span>
              </div>
              <div className={styles.infoItem}>
                <span className={styles.infoLabel}>尺寸 (points)</span>
                <span className={styles.infoValue}>
                  {layoutInfo.dimensions.width} × {layoutInfo.dimensions.height}
                </span>
              </div>
              <div className={styles.infoItem}>
                <span className={styles.infoLabel}>宽高比</span>
                <span className={styles.infoValue}>{layoutInfo.dimensions.aspectRatio}</span>
              </div>
              <div className={styles.infoItem}>
                <span className={styles.infoLabel}>尺寸来源</span>
                <span className={styles.infoValue}>
                  {layoutInfo.dimensions.isFromAPI ? "✅ API获取" : "⚠️ 默认值"}
                </span>
              </div>
              <div className={styles.infoItem}>
                <span className={styles.infoLabel}>元素数量</span>
                <span className={styles.infoValue}>{layoutInfo.elements.length}</span>
              </div>
              {layoutInfo.background && (
                <div className={styles.infoItem}>
                  <span className={styles.infoLabel}>背景类型</span>
                  <span className={styles.infoValue}>
                    {layoutInfo.background.type}
                    {layoutInfo.background.color && ` (${layoutInfo.background.color})`}
                  </span>
                </div>
              )}
            </div>
          </Card>

          <Divider />

          {/* 元素列表 */}
          {layoutInfo.elements.length > 0 && (
            <>
              <div className={styles.sectionTitle}>🎨 元素列表</div>
              <div className={styles.elementsList}>{layoutInfo.elements.map(renderElement)}</div>
            </>
          )}

          <Divider />

          {/* 复制 JSON 按钮 */}
          <Button
            appearance="secondary"
            size="large"
            icon={copied ? <CheckmarkCircle24Regular className={styles.successIcon} /> : <Copy24Regular />}
            onClick={copyToClipboard}
            className={styles.copyButton}
          >
            {copied ? "已复制到剪贴板" : "复制 JSON 数据"}
          </Button>
        </div>
      )}
    </div>
  );
};

export default SlideLayoutInfo;
