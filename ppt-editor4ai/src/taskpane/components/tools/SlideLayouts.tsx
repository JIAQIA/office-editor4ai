/**
 * 文件名: SlideLayouts.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 幻灯片布局模板工具，用于获取并显示可用的布局模板列表
 */

/* global console, navigator, setTimeout */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  makeStyles,
  tokens,
  Spinner,
  Switch,
  Label,
  Divider,
  Card,
  Input,
} from "@fluentui/react-components";
import {
  getAvailableSlideLayouts,
  createSlideWithLayout,
  getLayoutDescription,
  type SlideLayoutTemplate,
} from "../../../ppt-tools";
import {
  Copy24Regular,
  CheckmarkCircle24Regular,
  Add24Regular,
  LayoutRowTwo24Regular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
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
  switchContainer: {
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
  layoutsSection: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  layoutCard: {
    padding: "16px",
    cursor: "pointer",
    transition: "all 0.2s ease",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      border: `1px solid ${tokens.colorBrandStroke1}`,
    },
  },
  selectedCard: {
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke1}`,
    ":hover": {
      backgroundColor: tokens.colorBrandBackground2Hover,
    },
  },
  layoutHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
  },
  layoutName: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    color: tokens.colorNeutralForeground1,
  },
  layoutId: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    fontFamily: "monospace",
    wordBreak: "break-all",
  },
  layoutDescription: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginTop: "8px",
  },
  placeholderTags: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
    marginTop: "8px",
  },
  placeholderTag: {
    padding: "2px 8px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
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
  successMessage: {
    color: tokens.colorPaletteGreenForeground1,
    fontSize: tokens.fontSizeBase300,
    padding: "12px",
    textAlign: "center",
    backgroundColor: tokens.colorPaletteGreenBackground1,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorPaletteGreenBorder1}`,
  },
  actionButtons: {
    width: "100%",
    display: "flex",
    gap: "8px",
    marginTop: "8px",
  },
  actionButton: {
    flex: 1,
  },
  statsCard: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  statsText: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground2,
    textAlign: "center",
  },
  positionInput: {
    width: "100%",
    marginTop: "8px",
  },
});

const SlideLayoutsComponent: React.FC = () => {
  const styles = useStyles();
  const [layouts, setLayouts] = useState<SlideLayoutTemplate[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [includePlaceholders, setIncludePlaceholders] = useState(true);
  const [selectedLayoutId, setSelectedLayoutId] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);
  const [creating, setCreating] = useState(false);
  const [insertPosition, setInsertPosition] = useState<string>("");

  const fetchLayouts = async () => {
    setLoading(true);
    setError(null);
    setSuccessMessage(null);
    setCopied(false);

    try {
      const layoutsList = await getAvailableSlideLayouts({
        includePlaceholders,
      });

      setLayouts(layoutsList);
      setSuccessMessage(`成功获取 ${layoutsList.length} 个布局模板`);
      setTimeout(() => setSuccessMessage(null), 3000);
    } catch (err) {
      console.error("获取布局模板失败:", err);
      setError(err instanceof Error ? err.message : "获取布局模板失败");
    } finally {
      setLoading(false);
    }
  };

  const copyToClipboard = async () => {
    if (layouts.length === 0) return;

    try {
      const jsonString = JSON.stringify(layouts, null, 2);
      await navigator.clipboard.writeText(jsonString);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error("复制失败:", err);
      setError("复制到剪贴板失败");
    }
  };

  const handleCreateSlide = async () => {
    if (!selectedLayoutId) {
      setError("请先选择一个布局模板");
      return;
    }

    setCreating(true);
    setError(null);
    setSuccessMessage(null);

    try {
      const position = insertPosition.trim() === "" ? undefined : parseInt(insertPosition, 10);

      if (position !== undefined && (isNaN(position) || position < 0)) {
        setError("插入位置必须是大于等于0的整数");
        setCreating(false);
        return;
      }

      const newSlideId = await createSlideWithLayout(selectedLayoutId, position);
      setSuccessMessage(`成功创建新幻灯片！ID: ${newSlideId}`);
      setTimeout(() => setSuccessMessage(null), 5000);
    } catch (err) {
      console.error("创建幻灯片失败:", err);
      setError(err instanceof Error ? err.message : "创建幻灯片失败");
    } finally {
      setCreating(false);
    }
  };

  const selectedLayout = layouts.find((l) => l.id === selectedLayoutId);

  return (
    <div className={styles.root}>
      {/* 控制区域 */}
      <div className={styles.controlsSection}>
        <div className={styles.switchContainer}>
          <Label>包含占位符详细信息</Label>
          <Switch
            checked={includePlaceholders}
            onChange={(_e, data) => setIncludePlaceholders(data.checked)}
            disabled={loading}
          />
        </div>

        <div className={styles.buttonContainer}>
          <Button appearance="primary" size="large" onClick={fetchLayouts} disabled={loading}>
            {loading ? <Spinner size="tiny" /> : "获取布局模板"}
          </Button>
        </div>
      </div>

      {/* 错误信息 */}
      {error && <div className={styles.errorMessage}>❌ {error}</div>}

      {/* 成功信息 */}
      {successMessage && <div className={styles.successMessage}>✅ {successMessage}</div>}

      {/* 空状态 */}
      {!loading && !error && layouts.length === 0 && (
        <div className={styles.emptyState}>点击按钮获取可用的布局模板列表</div>
      )}

      {/* 布局列表 */}
      {layouts.length > 0 && (
        <>
          {/* 统计信息 */}
          <Card className={styles.statsCard}>
            <div className={styles.statsText}>
              共找到 <strong>{layouts.length}</strong> 个布局模板
              {selectedLayoutId && " · 已选择 1 个"}
            </div>
          </Card>

          <Divider />

          {/* 布局卡片列表 */}
          <div className={styles.layoutsSection}>
            {layouts.map((layout) => (
              <Card
                key={layout.id}
                className={`${styles.layoutCard} ${
                  selectedLayoutId === layout.id ? styles.selectedCard : ""
                }`}
                onClick={() => setSelectedLayoutId(layout.id)}
              >
                <div className={styles.layoutHeader}>
                  <div>
                    <div className={styles.layoutName}>
                      <LayoutRowTwo24Regular
                        style={{ marginRight: "8px", verticalAlign: "middle" }}
                      />
                      {layout.name}
                    </div>
                    <div className={styles.layoutDescription}>{getLayoutDescription(layout)}</div>
                  </div>
                </div>

                {/* 占位符标签 */}
                {layout.placeholderTypes.length > 0 && (
                  <div className={styles.placeholderTags}>
                    {Array.from(new Set(layout.placeholderTypes)).map((type, index) => (
                      <span key={index} className={styles.placeholderTag}>
                        {type}
                      </span>
                    ))}
                  </div>
                )}

                {/* 布局ID（折叠显示） */}
                <div className={styles.layoutId} style={{ marginTop: "8px" }}>
                  ID: {layout.id}
                </div>
              </Card>
            ))}
          </div>

          <Divider />

          {/* 操作按钮 */}
          <div className={styles.actionButtons}>
            <Button
              appearance="secondary"
              size="large"
              icon={
                copied ? (
                  <CheckmarkCircle24Regular style={{ color: tokens.colorPaletteGreenForeground1 }} />
                ) : (
                  <Copy24Regular />
                )
              }
              onClick={copyToClipboard}
              className={styles.actionButton}
            >
              {copied ? "已复制" : "复制 JSON"}
            </Button>
          </div>

          {/* 创建幻灯片区域 */}
          {selectedLayout && (
            <>
              <Divider />
              <Card className={styles.statsCard}>
                <div className={styles.statsText} style={{ marginBottom: "12px" }}>
                  使用布局 <strong>{selectedLayout.name}</strong> 创建新幻灯片
                </div>

                <div className={styles.positionInput}>
                  <Label>插入位置（可选，从0开始，不填则插入到末尾）</Label>
                  <Input
                    type="number"
                    min="0"
                    placeholder="例如: 0 表示插入到开头"
                    value={insertPosition}
                    onChange={(e) => setInsertPosition(e.target.value)}
                    disabled={creating}
                  />
                </div>

                <div className={styles.actionButtons}>
                  <Button
                    appearance="primary"
                    size="large"
                    icon={<Add24Regular />}
                    onClick={handleCreateSlide}
                    disabled={creating}
                    className={styles.actionButton}
                  >
                    {creating ? <Spinner size="tiny" /> : "创建新幻灯片"}
                  </Button>
                </div>
              </Card>
            </>
          )}
        </>
      )}
    </div>
  );
};

export default SlideLayoutsComponent;
