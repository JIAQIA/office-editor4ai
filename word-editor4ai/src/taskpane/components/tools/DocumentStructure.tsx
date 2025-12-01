/**
 * 文件名: DocumentStructure.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/01
 * 最后修改日期: 2025/12/01
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 文档结构获取工具，用于获取并显示文档大纲结构
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
  Tree,
  TreeItem,
  TreeItemLayout,
  Badge,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Tooltip,
} from "@fluentui/react-components";
import {
  DocumentBulletList24Regular,
  Navigation24Regular,
  ArrowDownload24Regular,
  ChevronRight20Regular,
  ChevronDown20Regular,
  Info24Regular,
} from "@fluentui/react-icons";
import {
  getDocumentOutline,
  getDocumentOutlineFlat,
  navigateToOutlineNode,
  exportOutlineAsMarkdown,
  exportOutlineAsJSON,
  type DocumentOutline,
  type OutlineNode,
  type GetDocumentOutlineOptions,
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
    flexWrap: "wrap",
  },
  resultContainer: {
    width: "100%",
    maxHeight: "calc(100vh - 300px)",
    overflowY: "auto",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    padding: "12px",
  },
  statsCard: {
    marginBottom: "12px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  statsGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(2, 1fr)",
    gap: "8px",
    marginTop: "8px",
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
  treeContainer: {
    width: "100%",
  },
  treeItem: {
    cursor: "pointer",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground3Hover,
    },
  },
  nodeContent: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    width: "100%",
  },
  nodeText: {
    flex: 1,
    fontSize: tokens.fontSizeBase300,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  nodeBadge: {
    flexShrink: 0,
  },
  formatInfo: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
  },
  emptyState: {
    textAlign: "center",
    padding: "32px 16px",
    color: tokens.colorNeutralForeground3,
  },
  errorMessage: {
    color: tokens.colorPaletteRedForeground1,
    padding: "12px",
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "8px",
  },
});

/**
 * 获取标题级别的颜色
 */
const getLevelColor = (level: number): "brand" | "success" | "warning" | "danger" | "important" | "informative" => {
  const colors: Array<"brand" | "success" | "warning" | "danger" | "important" | "informative"> = [
    "brand",
    "success",
    "warning",
    "danger",
    "important",
    "informative",
  ];
  return colors[(level - 1) % colors.length];
};

/**
 * 递归渲染大纲树节点
 */
const OutlineTreeNode: React.FC<{
  node: OutlineNode;
  includeFormat: boolean;
  onNavigate: (nodeId: string) => void;
}> = ({ node, includeFormat, onNavigate }) => {
  const styles = useStyles();
  const [isExpanded, setIsExpanded] = useState(true);

  const handleClick = () => {
    onNavigate(node.id);
  };

  const formatText = node.format
    ? `${node.format.font || "默认"} ${node.format.fontSize || ""}pt ${node.format.bold ? "粗体" : ""} ${
        node.format.italic ? "斜体" : ""
      }`
    : "";

  return (
    <TreeItem
      itemType={node.children.length > 0 ? "branch" : "leaf"}
      value={node.id}
      className={styles.treeItem}
    >
      <TreeItemLayout
        iconBefore={
          node.children.length > 0 ? (
            isExpanded ? (
              <ChevronDown20Regular />
            ) : (
              <ChevronRight20Regular />
            )
          ) : null
        }
        onClick={() => {
          if (node.children.length > 0) {
            setIsExpanded(!isExpanded);
          }
        }}
      >
        <div className={styles.nodeContent}>
          <Tooltip content={`点击跳转到此标题 (索引: ${node.index})`} relationship="label">
            <span className={styles.nodeText} onClick={handleClick}>
              {node.text || "(空标题)"}
            </span>
          </Tooltip>
          <Badge appearance="filled" color={getLevelColor(node.level)} className={styles.nodeBadge}>
            H{node.level}
          </Badge>
        </div>
        {includeFormat && node.format && (
          <div className={styles.formatInfo}>{formatText}</div>
        )}
      </TreeItemLayout>
      {node.children.length > 0 && isExpanded && (
        <Tree>
          {node.children.map((child) => (
            <OutlineTreeNode
              key={child.id}
              node={child}
              includeFormat={includeFormat}
              onNavigate={onNavigate}
            />
          ))}
        </Tree>
      )}
    </TreeItem>
  );
};

/**
 * 文档结构工具组件
 */
export const DocumentStructure: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [outline, setOutline] = useState<DocumentOutline | null>(null);
  const [error, setError] = useState<string | null>(null);

  // 选项状态
  const [includeFormat, setIncludeFormat] = useState(false);
  const [useTreeStructure, setUseTreeStructure] = useState(true);
  const [maxDepth, setMaxDepth] = useState(0);

  /**
   * 获取文档大纲
   */
  const handleGetOutline = async () => {
    setLoading(true);
    setError(null);

    try {
      const options: GetDocumentOutlineOptions = {
        includeFormat,
        maxDepth: maxDepth > 0 ? maxDepth : undefined,
      };

      if (useTreeStructure) {
        const result = await getDocumentOutline(options);
        setOutline(result);
      } else {
        const nodes = await getDocumentOutlineFlat(options);
        // 转换为 DocumentOutline 格式
        const levelCounts: Record<number, number> = {};
        let maxDepthFound = 0;

        nodes.forEach((node) => {
          levelCounts[node.level] = (levelCounts[node.level] || 0) + 1;
          maxDepthFound = Math.max(maxDepthFound, node.level);
        });

        setOutline({
          nodes,
          totalHeadings: nodes.length,
          maxDepth: maxDepthFound,
          levelCounts,
        });
      }
    } catch (err) {
      console.error("获取文档大纲失败:", err);
      setError(err.message || "获取文档大纲失败");
    } finally {
      setLoading(false);
    }
  };

  /**
   * 跳转到大纲节点
   */
  const handleNavigateToNode = async (nodeId: string) => {
    try {
      await navigateToOutlineNode(nodeId);
    } catch (err) {
      console.error("跳转失败:", err);
      setError(err.message || "跳转失败");
    }
  };

  /**
   * 导出为 Markdown
   */
  const handleExportMarkdown = () => {
    if (!outline) return;

    const markdown = exportOutlineAsMarkdown(outline);
    const blob = new Blob([markdown], { type: "text/markdown" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "document-outline.md";
    a.click();
    URL.revokeObjectURL(url);
  };

  /**
   * 导出为 JSON
   */
  const handleExportJSON = () => {
    if (!outline) return;

    const json = exportOutlineAsJSON(outline);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "document-outline.json";
    a.click();
    URL.revokeObjectURL(url);
  };

  /**
   * 复制到剪贴板
   */
  const handleCopyToClipboard = async () => {
    if (!outline) return;

    const json = exportOutlineAsJSON(outline);
    try {
      await navigator.clipboard.writeText(json);
      alert("已复制到剪贴板");
    } catch (err) {
      console.error("复制失败:", err);
      setError("复制到剪贴板失败");
    }
  };

  return (
    <div className={styles.container}>
      {/* 标题 */}
      <Card>
        <CardHeader
          image={<DocumentBulletList24Regular />}
          header={<strong>文档结构获取</strong>}
          description="获取文档的大纲结构（标题层级）"
        />
      </Card>

      {/* 选项配置 */}
      <div className={styles.optionsContainer}>
        <div className={styles.optionRow}>
          <Label>包含格式信息</Label>
          <Switch checked={includeFormat} onChange={(_, data) => setIncludeFormat(data.checked)} />
        </div>

        <div className={styles.optionRow}>
          <Label>树形结构</Label>
          <Switch checked={useTreeStructure} onChange={(_, data) => setUseTreeStructure(data.checked)} />
        </div>

        <div className={styles.optionRow}>
          <Label>
            最大层级深度 (0=不限制)
            <Tooltip content="限制获取的标题层级深度，例如设置为2只获取H1和H2" relationship="label">
              <Info24Regular style={{ marginLeft: "4px", fontSize: "14px" }} />
            </Tooltip>
          </Label>
          <input
            type="number"
            min="0"
            max="9"
            value={maxDepth}
            onChange={(e) => setMaxDepth(parseInt(e.target.value) || 0)}
            style={{
              width: "60px",
              padding: "4px 8px",
              borderRadius: "4px",
              border: `1px solid ${tokens.colorNeutralStroke1}`,
            }}
          />
        </div>
      </div>

      {/* 操作按钮 */}
      <div className={styles.buttonGroup}>
        <Button
          appearance="primary"
          icon={<DocumentBulletList24Regular />}
          onClick={handleGetOutline}
          disabled={loading}
        >
          {loading ? <Spinner size="tiny" /> : "获取大纲"}
        </Button>

        {outline && (
          <>
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <Button icon={<ArrowDownload24Regular />}>导出</Button>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem onClick={handleExportMarkdown}>导出为 Markdown</MenuItem>
                  <MenuItem onClick={handleExportJSON}>导出为 JSON</MenuItem>
                  <MenuItem onClick={handleCopyToClipboard}>复制到剪贴板</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </>
        )}
      </div>

      {/* 错误信息 */}
      {error && <div className={styles.errorMessage}>{error}</div>}

      {/* 结果显示 */}
      {outline && (
        <div className={styles.resultContainer}>
          {/* 统计信息 */}
          <div className={styles.statsCard}>
            <strong>统计信息</strong>
            <div className={styles.statsGrid}>
              <div className={styles.statItem}>
                <span className={styles.statLabel}>总标题数</span>
                <span className={styles.statValue}>{outline.totalHeadings}</span>
              </div>
              <div className={styles.statItem}>
                <span className={styles.statLabel}>最大层级</span>
                <span className={styles.statValue}>{outline.maxDepth}</span>
              </div>
              {Object.entries(outline.levelCounts).map(([level, count]) => (
                <div key={level} className={styles.statItem}>
                  <span className={styles.statLabel}>H{level} 数量</span>
                  <span className={styles.statValue}>{count}</span>
                </div>
              ))}
            </div>
          </div>

          <Divider />

          {/* 大纲树 */}
          {outline.nodes.length > 0 ? (
            <div className={styles.treeContainer}>
              <Tree>
                {outline.nodes.map((node) => (
                  <OutlineTreeNode
                    key={node.id}
                    node={node}
                    includeFormat={includeFormat}
                    onNavigate={handleNavigateToNode}
                  />
                ))}
              </Tree>
            </div>
          ) : (
            <div className={styles.emptyState}>
              <DocumentBulletList24Regular />
              <p>文档中没有找到标题</p>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default DocumentStructure;
