/**
 * 文件名: Sidebar.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components, @fluentui/react-icons
 * 描述: 侧边栏导航组件，支持二级菜单和展开/折叠
 */

import * as React from "react";
import { useState } from "react";
import { makeStyles, tokens, Button } from "@fluentui/react-components";
import {
  Home24Regular,
  Add24Regular,
  Delete24Regular,
  Edit24Regular,
  Search24Regular,
  TextGrammarSettings24Regular,
  ChevronDown24Regular,
  Navigation24Regular,
  Dismiss24Regular,
  Eye24Regular,
  DocumentBulletList24Regular,
  DocumentPageBreak24Regular,
  DocumentSplitHint24Regular,
  DocumentData24Regular,
  DocumentOnePage24Regular,
  TextEffects24Regular,
  DocumentHeader24Regular,
  Textbox24Regular,
  Comment24Regular,
  ArrowSwap24Regular,
  ArrowDown24Regular,
  Image24Regular,
  Table24Regular,
  Shapes24Regular,
  DocumentTable24Regular,
  MathFormula24Regular,
  DocumentArrowDown24Regular,
} from "@fluentui/react-icons";

interface SidebarProps {
  currentPage: string;
  currentTool: string;
  isCollapsed: boolean;
  onNavigate: (page: string, tool?: string) => void;
  onToggleCollapse: () => void;
}

const useStyles = makeStyles({
  sidebar: {
    width: "180px",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRight: `1px solid ${tokens.colorNeutralStroke1}`,
    display: "flex",
    flexDirection: "column",
    transition: "width 0.3s ease",
    flexShrink: 0,
  },
  sidebarCollapsed: {
    width: "48px",
  },
  scrollContainer: {
    flex: 1,
    overflowY: "auto",
    overflowX: "hidden",
    padding: "8px 0",
    // 隐藏滚动条但保持滚动功能 / Hide scrollbar but keep scroll functionality
    scrollbarWidth: "thin",
    scrollbarColor: "transparent transparent",
    "::-webkit-scrollbar": {
      width: "6px",
    },
    "::-webkit-scrollbar-track": {
      background: "transparent",
    },
    "::-webkit-scrollbar-thumb": {
      background: "transparent",
      borderRadius: "3px",
    },
    ":hover": {
      scrollbarColor: `${tokens.colorNeutralStroke2} transparent`,
      "::-webkit-scrollbar-thumb": {
        background: tokens.colorNeutralStroke2,
      },
    },
  },
  scrollContainerCollapsed: {
    padding: "8px 2px",
  },
  toggleButtonWrapper: {
    padding: "8px 0",
    flexShrink: 0,
  },
  toggleButtonWrapperCollapsed: {
    padding: "8px 2px",
  },
  toggleButton: {
    width: "100%",
    justifyContent: "center",
    padding: "8px 4px",
    marginBottom: "8px",
    border: "none",
    backgroundColor: "transparent",
    cursor: "pointer",
    height: "36px",
    minHeight: "36px",
    minWidth: "unset",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground3Hover,
    },
  },
  toggleButtonCollapsed: {
    padding: "6px",
    margin: "0 auto",
    marginBottom: "8px",
    width: "40px",
    minWidth: "40px",
    maxWidth: "40px",
  },
  menuItem: {
    width: "100%",
    justifyContent: "flex-start",
    padding: "10px 12px",
    border: "none",
    backgroundColor: "transparent",
    cursor: "pointer",
    transition: "background-color 0.2s",
    height: "44px",
    minHeight: "44px",
    minWidth: "unset",
    overflow: "hidden",
  },
  menuItemCollapsed: {
    padding: "10px 4px",
    justifyContent: "center",
    minWidth: "unset",
  },
  menuItemActive: {
    backgroundColor: tokens.colorNeutralBackground3Selected,
  },
  menuItemHover: {
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground3Hover,
    },
  },
  menuItemContent: {
    display: "flex",
    alignItems: "center",
    width: "100%",
    overflow: "hidden",
  },
  icon: {
    marginRight: "10px",
    fontSize: "24px",
    width: "24px",
    height: "24px",
    flexShrink: 0,
  },
  iconCollapsed: {
    marginRight: "0",
    fontSize: "24px",
  },
  label: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    flex: 1,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    opacity: 1,
    transition: "opacity 0.2s ease",
  },
  chevron: {
    fontSize: "16px",
    marginLeft: "auto",
    flexShrink: 0,
    transition: "transform 0.3s ease",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
  },
  submenu: {
    paddingLeft: "12px",
    overflow: "hidden",
    maxHeight: "500px",
    opacity: 1,
    transform: "translateY(0)",
    transition: "max-height 0.3s ease, opacity 0.3s ease, transform 0.3s ease",
  },
  submenuCollapsed: {
    maxHeight: "0",
    opacity: 0,
    transform: "translateY(-10px)",
  },
  submenuItem: {
    width: "100%",
    justifyContent: "flex-start",
    padding: "8px 12px 8px 38px",
    border: "none",
    backgroundColor: "transparent",
    cursor: "pointer",
    transition: "background-color 0.2s",
    height: "40px",
    minHeight: "40px",
    overflow: "hidden",
  },
  submenuItemActive: {
    backgroundColor: tokens.colorNeutralBackground3Selected,
  },
  submenuItemHover: {
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground3Hover,
    },
  },
});

const Sidebar: React.FC<SidebarProps> = ({
  currentPage,
  currentTool,
  isCollapsed,
  onNavigate,
  onToggleCollapse,
}) => {
  const styles = useStyles();
  const [createExpanded, setCreateExpanded] = useState(true);
  const [deleteExpanded, setDeleteExpanded] = useState(false);
  const [updateExpanded, setUpdateExpanded] = useState(false);
  const [queryExpanded, setQueryExpanded] = useState(false);

  return (
    <div className={`${styles.sidebar} ${isCollapsed ? styles.sidebarCollapsed : ""}`}>
      {/* 折叠/展开按钮 / Toggle button */}
      <div
        className={`${styles.toggleButtonWrapper} ${isCollapsed ? styles.toggleButtonWrapperCollapsed : ""}`}
      >
        <Button
          appearance="subtle"
          className={`${styles.toggleButton} ${isCollapsed ? styles.toggleButtonCollapsed : ""}`}
          onClick={onToggleCollapse}
          title={isCollapsed ? "展开侧边栏" : "折叠侧边栏"}
        >
          {isCollapsed ? <Navigation24Regular /> : <Dismiss24Regular />}
        </Button>
      </div>

      {/* 可滚动菜单容器 / Scrollable menu container */}
      <div
        className={`${styles.scrollContainer} ${isCollapsed ? styles.scrollContainerCollapsed : ""}`}
      >
        {/* 首页菜单项 */}
        <Button
          appearance="subtle"
          className={`${styles.menuItem} ${isCollapsed ? styles.menuItemCollapsed : ""} ${
            currentPage === "home" ? styles.menuItemActive : styles.menuItemHover
          }`}
          onClick={() => onNavigate("home")}
          title={isCollapsed ? "首页" : ""}
        >
          <div className={styles.menuItemContent}>
            <Home24Regular
              className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`}
            />
            {!isCollapsed && <span className={styles.label}>首页</span>}
          </div>
        </Button>

        {/* 创建元素类菜单项 */}
        <Button
          appearance="subtle"
          className={`${styles.menuItem} ${isCollapsed ? styles.menuItemCollapsed : ""} ${
            currentPage === "create" ? styles.menuItemActive : styles.menuItemHover
          }`}
          onClick={() => {
            if (isCollapsed) {
              onNavigate("create", "text-insertion");
            } else {
              setCreateExpanded(!createExpanded);
            }
          }}
          title={isCollapsed ? "创建元素类" : ""}
        >
          <div className={styles.menuItemContent}>
            <Add24Regular className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`} />
            {!isCollapsed && (
              <>
                <span className={styles.label}>创建元素类</span>
                <div
                  className={styles.chevron}
                  style={{ transform: createExpanded ? "rotate(0deg)" : "rotate(-90deg)" }}
                >
                  <ChevronDown24Regular />
                </div>
              </>
            )}
          </div>
        </Button>

        {/* 创建元素类二级菜单 */}
        {!isCollapsed && (
          <div className={`${styles.submenu} ${!createExpanded ? styles.submenuCollapsed : ""}`}>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "text-insertion"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "text-insertion")}
            >
              <div className={styles.menuItemContent}>
                <TextGrammarSettings24Regular className={styles.icon} />
                <span className={styles.label}>文本插入工具</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "append-text"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "append-text")}
            >
              <div className={styles.menuItemContent}>
                <ArrowDown24Regular className={styles.icon} />
                <span className={styles.label}>文档末尾追加文本</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "insert-image"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "insert-image")}
            >
              <div className={styles.menuItemContent}>
                <Image24Regular className={styles.icon} />
                <span className={styles.label}>插入图片</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "create-table"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "create-table")}
            >
              <div className={styles.menuItemContent}>
                <Table24Regular className={styles.icon} />
                <span className={styles.label}>创建表格</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "insert-textbox"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "insert-textbox")}
            >
              <div className={styles.menuItemContent}>
                <Textbox24Regular className={styles.icon} />
                <span className={styles.label}>插入文本框</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "insert-shape"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "insert-shape")}
            >
              <div className={styles.menuItemContent}>
                <Shapes24Regular className={styles.icon} />
                <span className={styles.label}>插入形状</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "insert-page-break"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "insert-page-break")}
            >
              <div className={styles.menuItemContent}>
                <DocumentPageBreak24Regular className={styles.icon} />
                <span className={styles.label}>插入分页符</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "insert-section-break"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "insert-section-break")}
            >
              <div className={styles.menuItemContent}>
                <DocumentSplitHint24Regular className={styles.icon} />
                <span className={styles.label}>插入分节符</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "insert-toc"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "insert-toc")}
            >
              <div className={styles.menuItemContent}>
                <DocumentTable24Regular className={styles.icon} />
                <span className={styles.label}>插入目录</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "create" && currentTool === "insert-equation"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("create", "insert-equation")}
            >
              <div className={styles.menuItemContent}>
                <MathFormula24Regular className={styles.icon} />
                <span className={styles.label}>插入公式</span>
              </div>
            </Button>
          </div>
        )}

        {/* 删除元素类菜单项 */}
        <Button
          appearance="subtle"
          className={`${styles.menuItem} ${isCollapsed ? styles.menuItemCollapsed : ""} ${
            currentPage === "delete" ? styles.menuItemActive : styles.menuItemHover
          }`}
          onClick={() => {
            if (isCollapsed) {
              onNavigate("delete", "delete-table");
            } else {
              setDeleteExpanded(!deleteExpanded);
            }
          }}
          title={isCollapsed ? "删除元素类" : ""}
        >
          <div className={styles.menuItemContent}>
            <Delete24Regular
              className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`}
            />
            {!isCollapsed && (
              <>
                <span className={styles.label}>删除元素类</span>
                <div
                  className={styles.chevron}
                  style={{ transform: deleteExpanded ? "rotate(0deg)" : "rotate(-90deg)" }}
                >
                  <ChevronDown24Regular />
                </div>
              </>
            )}
          </div>
        </Button>

        {/* 删除元素类二级菜单 */}
        {!isCollapsed && (
          <div className={`${styles.submenu} ${!deleteExpanded ? styles.submenuCollapsed : ""}`}>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "delete" && currentTool === "delete-table"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("delete", "delete-table")}
            >
              <div className={styles.menuItemContent}>
                <Table24Regular className={styles.icon} />
                <span className={styles.label}>删除表格</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "delete" && currentTool === "delete-toc"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("delete", "delete-toc")}
            >
              <div className={styles.menuItemContent}>
                <DocumentTable24Regular className={styles.icon} />
                <span className={styles.label}>删除目录</span>
              </div>
            </Button>
          </div>
        )}

        {/* 修改元素类菜单项 */}
        <Button
          appearance="subtle"
          className={`${styles.menuItem} ${isCollapsed ? styles.menuItemCollapsed : ""} ${
            currentPage === "update" ? styles.menuItemActive : styles.menuItemHover
          }`}
          onClick={() => {
            if (isCollapsed) {
              onNavigate("update", "replace-selection");
            } else {
              setUpdateExpanded(!updateExpanded);
            }
          }}
          title={isCollapsed ? "修改元素类" : ""}
        >
          <div className={styles.menuItemContent}>
            <Edit24Regular
              className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`}
            />
            {!isCollapsed && (
              <>
                <span className={styles.label}>修改元素类</span>
                <div
                  className={styles.chevron}
                  style={{ transform: updateExpanded ? "rotate(0deg)" : "rotate(-90deg)" }}
                >
                  <ChevronDown24Regular />
                </div>
              </>
            )}
          </div>
        </Button>

        {/* 修改元素类二级菜单 */}
        {!isCollapsed && (
          <div className={`${styles.submenu} ${!updateExpanded ? styles.submenuCollapsed : ""}`}>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "update" && currentTool === "replace-selection"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("update", "replace-selection")}
            >
              <div className={styles.menuItemContent}>
                <ArrowSwap24Regular className={styles.icon} />
                <span className={styles.label}>替换选中内容</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "update" && currentTool === "replace-text"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("update", "replace-text")}
            >
              <div className={styles.menuItemContent}>
                <TextEffects24Regular className={styles.icon} />
                <span className={styles.label}>替换文本</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "update" && currentTool === "replace-image"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("update", "replace-image")}
            >
              <div className={styles.menuItemContent}>
                <Image24Regular className={styles.icon} />
                <span className={styles.label}>替换图片</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "update" && currentTool === "update-table"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("update", "update-table")}
            >
              <div className={styles.menuItemContent}>
                <Table24Regular className={styles.icon} />
                <span className={styles.label}>更新表格</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "update" && currentTool === "update-toc"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("update", "update-toc")}
            >
              <div className={styles.menuItemContent}>
                <DocumentTable24Regular className={styles.icon} />
                <span className={styles.label}>更新目录</span>
              </div>
            </Button>
          </div>
        )}

        {/* 查询元素类菜单项 */}
        <Button
          appearance="subtle"
          className={`${styles.menuItem} ${isCollapsed ? styles.menuItemCollapsed : ""} ${
            currentPage === "query" ? styles.menuItemActive : styles.menuItemHover
          }`}
          onClick={() => {
            if (isCollapsed) {
              onNavigate("query", "visible-content");
            } else {
              setQueryExpanded(!queryExpanded);
            }
          }}
          title={isCollapsed ? "查询元素类" : ""}
        >
          <div className={styles.menuItemContent}>
            <Search24Regular
              className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`}
            />
            {!isCollapsed && (
              <>
                <span className={styles.label}>查询元素类</span>
                <div
                  className={styles.chevron}
                  style={{ transform: queryExpanded ? "rotate(0deg)" : "rotate(-90deg)" }}
                >
                  <ChevronDown24Regular />
                </div>
              </>
            )}
          </div>
        </Button>

        {/* 查询元素类二级菜单 */}
        {!isCollapsed && (
          <div className={`${styles.submenu} ${!queryExpanded ? styles.submenuCollapsed : ""}`}>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "visible-content"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "visible-content")}
            >
              <div className={styles.menuItemContent}>
                <Eye24Regular className={styles.icon} />
                <span className={styles.label}>可见内容获取</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "document-structure"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "document-structure")}
            >
              <div className={styles.menuItemContent}>
                <DocumentBulletList24Regular className={styles.icon} />
                <span className={styles.label}>文档结构获取</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "document-sections"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "document-sections")}
            >
              <div className={styles.menuItemContent}>
                <DocumentPageBreak24Regular className={styles.icon} />
                <span className={styles.label}>文档节信息获取</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "document-stats"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "document-stats")}
            >
              <div className={styles.menuItemContent}>
                <DocumentData24Regular className={styles.icon} />
                <span className={styles.label}>文档统计信息</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "page-content"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "page-content")}
            >
              <div className={styles.menuItemContent}>
                <DocumentOnePage24Regular className={styles.icon} />
                <span className={styles.label}>页面内容获取</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "selected-content"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "selected-content")}
            >
              <div className={styles.menuItemContent}>
                <TextEffects24Regular className={styles.icon} />
                <span className={styles.label}>选中内容获取</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "header-footer-content"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "header-footer-content")}
            >
              <div className={styles.menuItemContent}>
                <DocumentHeader24Regular className={styles.icon} />
                <span className={styles.label}>页眉页脚内容</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "textbox-content"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "textbox-content")}
            >
              <div className={styles.menuItemContent}>
                <Textbox24Regular className={styles.icon} />
                <span className={styles.label}>文本框内容获取</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "comments"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "comments")}
            >
              <div className={styles.menuItemContent}>
                <Comment24Regular className={styles.icon} />
                <span className={styles.label}>批注内容获取</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "query-table"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "query-table")}
            >
              <div className={styles.menuItemContent}>
                <Table24Regular className={styles.icon} />
                <span className={styles.label}>查询表格</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "query-toc"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "query-toc")}
            >
              <div className={styles.menuItemContent}>
                <DocumentTable24Regular className={styles.icon} />
                <span className={styles.label}>查询目录</span>
              </div>
            </Button>
            <Button
              appearance="subtle"
              className={`${styles.submenuItem} ${
                currentPage === "query" && currentTool === "export-content"
                  ? styles.submenuItemActive
                  : styles.submenuItemHover
              }`}
              onClick={() => onNavigate("query", "export-content")}
            >
              <div className={styles.menuItemContent}>
                <DocumentArrowDown24Regular className={styles.icon} />
                <span className={styles.label}>内容导出</span>
              </div>
            </Button>
          </div>
        )}
      </div>
    </div>
  );
};

export default Sidebar;
