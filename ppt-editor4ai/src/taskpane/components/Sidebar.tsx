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
  List24Regular,
  LayoutColumnTwoSplitLeft24Regular,
  LayoutRowTwo24Regular,
  Image24Regular,
  Video24Regular,
  Camera24Regular,
  Shapes24Regular,
  Table24Regular,
  DocumentOnePage24Regular
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
    boxSizing: "border-box",
  },
  sidebarCollapsed: {
    width: "56px",
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
    padding: "10px 8px",
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
  onToggleCollapse 
}) => {
  const styles = useStyles();
  const [createExpanded, setCreateExpanded] = useState(true);
  const [deleteExpanded, setDeleteExpanded] = useState(false);
  const [updateExpanded, setUpdateExpanded] = useState(false);
  const [queryExpanded, setQueryExpanded] = useState(false);

  return (
    <div className={`${styles.sidebar} ${isCollapsed ? styles.sidebarCollapsed : ""}`}>
      {/* 折叠/展开按钮 / Toggle button */}
      <div className={`${styles.toggleButtonWrapper} ${isCollapsed ? styles.toggleButtonWrapperCollapsed : ""}`}>
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
      <div className={`${styles.scrollContainer} ${isCollapsed ? styles.scrollContainerCollapsed : ""}`}>
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
          <Home24Regular className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`} />
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
                style={{ transform: createExpanded ? 'rotate(0deg)' : 'rotate(-90deg)' }}
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
              currentPage === "create" && currentTool === "image-insertion"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("create", "image-insertion")}
          >
            <div className={styles.menuItemContent}>
              <Image24Regular className={styles.icon} />
              <span className={styles.label}>图片插入工具</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "create" && currentTool === "video-insertion"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("create", "video-insertion")}
          >
            <div className={styles.menuItemContent}>
              <Video24Regular className={styles.icon} />
              <span className={styles.label}>视频插入工具</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "create" && currentTool === "shape-insertion"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("create", "shape-insertion")}
          >
            <div className={styles.menuItemContent}>
              <Shapes24Regular className={styles.icon} />
              <span className={styles.label}>形状插入工具</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "create" && currentTool === "table-insertion"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("create", "table-insertion")}
          >
            <div className={styles.menuItemContent}>
              <Table24Regular className={styles.icon} />
              <span className={styles.label}>表格插入工具</span>
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
            onNavigate("delete", "element-deletion");
          } else {
            setDeleteExpanded(!deleteExpanded);
          }
        }}
        title={isCollapsed ? "删除元素类" : ""}
      >
        <div className={styles.menuItemContent}>
          <Delete24Regular className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`} />
          {!isCollapsed && (
            <>
              <span className={styles.label}>删除元素类</span>
              <div 
                className={styles.chevron}
                style={{ transform: deleteExpanded ? 'rotate(0deg)' : 'rotate(-90deg)' }}
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
              currentPage === "delete" && currentTool === "element-deletion"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("delete", "element-deletion")}
          >
            <div className={styles.menuItemContent}>
              <Delete24Regular className={styles.icon} />
              <span className={styles.label}>元素删除工具</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "delete" && currentTool === "slide-deletion"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("delete", "slide-deletion")}
          >
            <div className={styles.menuItemContent}>
              <DocumentOnePage24Regular className={styles.icon} />
              <span className={styles.label}>幻灯片删除工具</span>
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
            onNavigate("update", "text-update");
          } else {
            setUpdateExpanded(!updateExpanded);
          }
        }}
        title={isCollapsed ? "修改元素类" : ""}
      >
        <div className={styles.menuItemContent}>
          <Edit24Regular className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`} />
          {!isCollapsed && (
            <>
              <span className={styles.label}>修改元素类</span>
              <div 
                className={styles.chevron}
                style={{ transform: updateExpanded ? 'rotate(0deg)' : 'rotate(-90deg)' }}
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
              currentPage === "update" && currentTool === "text-update"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("update", "text-update")}
          >
            <div className={styles.menuItemContent}>
              <TextGrammarSettings24Regular className={styles.icon} />
              <span className={styles.label}>文本框更新工具</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "update" && currentTool === "slide-move"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("update", "slide-move")}
          >
            <div className={styles.menuItemContent}>
              <DocumentOnePage24Regular className={styles.icon} />
              <span className={styles.label}>幻灯片移动工具</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "update" && currentTool === "image-replace"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("update", "image-replace")}
          >
            <div className={styles.menuItemContent}>
              <Image24Regular className={styles.icon} />
              <span className={styles.label}>图片替换工具</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "update" && currentTool === "table-cell-update"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("update", "table-cell-update")}
          >
            <div className={styles.menuItemContent}>
              <Table24Regular className={styles.icon} />
              <span className={styles.label}>表格单元格更新</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "update" && currentTool === "table-row-column-update"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("update", "table-row-column-update")}
          >
            <div className={styles.menuItemContent}>
              <Table24Regular className={styles.icon} />
              <span className={styles.label}>表格行/列更新</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "update" && currentTool === "table-format-update"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("update", "table-format-update")}
          >
            <div className={styles.menuItemContent}>
              <Table24Regular className={styles.icon} />
              <span className={styles.label}>表格格式更新</span>
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
            onNavigate("query", "elements-list");
          } else {
            setQueryExpanded(!queryExpanded);
          }
        }}
        title={isCollapsed ? "查询元素类" : ""}
      >
        <div className={styles.menuItemContent}>
          <Search24Regular className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`} />
          {!isCollapsed && (
            <>
              <span className={styles.label}>查询元素类</span>
              <div 
                className={styles.chevron}
                style={{ transform: queryExpanded ? 'rotate(0deg)' : 'rotate(-90deg)' }}
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
              currentPage === "query" && currentTool === "elements-list"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("query", "elements-list")}
          >
            <div className={styles.menuItemContent}>
              <List24Regular className={styles.icon} />
              <span className={styles.label}>元素列表</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "query" && currentTool === "slide-layout-info"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("query", "slide-layout-info")}
          >
            <div className={styles.menuItemContent}>
              <LayoutColumnTwoSplitLeft24Regular className={styles.icon} />
              <span className={styles.label}>页面布局信息</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "query" && currentTool === "slide-layouts"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("query", "slide-layouts")}
          >
            <div className={styles.menuItemContent}>
              <LayoutRowTwo24Regular className={styles.icon} />
              <span className={styles.label}>布局模板列表</span>
            </div>
          </Button>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "query" && currentTool === "slide-screenshot"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("query", "slide-screenshot")}
          >
            <div className={styles.menuItemContent}>
              <Camera24Regular className={styles.icon} />
              <span className={styles.label}>幻灯片截图</span>
            </div>
          </Button>
        </div>
      )}
      </div>
    </div>
  );
};

export default Sidebar;
