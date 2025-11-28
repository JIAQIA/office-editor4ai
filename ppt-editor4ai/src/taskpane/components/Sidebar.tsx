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
  Wrench24Regular, 
  TextGrammarSettings24Regular,
  ChevronDown24Regular,
  ChevronRight24Regular,
  Navigation24Regular,
  Dismiss24Regular
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
    padding: "8px 0",
    transition: "width 0.3s ease",
    flexShrink: 0,
  },
  sidebarCollapsed: {
    width: "48px",
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
    minHeight: "40px",
  },
  menuItemCollapsed: {
    padding: "10px 4px",
    justifyContent: "center",
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
  },
  icon: {
    marginRight: "10px",
    fontSize: "20px",
  },
  iconCollapsed: {
    marginRight: "0",
    fontSize: "22px",
  },
  label: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    flex: 1,
  },
  chevron: {
    fontSize: "16px",
    marginLeft: "auto",
  },
  submenu: {
    paddingLeft: "12px",
  },
  submenuItem: {
    width: "100%",
    justifyContent: "flex-start",
    padding: "8px 12px 8px 38px",
    border: "none",
    backgroundColor: "transparent",
    cursor: "pointer",
    transition: "background-color 0.2s",
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
  const [toolsExpanded, setToolsExpanded] = useState(true);

  return (
    <div className={`${styles.sidebar} ${isCollapsed ? styles.sidebarCollapsed : ""}`}>
      {/* 折叠/展开按钮 / Toggle button */}
      <Button
        appearance="subtle"
        className={`${styles.toggleButton} ${isCollapsed ? styles.toggleButtonCollapsed : ""}`}
        onClick={onToggleCollapse}
        title={isCollapsed ? "展开侧边栏" : "折叠侧边栏"}
      >
        {isCollapsed ? <Navigation24Regular /> : <Dismiss24Regular />}
      </Button>

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

      {/* 工具调试页菜单项 */}
      <Button
        appearance="subtle"
        className={`${styles.menuItem} ${isCollapsed ? styles.menuItemCollapsed : ""} ${
          currentPage === "tools" ? styles.menuItemActive : styles.menuItemHover
        }`}
        onClick={() => {
          if (isCollapsed) {
            onNavigate("tools", "text-insertion");
          } else {
            setToolsExpanded(!toolsExpanded);
          }
        }}
        title={isCollapsed ? "工具调试页" : ""}
      >
        <div className={styles.menuItemContent}>
          <Wrench24Regular className={`${styles.icon} ${isCollapsed ? styles.iconCollapsed : ""}`} />
          {!isCollapsed && (
            <>
              <span className={styles.label}>工具调试页</span>
              {toolsExpanded ? (
                <ChevronDown24Regular className={styles.chevron} />
              ) : (
                <ChevronRight24Regular className={styles.chevron} />
              )}
            </>
          )}
        </div>
      </Button>

      {/* 工具调试页二级菜单 */}
      {!isCollapsed && toolsExpanded && (
        <div className={styles.submenu}>
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "tools" && currentTool === "text-insertion"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("tools", "text-insertion")}
          >
            <div className={styles.menuItemContent}>
              <TextGrammarSettings24Regular className={styles.icon} />
              <span className={styles.label}>文本插入工具</span>
            </div>
          </Button>
        </div>
      )}
    </div>
  );
};

export default Sidebar;
