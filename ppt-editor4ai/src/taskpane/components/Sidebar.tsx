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
  Navigation24Regular,
  Dismiss24Regular,
  List24Regular
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
              <div 
                className={styles.chevron}
                style={{ transform: toolsExpanded ? 'rotate(0deg)' : 'rotate(-90deg)' }}
              >
                <ChevronDown24Regular />
              </div>
            </>
          )}
        </div>
      </Button>

      {/* 工具调试页二级菜单 */}
      {!isCollapsed && (
        <div className={`${styles.submenu} ${!toolsExpanded ? styles.submenuCollapsed : ""}`}>
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
          
          <Button
            appearance="subtle"
            className={`${styles.submenuItem} ${
              currentPage === "tools" && currentTool === "elements-list"
                ? styles.submenuItemActive
                : styles.submenuItemHover
            }`}
            onClick={() => onNavigate("tools", "elements-list")}
          >
            <div className={styles.menuItemContent}>
              <List24Regular className={styles.icon} />
              <span className={styles.label}>元素列表</span>
            </div>
          </Button>
        </div>
      )}
    </div>
  );
};

export default Sidebar;
