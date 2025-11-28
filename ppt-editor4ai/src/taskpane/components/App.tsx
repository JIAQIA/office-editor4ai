/**
 * 文件名: App.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 主应用组件，包含侧边栏导航和页面路由
 */

import * as React from "react";
import { useState } from "react";
import { makeStyles } from "@fluentui/react-components";
import Sidebar from "./Sidebar";
import HomePage from "./HomePage";
import ToolsDebugPage from "./ToolsDebugPage";

const useStyles = makeStyles({
  root: {
    display: "flex",
    minHeight: "100vh",
    width: "100%",
    overflow: "hidden",
  },
  content: {
    minWidth: "600px",
    width: "600px",
    overflow: "auto",
  },
});

const App: React.FC = () => {
  const styles = useStyles();
  const [currentPage, setCurrentPage] = useState<string>("home");
  const [currentTool, setCurrentTool] = useState<string>("");
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(false);

  // 处理导航 / Handle navigation
  const handleNavigate = (page: string, tool?: string) => {
    setCurrentPage(page);
    if (tool) {
      setCurrentTool(tool);
    }
  };

  // 切换侧边栏折叠状态 / Toggle sidebar collapse state
  const handleToggleSidebar = () => {
    setIsSidebarCollapsed(!isSidebarCollapsed);
  };

  // 根据当前页面渲染内容 / Render content based on current page
  const renderContent = () => {
    switch (currentPage) {
      case "home":
        return <HomePage />;
      case "tools":
        return <ToolsDebugPage selectedTool={currentTool} />;
      default:
        return <HomePage />;
    }
  };

  return (
    <div className={styles.root}>
      <Sidebar 
        currentPage={currentPage} 
        currentTool={currentTool} 
        isCollapsed={isSidebarCollapsed}
        onNavigate={handleNavigate}
        onToggleCollapse={handleToggleSidebar}
      />
      <div className={styles.content}>{renderContent()}</div>
    </div>
  );
};

export default App;
