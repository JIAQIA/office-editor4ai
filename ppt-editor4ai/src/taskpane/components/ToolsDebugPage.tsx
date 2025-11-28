/**
 * 文件名: ToolsDebugPage.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 工具调试页面，包含各种工具的调试界面
 */

import * as React from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import { getToolConfig } from "./tools/toolsConfig";

interface ToolsDebugPageProps {
  selectedTool: string;
}

const useStyles = makeStyles({
  container: {
    padding: "16px",
    minHeight: "100vh",
    minWidth: "280px", // 确保内容区有最小宽度
    backgroundColor: tokens.colorNeutralBackground1,
  },
  title: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    marginBottom: "8px",
  },
  subtitle: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground3,
    marginBottom: "20px",
    lineHeight: "1.4",
  },
  toolContainer: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    padding: "16px",
  },
});

const ToolsDebugPage: React.FC<ToolsDebugPageProps> = ({ selectedTool }) => {
  const styles = useStyles();

  // 根据选中的工具渲染不同的组件
  // Render different components based on selected tool
  const renderTool = () => {
    // 获取工具配置
    const toolConfig = getToolConfig(selectedTool);
    
    if (toolConfig) {
      return (
        <>
          <h1 className={styles.title}>
            {toolConfig.title}
          </h1>
          <p className={styles.subtitle}>
            {toolConfig.subtitle}
          </p>
          <div className={styles.toolContainer}>
            {toolConfig.component}
          </div>
        </>
      );
    }
    
    // 默认状态：未选择工具
    return (
      <>
        <h1 className={styles.title}>
          请选择工具
        </h1>
        <p className={styles.subtitle}>
          从左侧菜单选择要调试的工具
        </p>
      </>
    );
  };

  return <div className={styles.container}>{renderTool()}</div>;
};

export default ToolsDebugPage;
