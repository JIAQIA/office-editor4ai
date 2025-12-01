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
import TextInsertion from "./TextInsertion";
import VisibleContent from "./tools/VisibleContent";
import DocumentStructure from "./tools/DocumentStructure";
import DocumentSections from "./tools/DocumentSections";
import { insertText } from "../taskpane";

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
    switch (selectedTool) {
      case "text-insertion":
        return (
          <>
            <h1 className={styles.title}>
              文本插入工具
            </h1>
            <p className={styles.subtitle}>
              在文档中插入文本段落
            </p>
            <div className={styles.toolContainer}>
              <TextInsertion insertText={insertText} />
            </div>
          </>
        );
      case "visible-content":
        return (
          <>
            <h1 className={styles.title}>
              可见内容获取工具
            </h1>
            <p className={styles.subtitle}>
              获取用户当前可见范围的文档内容，包括段落、表格、图片等元素
            </p>
            <div className={styles.toolContainer}>
              <VisibleContent />
            </div>
          </>
        );
      case "document-structure":
        return (
          <>
            <h1 className={styles.title}>
              文档结构获取工具
            </h1>
            <p className={styles.subtitle}>
              获取文档的大纲结构（标题层级），支持树形展示和导出功能
            </p>
            <div className={styles.toolContainer}>
              <DocumentStructure />
            </div>
          </>
        );
      case "document-sections":
        return (
          <>
            <h1 className={styles.title}>
              文档节信息获取工具
            </h1>
            <p className={styles.subtitle}>
              获取文档的分节符、页眉页脚配置、页面设置等信息
            </p>
            <div className={styles.toolContainer}>
              <DocumentSections />
            </div>
          </>
        );
      default:
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
    }
  };

  return <div className={styles.container}>{renderTool()}</div>;
};

export default ToolsDebugPage;
