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
import { DocumentStats } from "./tools/DocumentStats";
import PageContent from "./tools/PageContent";
import SelectedContent from "./tools/SelectedContent";
import HeaderFooterContent from "./tools/HeaderFooterContent";
import TextBoxContent from "./tools/TextBoxContent";
import Comments from "./tools/Comments";
import { ReplaceSelection } from "./tools/ReplaceSelection";
import { ReplaceTextDebug } from "./tools/ReplaceTextDebug";
import { ReplaceImageDebug } from "./tools/ReplaceImageDebug";
import { insertText } from "../taskpane";
import { appendText } from "../../word-tools";
import { AppendTextDebug } from "./tools/AppendTextDebug";
import { InsertImageDebug } from "./tools/InsertImageDebug";
import { CreateTableDebug } from "./tools/CreateTableDebug";
import { DeleteTableDebug } from "./tools/DeleteTableDebug";
import { UpdateTableDebug } from "./tools/UpdateTableDebug";
import { QueryTableDebug } from "./tools/QueryTableDebug";
import { InsertTextBoxDebug } from "./tools/InsertTextBoxDebug";
import { InsertShapeDebug } from "./tools/InsertShapeDebug";
import { InsertPageBreakDebug } from "./tools/InsertPageBreakDebug";
import { InsertSectionBreakDebug } from "./tools/InsertSectionBreakDebug";
import { TableOfContentsDebug } from "./tools/TableOfContentsDebug";
import { InsertTableOfContents } from "./tools/InsertTableOfContents";
import { DeleteTableOfContents } from "./tools/DeleteTableOfContents";
import { UpdateTableOfContents } from "./tools/UpdateTableOfContents";
import { QueryTableOfContents } from "./tools/QueryTableOfContents";
import { InsertEquationDebug } from "./tools/InsertEquationDebug";

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
      case "document-stats":
        return (
          <>
            <h1 className={styles.title}>
              文档统计信息工具
            </h1>
            <p className={styles.subtitle}>
              获取文档的字数、段落数、页数等统计信息，支持自定义统计选项
            </p>
            <div className={styles.toolContainer}>
              <DocumentStats />
            </div>
          </>
        );
      case "page-content":
        return (
          <>
            <h1 className={styles.title}>
              页面内容获取工具
            </h1>
            <p className={styles.subtitle}>
              获取指定页面的内容和统计信息，支持自定义获取选项
            </p>
            <div className={styles.toolContainer}>
              <PageContent />
            </div>
          </>
        );
      case "selected-content":
        return (
          <>
            <h1 className={styles.title}>
              选中内容获取工具
            </h1>
            <p className={styles.subtitle}>
              获取用户当前选中的内容，包括文本、段落、表格、图片等元素
            </p>
            <div className={styles.toolContainer}>
              <SelectedContent />
            </div>
          </>
        );
      case "header-footer-content":
        return (
          <>
            <h1 className={styles.title}>
              页眉页脚内容获取工具
            </h1>
            <p className={styles.subtitle}>
              获取文档所有节的页眉页脚内容，支持按节索引查询和详细元素解析
            </p>
            <div className={styles.toolContainer}>
              <HeaderFooterContent />
            </div>
          </>
        );
      case "textbox-content":
        return (
          <>
            <h1 className={styles.title}>
              文本框内容获取工具
            </h1>
            <p className={styles.subtitle}>
              获取文档中的文本框内容，支持选择范围或可见区域，可获取文本框的文本、段落和元数据信息
            </p>
            <div className={styles.toolContainer}>
              <TextBoxContent />
            </div>
          </>
        );
      case "comments":
        return (
          <>
            <h1 className={styles.title}>
              批注内容获取工具
            </h1>
            <p className={styles.subtitle}>
              获取文档中的批注内容，支持选择范围或整个文档，可获取批注的内容、作者、回复和关联文本等信息
            </p>
            <div className={styles.toolContainer}>
              <Comments />
            </div>
          </>
        );
      case "replace-selection":
        return (
          <>
            <h1 className={styles.title}>
              替换选中内容工具
            </h1>
            <p className={styles.subtitle}>
              替换或插入文本到选中位置，支持自定义格式（字体、字号、颜色等）或保持原格式
            </p>
            <div className={styles.toolContainer}>
              <ReplaceSelection />
            </div>
          </>
        );
      case "replace-text":
        return (
          <>
            <h1 className={styles.title}>
              替换文本工具
            </h1>
            <p className={styles.subtitle}>
              统一的文本替换工具，支持三种定位方式：当前选区、搜索匹配、指定范围（书签、标题、段落、节、内容控件）
            </p>
            <div className={styles.toolContainer}>
              <ReplaceTextDebug />
            </div>
          </>
        );
      case "replace-image":
        return (
          <>
            <h1 className={styles.title}>
              替换图片工具
            </h1>
            <p className={styles.subtitle}>
              统一的图片替换工具，支持四种定位方式：当前选区、按索引、搜索匹配、指定范围。可替换图片内容或更新图片属性
            </p>
            <div className={styles.toolContainer}>
              <ReplaceImageDebug />
            </div>
          </>
        );
      case "append-text":
        return (
          <>
            <h1 className={styles.title}>
              文档末尾追加文本工具
            </h1>
            <p className={styles.subtitle}>
              在文档末尾追加文本或图片，支持自定义文本格式
            </p>
            <div className={styles.toolContainer}>
              <AppendTextDebug 
                appendText={appendText}
              />
            </div>
          </>
        );
      case "insert-image":
        return (
          <>
            <h1 className={styles.title}>
              插入图片工具
            </h1>
            <p className={styles.subtitle}>
              在文档中插入图片，支持内联和浮动布局，可配置尺寸、位置、文字环绕等选项
            </p>
            <div className={styles.toolContainer}>
              <InsertImageDebug />
            </div>
          </>
        );
      case "create-table":
        return (
          <>
            <h1 className={styles.title}>
              创建表格工具
            </h1>
            <p className={styles.subtitle}>
              在文档中插入表格，支持自定义行列数、样式、边框、对齐方式等选项
            </p>
            <div className={styles.toolContainer}>
              <CreateTableDebug />
            </div>
          </>
        );
      case "delete-table":
        return (
          <>
            <h1 className={styles.title}>
              删除表格工具
            </h1>
            <p className={styles.subtitle}>
              删除文档中的表格，支持按索引删除或删除选中的表格
            </p>
            <div className={styles.toolContainer}>
              <DeleteTableDebug />
            </div>
          </>
        );
      case "update-table":
        return (
          <>
            <h1 className={styles.title}>
              更新表格工具
            </h1>
            <p className={styles.subtitle}>
              更新表格内容和格式，支持更新整个表格、单个单元格、行列操作和单元格合并
            </p>
            <div className={styles.toolContainer}>
              <UpdateTableDebug />
            </div>
          </>
        );
      case "query-table":
        return (
          <>
            <h1 className={styles.title}>
              查询表格工具
            </h1>
            <p className={styles.subtitle}>
              查询文档中的表格信息，支持查询单个表格或所有表格的详细信息
            </p>
            <div className={styles.toolContainer}>
              <QueryTableDebug />
            </div>
          </>
        );
      case "insert-textbox":
        return (
          <>
            <h1 className={styles.title}>
              插入文本框工具
            </h1>
            <p className={styles.subtitle}>
              在文档中插入文本框，支持自定义尺寸、位置、旋转角度和文本格式
            </p>
            <div className={styles.toolContainer}>
              <InsertTextBoxDebug />
            </div>
          </>
        );
      case "insert-shape":
        return (
          <>
            <h1 className={styles.title}>
              插入形状工具
            </h1>
            <p className={styles.subtitle}>
              在文档中插入各种形状，支持自定义样式、尺寸和位置
            </p>
            <div className={styles.toolContainer}>
              <InsertShapeDebug />
            </div>
          </>
        );
      case "insert-page-break":
        return (
          <>
            <h1 className={styles.title}>
              插入分页符工具
            </h1>
            <p className={styles.subtitle}>
              在文档中插入分页符，强制在指定位置开始新的一页
            </p>
            <div className={styles.toolContainer}>
              <InsertPageBreakDebug />
            </div>
          </>
        );
      case "insert-section-break":
        return (
          <>
            <h1 className={styles.title}>
              插入分节符工具
            </h1>
            <p className={styles.subtitle}>
              在文档中插入分节符，将文档分成不同的节，每节可以有独立的页面设置
            </p>
            <div className={styles.toolContainer}>
              <InsertSectionBreakDebug />
            </div>
          </>
        );
      case "insert-toc":
        return (
          <>
            <h1 className={styles.title}>
              插入目录工具
            </h1>
            <p className={styles.subtitle}>
              在文档中插入目录，支持自定义标题级别、页码显示等选项
            </p>
            <div className={styles.toolContainer}>
              <InsertTableOfContents />
            </div>
          </>
        );
      case "delete-toc":
        return (
          <>
            <h1 className={styles.title}>
              删除目录工具
            </h1>
            <p className={styles.subtitle}>
              删除文档中的目录，支持按索引删除或删除所有目录
            </p>
            <div className={styles.toolContainer}>
              <DeleteTableOfContents />
            </div>
          </>
        );
      case "update-toc":
        return (
          <>
            <h1 className={styles.title}>
              更新目录工具
            </h1>
            <p className={styles.subtitle}>
              更新文档中的目录，支持按索引更新或更新所有目录
            </p>
            <div className={styles.toolContainer}>
              <UpdateTableOfContents />
            </div>
          </>
        );
      case "query-toc":
        return (
          <>
            <h1 className={styles.title}>
              查询目录工具
            </h1>
            <p className={styles.subtitle}>
              查询文档中的目录信息，获取所有目录的详细信息
            </p>
            <div className={styles.toolContainer}>
              <QueryTableOfContents />
            </div>
          </>
        );
      case "insert-equation":
        return (
          <>
            <h1 className={styles.title}>
              插入公式工具
            </h1>
            <p className={styles.subtitle}>
              在文档中插入数学公式，支持 LaTeX 格式，包括分数、根号、求和、积分等常见数学符号
            </p>
            <div className={styles.toolContainer}>
              <InsertEquationDebug />
            </div>
          </>
        );
      case "table-of-contents":
        return (
          <>
            <h1 className={styles.title}>
              目录管理工具（已废弃）
            </h1>
            <p className={styles.subtitle}>
              此工具已拆分为插入、删除、更新、查询四个独立工具，请从左侧菜单选择对应的工具
            </p>
            <div className={styles.toolContainer}>
              <TableOfContentsDebug />
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
