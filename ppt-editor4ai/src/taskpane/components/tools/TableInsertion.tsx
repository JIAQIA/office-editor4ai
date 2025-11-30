/**
 * 文件名: TableInsertion.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 表格插入工具 UI 组件
 */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  Field,
  Input,
  tokens,
  makeStyles,
  Label,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Dropdown,
  Option,
  Switch,
  Textarea,
} from "@fluentui/react-components";
import { insertTableToSlide, TABLE_TEMPLATES } from "../../../ppt-tools";
import { Table24Regular } from "@fluentui/react-icons";

/* global console */

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
interface TableInsertionProps {}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
    padding: "0 8px",
  },
  section: {
    width: "100%",
    marginBottom: "16px",
  },
  row: {
    display: "flex",
    gap: "12px",
    width: "100%",
    marginBottom: "12px",
  },
  field: {
    flex: 1,
  },
  colorInput: {
    width: "100%",
    height: "32px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: "pointer",
    ":hover": {
      border: `1px solid ${tokens.colorNeutralStroke1Hover}`,
    },
    ":focus": {
      outline: `2px solid ${tokens.colorBrandStroke1}`,
      outlineOffset: "1px",
    },
  },
  hint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginBottom: "12px",
    width: "100%",
    textAlign: "center",
    lineHeight: "1.4",
  },
  messageBar: {
    marginBottom: "12px",
    width: "100%",
  },
  templatePreview: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px",
    marginTop: "8px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
  },
  templateIcon: {
    fontSize: "24px",
    color: tokens.colorBrandForeground1,
  },
  templateInfo: {
    flex: 1,
  },
  templateDescription: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
  },
  switchRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    width: "100%",
    marginBottom: "12px",
  },
  dataSection: {
    width: "100%",
    marginBottom: "12px",
  },
  dataHint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
    lineHeight: "1.4",
  },
});

const TableInsertion: React.FC<TableInsertionProps> = () => {
  const styles = useStyles();

  // 模板选择
  const [selectedTemplate, setSelectedTemplate] = useState<string>("custom");
  const [templateName, setTemplateName] = useState<string>("自定义");

  // 表格尺寸
  const [rowCount, setRowCount] = useState<string>("3");
  const [columnCount, setColumnCount] = useState<string>("3");

  // 位置和尺寸
  const [left, setLeft] = useState<string>("");
  const [top, setTop] = useState<string>("");
  const [width, setWidth] = useState<string>("400");
  const [height, setHeight] = useState<string>("");

  // 样式
  const [showHeader, setShowHeader] = useState<boolean>(true);
  const [headerColor, setHeaderColor] = useState<string>("#4472C4");
  const [borderColor, setBorderColor] = useState<string>("#D0D0D0");

  // 数据
  const [useData, setUseData] = useState<boolean>(false);
  const [dataText, setDataText] = useState<string>("");

  // 状态
  const [isInserting, setIsInserting] = useState<boolean>(false);
  const [message, setMessage] = useState<{
    type: "success" | "error" | "warning" | "info";
    title: string;
    content: string;
  } | null>(null);

  // 处理模板选择
  const handleTemplateChange = (
    _event: React.SyntheticEvent,
    data: { optionValue?: string }
  ) => {
    const templateId = data.optionValue as string;
    setSelectedTemplate(templateId);

    if (templateId === "custom") {
      setTemplateName("自定义");
      return;
    }

    const template = TABLE_TEMPLATES.find((t) => t.id === templateId);
    if (template) {
      setTemplateName(template.name);
      setRowCount(template.rowCount.toString());
      setColumnCount(template.columnCount.toString());
    }
  };

  // 解析数据文本
  const parseDataText = (): string[][] | null => {
    if (!useData || !dataText.trim()) {
      return null;
    }

    try {
      // 按行分割
      const lines = dataText.trim().split("\n");
      const result: string[][] = [];

      for (const line of lines) {
        // 支持逗号、制表符或多个空格分隔
        const cells = line
          .split(/[,\t]|\s{2,}/)
          .map((cell) => cell.trim())
          .filter((cell) => cell !== "");

        if (cells.length > 0) {
          result.push(cells);
        }
      }

      return result.length > 0 ? result : null;
    } catch (error) {
      console.error("解析数据失败:", error);
      return null;
    }
  };

  // 处理插入表格
  const handleInsertTable = async () => {
    setIsInserting(true);

    try {
      // 解析参数
      const rowCountValue = parseInt(rowCount, 10);
      const columnCountValue = parseInt(columnCount, 10);
      const leftValue = left.trim() === "" ? undefined : parseFloat(left);
      const topValue = top.trim() === "" ? undefined : parseFloat(top);
      const widthValue = width.trim() === "" ? 400 : parseFloat(width);
      const heightValue = height.trim() === "" ? undefined : parseFloat(height);

      // 验证数值
      if (isNaN(rowCountValue) || rowCountValue <= 0) {
        setMessage({
          type: "warning",
          title: "参数错误",
          content: "行数必须是大于 0 的整数",
        });
        return;
      }

      if (isNaN(columnCountValue) || columnCountValue <= 0) {
        setMessage({
          type: "warning",
          title: "参数错误",
          content: "列数必须是大于 0 的整数",
        });
        return;
      }

      if (rowCountValue > 100) {
        setMessage({
          type: "warning",
          title: "参数错误",
          content: "行数不能超过 100",
        });
        return;
      }

      if (columnCountValue > 50) {
        setMessage({
          type: "warning",
          title: "参数错误",
          content: "列数不能超过 50",
        });
        return;
      }

      if (widthValue <= 0) {
        setMessage({
          type: "warning",
          title: "参数错误",
          content: "宽度必须大于 0",
        });
        return;
      }

      // 解析数据
      const values = parseDataText();

      // 如果提供了数据，验证维度
      if (values) {
        if (values.length !== rowCountValue) {
          setMessage({
            type: "warning",
            title: "数据维度不匹配",
            content: `数据有 ${values.length} 行，但指定了 ${rowCountValue} 行`,
          });
          return;
        }

        const firstRowLength = values[0].length;
        if (firstRowLength !== columnCountValue) {
          setMessage({
            type: "warning",
            title: "数据维度不匹配",
            content: `数据有 ${firstRowLength} 列，但指定了 ${columnCountValue} 列`,
          });
          return;
        }
      }

      // 插入表格
      const result = await insertTableToSlide({
        rowCount: rowCountValue,
        columnCount: columnCountValue,
        left: leftValue,
        top: topValue,
        width: widthValue,
        height: heightValue,
        values: values ?? undefined,
        showHeader,
        headerColor: headerColor.trim() || "#4472C4",
        borderColor: borderColor.trim() || "#D0D0D0",
      });

      setMessage({
        type: "success",
        title: "插入成功",
        content: `表格已插入！${result.rowCount} 行 × ${result.columnCount} 列，位置: (${result.left.toFixed(
          1
        )}, ${result.top.toFixed(1)})`,
      });
    } catch (error) {
      console.error("插入表格失败:", error);
      setMessage({
        type: "error",
        title: "插入失败",
        content: `${(error as Error).message}`,
      });
    } finally {
      setIsInserting(false);
    }
  };

  return (
    <div className={styles.container}>
      {/* 消息提示 */}
      {message && (
        <MessageBar
          key={message.type + message.title}
          intent={message.type}
          className={styles.messageBar}
        >
          <MessageBarBody>
            <MessageBarTitle>{message.title}</MessageBarTitle>
            {message.content}
          </MessageBarBody>
        </MessageBar>
      )}

      {/* 模板选择 */}
      <div className={styles.section}>
        <Field label="选择表格模板">
          <Dropdown
            placeholder="选择模板"
            value={templateName}
            selectedOptions={[selectedTemplate]}
            onOptionSelect={handleTemplateChange}
          >
            <Option value="custom" text="自定义">
              自定义
            </Option>
            {TABLE_TEMPLATES.map((template) => (
              <Option key={template.id} value={template.id} text={template.name}>
                {template.name}
              </Option>
            ))}
          </Dropdown>
        </Field>

        {/* 模板预览 */}
        {selectedTemplate !== "custom" && (
          <div className={styles.templatePreview}>
            <Table24Regular className={styles.templateIcon} />
            <div className={styles.templateInfo}>
              <div>
                <strong>{templateName}</strong>
              </div>
              <div className={styles.templateDescription}>
                {TABLE_TEMPLATES.find((t) => t.id === selectedTemplate)?.description}
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 表格尺寸 */}
      <div className={styles.section}>
        <Label weight="semibold">表格尺寸</Label>
        <div className={styles.row}>
          <Field className={styles.field} label="行数">
            <Input
              type="number"
              value={rowCount}
              onChange={(e) => setRowCount(e.target.value)}
              placeholder="默认 3"
            />
          </Field>
          <Field className={styles.field} label="列数">
            <Input
              type="number"
              value={columnCount}
              onChange={(e) => setColumnCount(e.target.value)}
              placeholder="默认 3"
            />
          </Field>
        </div>
      </div>

      {/* 位置和尺寸 */}
      <div className={styles.section}>
        <Label weight="semibold">位置和尺寸</Label>
        <div className={styles.row}>
          <Field className={styles.field} label="X 坐标">
            <Input
              type="number"
              value={left}
              onChange={(e) => setLeft(e.target.value)}
              placeholder="留空居中"
            />
          </Field>
          <Field className={styles.field} label="Y 坐标">
            <Input
              type="number"
              value={top}
              onChange={(e) => setTop(e.target.value)}
              placeholder="留空居中"
            />
          </Field>
        </div>
        <div className={styles.row}>
          <Field className={styles.field} label="宽度（磅）">
            <Input
              type="number"
              value={width}
              onChange={(e) => setWidth(e.target.value)}
              placeholder="默认 400"
            />
          </Field>
          <Field className={styles.field} label="高度（磅）">
            <Input
              type="number"
              value={height}
              onChange={(e) => setHeight(e.target.value)}
              placeholder="自动计算"
            />
          </Field>
        </div>
      </div>

      {/* 样式设置 */}
      <div className={styles.section}>
        <Label weight="semibold">样式设置</Label>
        <div className={styles.switchRow}>
          <Label>显示表头样式</Label>
          <Switch checked={showHeader} onChange={(e) => setShowHeader(e.currentTarget.checked)} />
        </div>
        <div className={styles.row}>
          <Field className={styles.field} label="表头颜色">
            <input
              type="color"
              className={styles.colorInput}
              value={headerColor}
              onChange={(e) => setHeaderColor(e.target.value)}
              disabled={!showHeader}
            />
          </Field>
          <Field className={styles.field} label="边框颜色">
            <input
              type="color"
              className={styles.colorInput}
              value={borderColor}
              onChange={(e) => setBorderColor(e.target.value)}
            />
          </Field>
        </div>
      </div>

      {/* 数据输入 */}
      <div className={styles.section}>
        <div className={styles.switchRow}>
          <Label weight="semibold">填充表格数据</Label>
          <Switch checked={useData} onChange={(e) => setUseData(e.currentTarget.checked)} />
        </div>
        {useData && (
          <div className={styles.dataSection}>
            <Field label="表格数据（每行一行，用逗号或制表符分隔）">
              <Textarea
                value={dataText}
                onChange={(e) => setDataText(e.target.value)}
                placeholder="例如：&#10;姓名,年龄,城市&#10;张三,25,北京&#10;李四,30,上海"
                rows={6}
              />
            </Field>
            <div className={styles.dataHint}>
              💡 提示：每行代表表格的一行，单元格之间用逗号、制表符或多个空格分隔
            </div>
          </div>
        )}
      </div>

      <div className={styles.hint}>
        💡 位置范围提示: <br />
        标准 16:9 幻灯片尺寸约为 720×540 磅 (points)
        <br />X 范围: 0-720, Y 范围: 0-540
      </div>

      {/* 操作按钮 */}
      <div className={styles.section}>
        <Button
          appearance="primary"
          size="large"
          onClick={handleInsertTable}
          disabled={isInserting}
        >
          {isInserting ? "插入中..." : "确认插入"}
        </Button>
      </div>
    </div>
  );
};

export default TableInsertion;
