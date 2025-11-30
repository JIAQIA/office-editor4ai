/**
 * 文件名: ShapeInsertion.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 形状插入工具 UI 组件
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
} from "@fluentui/react-components";
import { insertShapeToSlide, COMMON_SHAPES, ShapeType } from "../../../ppt-tools";
import { Shapes24Regular } from "@fluentui/react-icons";

/* global console */

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
interface ShapeInsertionProps {}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
    padding: "0 8px",
  },
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "16px",
    marginBottom: "8px",
    fontSize: tokens.fontSizeBase300,
  },
  section: {
    width: "100%",
    marginBottom: "16px",
  },
  positionContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    width: "100%",
    marginBottom: "12px",
  },
  positionRow: {
    display: "flex",
    gap: "12px",
    width: "100%",
  },
  positionField: {
    flex: 1,
  },
  colorRow: {
    display: "flex",
    gap: "12px",
    width: "100%",
    marginBottom: "12px",
  },
  colorField: {
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
  shapePreview: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px",
    marginTop: "8px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
  },
  shapeIcon: {
    fontSize: "24px",
    color: tokens.colorBrandForeground1,
  },
  categoryLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
  },
});

const ShapeInsertion: React.FC<ShapeInsertionProps> = () => {
  const styles = useStyles();

  // 形状类型
  const [shapeType, setShapeType] = useState<ShapeType>("rectangle");
  const [selectedShapeLabel, setSelectedShapeLabel] = useState<string>("矩形");
  const [selectedShapeCategory, setSelectedShapeCategory] = useState<string>("基础形状");

  // 位置和尺寸
  const [left, setLeft] = useState<string>("");
  const [top, setTop] = useState<string>("");
  const [width, setWidth] = useState<string>("100");
  const [height, setHeight] = useState<string>("100");

  // 样式
  const [fillColor, setFillColor] = useState<string>("#4472C4");
  const [lineColor, setLineColor] = useState<string>("#2E5090");
  const [lineWeight, setLineWeight] = useState<string>("2");

  // 文本
  const [text, setText] = useState<string>("");

  // 状态
  const [isInserting, setIsInserting] = useState<boolean>(false);
  const [message, setMessage] = useState<{
    type: "success" | "error" | "warning" | "info";
    title: string;
    content: string;
  } | null>(null);

  // 按分类组织形状
  const shapesByCategory = COMMON_SHAPES.reduce(
    (acc, shape) => {
      if (!acc[shape.category]) {
        acc[shape.category] = [];
      }
      acc[shape.category].push(shape);
      return acc;
    },
    {} as Record<string, typeof COMMON_SHAPES>
  );

  // 处理形状选择
  const handleShapeChange = (
    _event: React.SyntheticEvent,
    data: { optionValue?: string }
  ) => {
    const selectedType = data.optionValue as ShapeType;
    const selectedShape = COMMON_SHAPES.find((s) => s.type === selectedType);
    if (selectedShape) {
      setShapeType(selectedType);
      setSelectedShapeLabel(selectedShape.label);
      setSelectedShapeCategory(selectedShape.category);
    }
  };

  // 处理插入形状
  const handleInsertShape = async () => {
    setIsInserting(true);

    try {
      // 解析位置和尺寸参数
      const leftValue = left.trim() === "" ? undefined : parseFloat(left);
      const topValue = top.trim() === "" ? undefined : parseFloat(top);
      const widthValue = width.trim() === "" ? 100 : parseFloat(width);
      const heightValue = height.trim() === "" ? 100 : parseFloat(height);
      const lineWeightValue = lineWeight.trim() === "" ? 2 : parseFloat(lineWeight);

      // 验证数值
      if (widthValue <= 0 || heightValue <= 0) {
        setMessage({
          type: "warning",
          title: "参数错误",
          content: "宽度和高度必须大于 0",
        });
        return;
      }

      if (lineWeightValue < 0) {
        setMessage({
          type: "warning",
          title: "参数错误",
          content: "边框粗细不能为负数",
        });
        return;
      }

      // 插入形状
      const result = await insertShapeToSlide({
        shapeType,
        left: leftValue,
        top: topValue,
        width: widthValue,
        height: heightValue,
        fillColor: fillColor.trim() || "#4472C4",
        lineColor: lineColor.trim() || "#2E5090",
        lineWeight: lineWeightValue,
        text: text.trim() || undefined,
      });

      setMessage({
        type: "success",
        title: "插入成功",
        content: `形状已插入！位置: (${result.left.toFixed(1)}, ${result.top.toFixed(
          1
        )})，尺寸: ${result.width.toFixed(1)} × ${result.height.toFixed(1)} 磅`,
      });
    } catch (error) {
      console.error("插入形状失败:", error);
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

      {/* 形状类型选择 */}
      <div className={styles.section}>
        <Field label="选择形状类型">
          <Dropdown
            placeholder="选择形状"
            value={selectedShapeLabel}
            selectedOptions={[shapeType]}
            onOptionSelect={handleShapeChange}
          >
            {Object.entries(shapesByCategory).map(([category, shapes]) => (
              <React.Fragment key={category}>
                <Option text={category} disabled>
                  {category}
                </Option>
                {shapes.map((shape) => (
                  <Option key={shape.type} value={shape.type} text={shape.label}>
                    {shape.label}
                  </Option>
                ))}
              </React.Fragment>
            ))}
          </Dropdown>
        </Field>

        {/* 形状预览信息 */}
        <div className={styles.shapePreview}>
          <Shapes24Regular className={styles.shapeIcon} />
          <div>
            <div>
              <strong>{selectedShapeLabel}</strong>
            </div>
            <div className={styles.categoryLabel}>{selectedShapeCategory}</div>
          </div>
        </div>
      </div>

      {/* 位置和尺寸设置 */}
      <div className={styles.section}>
        <Label weight="semibold">位置和尺寸</Label>
        <div className={styles.positionContainer}>
          <div className={styles.positionRow}>
            <Field className={styles.positionField} label="X 坐标">
              <Input
                type="number"
                value={left}
                onChange={(e) => setLeft(e.target.value)}
                placeholder="留空居中"
              />
            </Field>
            <Field className={styles.positionField} label="Y 坐标">
              <Input
                type="number"
                value={top}
                onChange={(e) => setTop(e.target.value)}
                placeholder="留空居中"
              />
            </Field>
          </div>
          <div className={styles.positionRow}>
            <Field className={styles.positionField} label="宽度">
              <Input
                type="number"
                value={width}
                onChange={(e) => setWidth(e.target.value)}
                placeholder="默认 100"
              />
            </Field>
            <Field className={styles.positionField} label="高度">
              <Input
                type="number"
                value={height}
                onChange={(e) => setHeight(e.target.value)}
                placeholder="默认 100"
              />
            </Field>
          </div>
        </div>
      </div>

      {/* 样式设置 */}
      <div className={styles.section}>
        <Label weight="semibold">样式设置</Label>
        <div className={styles.positionContainer}>
          <div className={styles.colorRow}>
            <Field className={styles.colorField} label="填充颜色">
              <input
                type="color"
                className={styles.colorInput}
                value={fillColor}
                onChange={(e) => setFillColor(e.target.value)}
              />
            </Field>
            <Field className={styles.colorField} label="边框颜色">
              <input
                type="color"
                className={styles.colorInput}
                value={lineColor}
                onChange={(e) => setLineColor(e.target.value)}
              />
            </Field>
          </div>
          <Field label="边框粗细（磅）">
            <Input
              type="number"
              value={lineWeight}
              onChange={(e) => setLineWeight(e.target.value)}
              placeholder="默认 2"
            />
          </Field>
        </div>
      </div>

      {/* 文本设置 */}
      <div className={styles.section}>
        <Field label="形状内文本（可选）">
          <Input
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder="输入文本内容"
          />
        </Field>
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
          onClick={handleInsertShape}
          disabled={isInserting}
        >
          {isInserting ? "插入中..." : "确认插入"}
        </Button>
      </div>
    </div>
  );
};

export default ShapeInsertion;
