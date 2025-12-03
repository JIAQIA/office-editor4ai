/**
 * 文件名: InsertShapeDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: insertShape工具的调试组件
 */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  Input,
  Label,
  makeStyles,
  Spinner,
  Switch,
  Dropdown,
  Option,
  tokens,
  Field,
} from "@fluentui/react-components";
import { Shapes24Regular } from "@fluentui/react-icons";
import {
  insertShape,
  type ShapeOptions,
  type InsertLocation,
  type WordShapeType,
} from "../../../word-tools";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  formRow: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  formGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "12px",
  },
  optionsContainer: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "8px",
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    marginTop: "12px",
  },
  resultMessage: {
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "12px",
  },
  success: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
  },
  error: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
  },
  warningBox: {
    padding: "8px",
    backgroundColor: tokens.colorPaletteYellowBackground1,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: "12px",
    marginTop: "8px",
  },
  colorInputGroup: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  colorPicker: {
    width: "40px",
    height: "32px",
    border: "none",
    cursor: "pointer",
    borderRadius: tokens.borderRadiusSmall,
  },
});

// 常用形状类型列表 / Common shape types list
const COMMON_SHAPES: Array<{ value: WordShapeType; label: string }> = [
  { value: "Rectangle", label: "矩形" },
  { value: "RoundRectangle", label: "圆角矩形" },
  { value: "Ellipse", label: "椭圆" },
  { value: "Diamond", label: "菱形" },
  { value: "Triangle", label: "三角形" },
  { value: "RightTriangle", label: "直角三角形" },
  { value: "Parallelogram", label: "平行四边形" },
  { value: "Trapezoid", label: "梯形" },
  { value: "Hexagon", label: "六边形" },
  { value: "Octagon", label: "八边形" },
  { value: "Plus", label: "加号" },
  { value: "Star", label: "星形" },
  { value: "Arrow", label: "箭头" },
  { value: "LeftArrow", label: "左箭头" },
  { value: "RightArrow", label: "右箭头" },
  { value: "UpArrow", label: "上箭头" },
  { value: "DownArrow", label: "下箭头" },
  { value: "LeftRightArrow", label: "左右箭头" },
  { value: "UpDownArrow", label: "上下箭头" },
  { value: "Heart", label: "心形" },
  { value: "Sun", label: "太阳" },
  { value: "Moon", label: "月亮" },
  { value: "LightningBolt", label: "闪电" },
  { value: "FlowChartProcess", label: "流程图-过程" },
  { value: "FlowChartDecision", label: "流程图-决策" },
  { value: "FlowChartInputOutput", label: "流程图-输入输出" },
  { value: "FlowChartDocument", label: "流程图-文档" },
  { value: "FlowChartTerminator", label: "流程图-终止" },
];

export const InsertShapeDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);

  // 基本选项 / Basic options
  const [shapeType, setShapeType] = useState<string>("Rectangle");
  const [insertLocation, setInsertLocation] = useState<InsertLocation>("End");
  const [width, setWidth] = useState<string>("100");
  const [height, setHeight] = useState<string>("100");
  const [name, setName] = useState<string>("");
  const [lockAspectRatio, setLockAspectRatio] = useState(false);
  const [visible, setVisible] = useState(true);
  const [left, setLeft] = useState<string>("");
  const [top, setTop] = useState<string>("");
  const [rotation, setRotation] = useState<string>("");

  // 样式选项 / Style options
  const [enableStyle, setEnableStyle] = useState(false);
  const [fillColor, setFillColor] = useState<string>("#0078D4");
  const [lineColor, setLineColor] = useState<string>("#000000");
  const [lineWeight, setLineWeight] = useState<string>("1");
  const [lineStyle, setLineStyle] = useState<string>("Single");

  // 文本选项 / Text options
  const [enableText, setEnableText] = useState(false);
  const [text, setText] = useState<string>("示例文本");

  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

  const handleInsertShape = async () => {
    if (!shapeType) {
      setResult({
        success: false,
        message: "请选择形状类型",
      });
      return;
    }

    setLoading(true);
    setResult(null);

    try {
      const options: ShapeOptions = {
        width: width ? parseFloat(width) : undefined,
        height: height ? parseFloat(height) : undefined,
        name: name || undefined,
        lockAspectRatio,
        visible,
        left: left ? parseFloat(left) : undefined,
        top: top ? parseFloat(top) : undefined,
        rotation: rotation ? parseFloat(rotation) : undefined,
      };

      // 如果启用样式，添加样式选项 / Add style options if enabled
      if (enableStyle) {
        options.fillColor = fillColor || undefined;
        // 注意：以下线条属性已弃用，当前 Word JavaScript API 不支持
        // Note: The following line properties are deprecated and not supported by current Word JavaScript API
        options.lineColor = lineColor || undefined;
        options.lineWeight = lineWeight ? parseFloat(lineWeight) : undefined;
        options.lineStyle = lineStyle !== "Single" ? lineStyle : undefined;
      }

      // 如果启用文本，添加文本选项 / Add text options if enabled
      if (enableText && text) {
        options.text = text;
      }

      const insertResult = await insertShape(shapeType, insertLocation, options);

      if (insertResult.success) {
        setResult({
          success: true,
          message: insertResult.shapeId
            ? `形状插入成功！标识符: ${insertResult.shapeId}`
            : "形状插入成功！",
        });
      } else {
        setResult({
          success: false,
          message: `形状插入失败: ${insertResult.error}`,
        });
      }
    } catch (error) {
      console.error("插入形状失败:", error);
      setResult({
        success: false,
        message: `插入形状失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      {/* 形状类型 / Shape type */}
      <div className={styles.formRow}>
        <Field label="形状类型" required>
          <Dropdown
            value={COMMON_SHAPES.find((s) => s.value === shapeType)?.label || shapeType}
            selectedOptions={[shapeType]}
            onOptionSelect={(_, data) => setShapeType(data.optionValue as string)}
          >
            {COMMON_SHAPES.map((shape) => (
              <Option key={shape.value} value={shape.value}>
                {shape.label}
              </Option>
            ))}
          </Dropdown>
        </Field>
      </div>

      {/* 基本选项 / Basic options */}
      <div className={styles.formRow}>
        <Label weight="semibold">基本选项</Label>
        <div className={styles.optionsContainer}>
          <Field label="插入位置">
            <Dropdown
              value={insertLocation}
              selectedOptions={[insertLocation]}
              onOptionSelect={(_, data) => setInsertLocation(data.optionValue as InsertLocation)}
            >
              <Option value="Start">文档开头</Option>
              <Option value="End">文档末尾</Option>
              <Option value="Before">选区之前</Option>
              <Option value="After">选区之后</Option>
              <Option value="Replace">替换选区</Option>
            </Dropdown>
          </Field>

          <div className={styles.formGrid}>
            <Field label="宽度（磅）">
              <Input
                type="number"
                value={width}
                onChange={(e) => setWidth(e.target.value)}
                placeholder="100"
              />
            </Field>

            <Field label="高度（磅）">
              <Input
                type="number"
                value={height}
                onChange={(e) => setHeight(e.target.value)}
                placeholder="100"
              />
            </Field>
          </div>

          <Field label="形状名称（可选）">
            <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="MyShape" />
          </Field>

          <Switch
            label="锁定纵横比"
            checked={lockAspectRatio}
            onChange={(_, data) => setLockAspectRatio(data.checked)}
          />

          <Switch label="可见" checked={visible} onChange={(_, data) => setVisible(data.checked)} />
        </div>
      </div>

      {/* 位置和旋转选项 / Position and rotation options */}
      <div className={styles.formRow}>
        <Label weight="semibold">位置和旋转</Label>
        <div className={styles.optionsContainer}>
          <div className={styles.formGrid}>
            <Field label="左边距（磅，可选）">
              <Input
                type="number"
                value={left}
                onChange={(e) => setLeft(e.target.value)}
                placeholder="0"
              />
            </Field>

            <Field label="上边距（磅，可选）">
              <Input
                type="number"
                value={top}
                onChange={(e) => setTop(e.target.value)}
                placeholder="0"
              />
            </Field>
          </div>

          <Field label="旋转角度（度，可选）">
            <Input
              type="number"
              value={rotation}
              onChange={(e) => setRotation(e.target.value)}
              placeholder="0"
            />
          </Field>

          <div className={styles.warningBox}>
            注意：位置参数在某些情况下可能无法完全生效，取决于 Word 的布局设置。
          </div>
        </div>
      </div>

      {/* 样式选项 / Style options */}
      <div className={styles.formRow}>
        <Switch
          label="启用样式设置"
          checked={enableStyle}
          onChange={(_, data) => setEnableStyle(data.checked)}
        />

        {enableStyle && (
          <div className={styles.optionsContainer}>
            <Field label="填充颜色">
              <div className={styles.colorInputGroup}>
                <input
                  type="color"
                  value={fillColor}
                  onChange={(e) => setFillColor(e.target.value)}
                  className={styles.colorPicker}
                />
                <Input
                  value={fillColor}
                  onChange={(e) => setFillColor(e.target.value)}
                  placeholder="#0078D4"
                />
              </div>
            </Field>

            <Field label="线条颜色（当前不支持）">
              <div className={styles.colorInputGroup}>
                <input
                  type="color"
                  value={lineColor}
                  onChange={(e) => setLineColor(e.target.value)}
                  className={styles.colorPicker}
                  disabled
                />
                <Input
                  value={lineColor}
                  onChange={(e) => setLineColor(e.target.value)}
                  placeholder="#000000"
                  disabled
                />
              </div>
            </Field>

            <div className={styles.formGrid}>
              <Field label="线条宽度（磅，当前不支持）">
                <Input
                  type="number"
                  value={lineWeight}
                  onChange={(e) => setLineWeight(e.target.value)}
                  placeholder="1"
                  disabled
                />
              </Field>

              <Field label="线条样式（当前不支持）">
                <Dropdown
                  value={lineStyle}
                  selectedOptions={[lineStyle]}
                  onOptionSelect={(_, data) => setLineStyle(data.optionValue as string)}
                  disabled
                >
                  <Option value="Single">实线</Option>
                  <Option value="Dash">虚线</Option>
                  <Option value="Dot">点线</Option>
                  <Option value="DashDot">点划线</Option>
                  <Option value="DashDotDot">双点划线</Option>
                </Dropdown>
              </Field>
            </div>
            <div className={styles.warningBox}>
              ⚠️ 注意：Word JavaScript API 当前不支持设置形状线条样式。这些选项已禁用。
            </div>
          </div>
        )}
      </div>

      {/* 文本选项 / Text options */}
      <div className={styles.formRow}>
        <Switch
          label="添加文本内容"
          checked={enableText}
          onChange={(_, data) => setEnableText(data.checked)}
        />

        {enableText && (
          <div className={styles.optionsContainer}>
            <Field label="文本内容">
              <Input
                value={text}
                onChange={(e) => setText(e.target.value)}
                placeholder="请输入文本"
              />
            </Field>
            <div className={styles.warningBox}>注意：并非所有形状类型都支持文本内容。</div>
          </div>
        )}
      </div>

      {/* 操作按钮 / Action buttons */}
      <div className={styles.buttonGroup}>
        <Button
          appearance="primary"
          onClick={handleInsertShape}
          disabled={loading || !shapeType}
          icon={loading ? <Spinner size="tiny" /> : <Shapes24Regular />}
        >
          {loading ? "正在插入..." : "插入形状"}
        </Button>
      </div>

      {/* 结果显示 / Result display */}
      {result && (
        <div
          className={`${styles.resultMessage} ${result.success ? styles.success : styles.error}`}
        >
          {result.message}
        </div>
      )}
    </div>
  );
};
