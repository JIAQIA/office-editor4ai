/**
 * 文件名: InsertTextBoxDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: insertTextBox工具的调试组件
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
  Textarea,
} from "@fluentui/react-components";
import { Textbox24Regular } from "@fluentui/react-icons";
import {
  insertTextBox,
  type TextBoxOptions,
  type InsertLocation,
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
});

export const InsertTextBoxDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);

  // 基本选项 / Basic options
  const [text, setText] = useState<string>("示例文本框内容");
  const [insertLocation, setInsertLocation] = useState<InsertLocation>("End");
  const [width, setWidth] = useState<string>("150");
  const [height, setHeight] = useState<string>("100");
  const [name, setName] = useState<string>("");
  const [lockAspectRatio, setLockAspectRatio] = useState(false);
  const [visible, setVisible] = useState(true);
  const [left, setLeft] = useState<string>("");
  const [top, setTop] = useState<string>("");
  const [rotation, setRotation] = useState<string>("");

  // 文本格式选项 / Text format options
  const [enableFormat, setEnableFormat] = useState(false);
  const [fontName, setFontName] = useState<string>("Arial");
  const [fontSize, setFontSize] = useState<string>("12");
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const [underline, setUnderline] = useState<string>("None");
  const [color, setColor] = useState<string>("#000000");
  const [highlightColor, setHighlightColor] = useState<string>("");
  const [strikeThrough, setStrikeThrough] = useState(false);

  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

  const handleInsertTextBox = async () => {
    if (!text.trim()) {
      setResult({
        success: false,
        message: "请输入文本框内容",
      });
      return;
    }

    setLoading(true);
    setResult(null);

    try {
      const options: TextBoxOptions = {
        width: width ? parseFloat(width) : undefined,
        height: height ? parseFloat(height) : undefined,
        name: name || undefined,
        lockAspectRatio,
        visible,
        left: left ? parseFloat(left) : undefined,
        top: top ? parseFloat(top) : undefined,
        rotation: rotation ? parseFloat(rotation) : undefined,
      };

      // 如果启用格式，添加格式选项 / Add format options if enabled
      if (enableFormat) {
        options.format = {
          fontName: fontName || undefined,
          fontSize: fontSize ? parseFloat(fontSize) : undefined,
          bold,
          italic,
          underline: underline !== "None" ? (underline as Word.UnderlineType) : undefined,
          color: color || undefined,
          highlightColor: highlightColor || undefined,
          strikeThrough,
        };
      }

      const insertResult = await insertTextBox(text, insertLocation, options);

      if (insertResult.success) {
        setResult({
          success: true,
          message: insertResult.textBoxId
            ? `文本框插入成功！标识符: ${insertResult.textBoxId}`
            : "文本框插入成功！",
        });
      } else {
        setResult({
          success: false,
          message: `文本框插入失败: ${insertResult.error}`,
        });
      }
    } catch (error) {
      console.error("插入文本框失败:", error);
      setResult({
        success: false,
        message: `插入文本框失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      {/* 文本内容 / Text content */}
      <div className={styles.formRow}>
        <Field label="文本框内容" required>
          <Textarea
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder="请输入文本框内容"
            rows={4}
          />
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
                placeholder="150"
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

          <Field label="文本框名称（可选）">
            <Input
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="MyTextBox"
            />
          </Field>

          <Switch
            label="锁定纵横比"
            checked={lockAspectRatio}
            onChange={(_, data) => setLockAspectRatio(data.checked)}
          />

          <Switch
            label="可见"
            checked={visible}
            onChange={(_, data) => setVisible(data.checked)}
          />
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

      {/* 文本格式选项 / Text format options */}
      <div className={styles.formRow}>
        <Switch
          label="启用文本格式"
          checked={enableFormat}
          onChange={(_, data) => setEnableFormat(data.checked)}
        />

        {enableFormat && (
          <div className={styles.optionsContainer}>
            <div className={styles.formGrid}>
              <Field label="字体">
                <Input
                  value={fontName}
                  onChange={(e) => setFontName(e.target.value)}
                  placeholder="Arial"
                />
              </Field>

              <Field label="字号">
                <Input
                  type="number"
                  value={fontSize}
                  onChange={(e) => setFontSize(e.target.value)}
                  placeholder="12"
                />
              </Field>
            </div>

            <div className={styles.formGrid}>
              <Field label="文字颜色">
                <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                  <input
                    type="color"
                    value={color}
                    onChange={(e) => setColor(e.target.value)}
                    style={{ width: "40px", height: "32px", border: "none", cursor: "pointer" }}
                  />
                  <Input
                    value={color}
                    onChange={(e) => setColor(e.target.value)}
                    placeholder="#000000"
                  />
                </div>
              </Field>

              <Field label="高亮颜色（可选）">
                <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                  <input
                    type="color"
                    value={highlightColor || "#FFFF00"}
                    onChange={(e) => setHighlightColor(e.target.value)}
                    style={{ width: "40px", height: "32px", border: "none", cursor: "pointer" }}
                  />
                  <Input
                    value={highlightColor}
                    onChange={(e) => setHighlightColor(e.target.value)}
                    placeholder="#FFFF00"
                  />
                </div>
              </Field>
            </div>

            <Field label="下划线">
              <Dropdown
                value={underline}
                selectedOptions={[underline]}
                onOptionSelect={(_, data) => setUnderline(data.optionValue as string)}
              >
                <Option value="None">无</Option>
                <Option value="Single">单下划线</Option>
                <Option value="Double">双下划线</Option>
                <Option value="Dotted">点状下划线</Option>
                <Option value="Dash">虚线下划线</Option>
              </Dropdown>
            </Field>

            <Switch label="粗体" checked={bold} onChange={(_, data) => setBold(data.checked)} />

            <Switch label="斜体" checked={italic} onChange={(_, data) => setItalic(data.checked)} />

            <Switch
              label="删除线"
              checked={strikeThrough}
              onChange={(_, data) => setStrikeThrough(data.checked)}
            />
          </div>
        )}
      </div>

      {/* 操作按钮 / Action buttons */}
      <div className={styles.buttonGroup}>
        <Button
          appearance="primary"
          onClick={handleInsertTextBox}
          disabled={loading || !text.trim()}
          icon={loading ? <Spinner size="tiny" /> : <Textbox24Regular />}
        >
          {loading ? "正在插入..." : "插入文本框"}
        </Button>
      </div>

      {/* 结果显示 / Result display */}
      {result && (
        <div className={`${styles.resultMessage} ${result.success ? styles.success : styles.error}`}>
          {result.message}
        </div>
      )}
    </div>
  );
};
