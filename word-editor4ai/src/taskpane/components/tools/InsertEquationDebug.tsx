/**
 * 文件名: InsertEquationDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: insertEquation工具的调试组件 / Debug component for insertEquation tool
 */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  makeStyles,
  Spinner,
  Dropdown,
  Option,
  tokens,
  Field,
  Textarea,
} from "@fluentui/react-components";
import { MathFormula24Regular } from "@fluentui/react-icons";
import { insertEquation, type InsertLocation } from "../../../word-tools";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    paddingBottom: "8px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  title: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  description: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    lineHeight: "1.5",
  },
  formRow: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
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
  infoBox: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    fontSize: tokens.fontSizeBase200,
    lineHeight: "1.6",
  },
  exampleBox: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    fontSize: tokens.fontSizeBase200,
    fontFamily: "monospace",
  },
  exampleItem: {
    marginBottom: "8px",
    cursor: "pointer",
    padding: "4px",
    borderRadius: tokens.borderRadiusSmall,
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
});

const INSERT_LOCATIONS: Array<{ value: InsertLocation; label: string; description: string }> = [
  { value: "Start", label: "文档开头", description: "在文档最开始插入公式" },
  { value: "End", label: "文档末尾", description: "在文档最后插入公式" },
  { value: "Before", label: "选中内容之前", description: "在当前选中内容之前插入公式" },
  { value: "After", label: "选中内容之后", description: "在当前选中内容之后插入公式" },
  { value: "Replace", label: "替换选中内容", description: "用公式替换当前选中的内容" },
];

const EXAMPLE_EQUATIONS = [
  { label: "质能方程", latex: "E = mc^2" },
  { label: "勾股定理", latex: "a^2 + b^2 = c^2" },
  { label: "分数", latex: "\\frac{a}{b}" },
  { label: "根号", latex: "\\sqrt{x}" },
  { label: "求和", latex: "\\sum_{i=1}^{n} x_i" },
  { label: "积分", latex: "\\int_{a}^{b} f(x) dx" },
  { label: "二次方程", latex: "x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}" },
  { label: "极限", latex: "\\lim_{x \\to \\infty} f(x)" },
];

export const InsertEquationDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [latex, setLatex] = useState<string>("E = mc^2");
  const [insertLocation, setInsertLocation] = useState<InsertLocation>("End");
  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

  const handleInsert = async () => {
    if (!latex.trim()) {
      setResult({
        success: false,
        message: "请输入 LaTeX 公式",
      });
      return;
    }

    setLoading(true);
    setResult(null);

    try {
      const insertResult = await insertEquation(latex, insertLocation);

      if (insertResult.success) {
        setResult({
          success: true,
          message: `成功在 ${INSERT_LOCATIONS.find((loc) => loc.value === insertLocation)?.label} 插入公式`,
        });
      } else {
        setResult({
          success: false,
          message: `插入失败: ${insertResult.error || "未知错误"}`,
        });
      }
    } catch (error) {
      setResult({
        success: false,
        message: `插入失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const handleReset = () => {
    setLatex("E = mc^2");
    setInsertLocation("End");
    setResult(null);
  };

  const handleExampleClick = (exampleLatex: string) => {
    setLatex(exampleLatex);
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <MathFormula24Regular />
        <span className={styles.title}>插入公式</span>
      </div>

      <div className={styles.description}>
        使用 LaTeX 格式插入数学公式。支持常见的数学符号、分数、根号、上下标等。
      </div>

      <Field label="LaTeX 公式" required>
        <Textarea
          value={latex}
          onChange={(_, data) => setLatex(data.value)}
          placeholder="输入 LaTeX 格式的公式，例如: E = mc^2"
          rows={3}
        />
      </Field>

      <Field label="插入位置" required>
        <Dropdown
          placeholder="选择插入位置"
          value={INSERT_LOCATIONS.find((loc) => loc.value === insertLocation)?.label || ""}
          selectedOptions={[insertLocation]}
          onOptionSelect={(_, data) => {
            setInsertLocation(data.optionValue as InsertLocation);
          }}
        >
          {INSERT_LOCATIONS.map((location) => (
            <Option key={location.value} value={location.value} text={location.label}>
              <div>
                <div>{location.label}</div>
                <div style={{ fontSize: "12px", color: tokens.colorNeutralForeground3 }}>
                  {location.description}
                </div>
              </div>
            </Option>
          ))}
        </Dropdown>
      </Field>

      <div className={styles.infoBox}>
        <strong>LaTeX 语法示例：</strong>
        <div className={styles.exampleBox}>
          {EXAMPLE_EQUATIONS.map((example, index) => (
            <div
              key={index}
              className={styles.exampleItem}
              onClick={() => handleExampleClick(example.latex)}
              title="点击使用此示例"
            >
              <strong>{example.label}:</strong> {example.latex}
            </div>
          ))}
        </div>
      </div>

      <div className={styles.infoBox}>
        <strong>支持的 LaTeX 语法：</strong>
        <ul style={{ margin: "8px 0", paddingLeft: "20px" }}>
          <li>上标: x^2 或 x^{"{"}2{"}"}</li>
          <li>下标: x_i 或 x_{"{"}i{"}"}</li>
          <li>分数: \frac{"{"}a{"}"}{"{"}b{"}"}</li>
          <li>根号: \sqrt{"{"}x{"}"}</li>
          <li>简单公式会被转换为 Word 原生数学公式（OMML 格式）</li>
          <li>复杂公式可能需要使用 Word 内置的公式编辑器</li>
        </ul>
      </div>

      <div className={styles.buttonGroup}>
        <Button appearance="primary" onClick={handleInsert} disabled={loading}>
          {loading ? <Spinner size="tiny" /> : "插入公式"}
        </Button>
        <Button appearance="secondary" onClick={handleReset} disabled={loading}>
          重置
        </Button>
      </div>

      {result && (
        <div className={`${styles.resultMessage} ${result.success ? styles.success : styles.error}`}>
          {result.message}
        </div>
      )}
    </div>
  );
};
