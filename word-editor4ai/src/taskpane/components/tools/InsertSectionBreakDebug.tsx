/**
 * 文件名: InsertSectionBreakDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: insertSectionBreak工具的调试组件 / Debug component for insertSectionBreak tool
 */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  Label,
  makeStyles,
  Spinner,
  Dropdown,
  Option,
  tokens,
  Field,
} from "@fluentui/react-components";
import { DocumentSplitHint24Regular } from "@fluentui/react-icons";
import { insertSectionBreak, type InsertLocation, type SectionBreakType } from "../../../word-tools";

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
});

// 分节符类型选项 / Section break type options
const SECTION_BREAK_TYPES: Array<{ value: SectionBreakType; label: string; description: string }> = [
  { value: "Continuous", label: "连续", description: "新节从当前页继续，不换页" },
  { value: "NextPage", label: "下一页", description: "新节从下一页开始" },
  { value: "OddPage", label: "奇数页", description: "新节从下一个奇数页开始" },
  { value: "EvenPage", label: "偶数页", description: "新节从下一个偶数页开始" },
];

// 插入位置选项 / Insert location options
const INSERT_LOCATIONS: Array<{ value: InsertLocation; label: string; description: string }> = [
  { value: "Start", label: "文档开头", description: "在文档最开始插入分节符" },
  { value: "End", label: "文档末尾", description: "在文档最后插入分节符" },
  { value: "Before", label: "选中内容之前", description: "在当前选中内容之前插入分节符" },
  { value: "After", label: "选中内容之后", description: "在当前选中内容之后插入分节符" },
  { value: "Replace", label: "替换选中内容", description: "用分节符替换当前选中的内容" },
];

export const InsertSectionBreakDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [breakType, setBreakType] = useState<SectionBreakType>("NextPage");
  const [insertLocation, setInsertLocation] = useState<InsertLocation>("End");
  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

  const handleInsert = async () => {
    setLoading(true);
    setResult(null);

    try {
      const insertResult = await insertSectionBreak(breakType, insertLocation);

      if (insertResult.success) {
        setResult({
          success: true,
          message: `成功在 ${INSERT_LOCATIONS.find((loc) => loc.value === insertLocation)?.label} 插入${SECTION_BREAK_TYPES.find((type) => type.value === breakType)?.label}分节符${insertResult.sectionIndex !== undefined ? `，新节索引: ${insertResult.sectionIndex}` : ""}`,
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
    setBreakType("NextPage");
    setInsertLocation("End");
    setResult(null);
  };

  return (
    <div className={styles.container}>
      {/* 标题 / Header */}
      <div className={styles.header}>
        <DocumentSplitHint24Regular />
        <span className={styles.title}>插入分节符</span>
      </div>

      {/* 描述 / Description */}
      <div className={styles.description}>
        分节符用于将文档分成不同的节，每节可以有独立的页面设置（如页眉、页脚、页边距、纸张方向等）。
      </div>

      {/* 分节符类型选择 / Section break type selection */}
      <Field label="分节符类型" required>
        <Dropdown
          placeholder="选择分节符类型"
          value={SECTION_BREAK_TYPES.find((type) => type.value === breakType)?.label || ""}
          selectedOptions={[breakType]}
          onOptionSelect={(_, data) => {
            setBreakType(data.optionValue as SectionBreakType);
          }}
        >
          {SECTION_BREAK_TYPES.map((type) => (
            <Option key={type.value} value={type.value} text={type.label}>
              <div>
                <div>{type.label}</div>
                <div style={{ fontSize: "12px", color: tokens.colorNeutralForeground3 }}>
                  {type.description}
                </div>
              </div>
            </Option>
          ))}
        </Dropdown>
      </Field>

      {/* 插入位置选择 / Insert location selection */}
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

      {/* 提示信息 / Info box */}
      <div className={styles.infoBox}>
        <strong>使用提示：</strong>
        <ul style={{ margin: "8px 0", paddingLeft: "20px" }}>
          <li>
            <strong>连续分节符：</strong>新节从当前页继续，适合在同一页内改变列数或页边距
          </li>
          <li>
            <strong>下一页分节符：</strong>最常用，新节从下一页开始
          </li>
          <li>
            <strong>奇数页/偶数页分节符：</strong>用于书籍排版，确保章节从特定页面开始
          </li>
          <li>选择"选中内容之前/之后"时，请先在文档中选中内容</li>
        </ul>
      </div>

      {/* 操作按钮 / Action buttons */}
      <div className={styles.buttonGroup}>
        <Button appearance="primary" onClick={handleInsert} disabled={loading}>
          {loading ? <Spinner size="tiny" /> : "插入分节符"}
        </Button>
        <Button appearance="secondary" onClick={handleReset} disabled={loading}>
          重置
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
