/**
 * 文件名: InsertPageBreakDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/04
 * 最后修改日期: 2025/12/04
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: insertPageBreak工具的调试组件 / Debug component for insertPageBreak tool
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
import { DocumentPageBreak24Regular } from "@fluentui/react-icons";
import { insertPageBreak, type InsertLocation } from "../../../word-tools";

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

// 插入位置选项 / Insert location options
const INSERT_LOCATIONS: Array<{ value: InsertLocation; label: string; description: string }> = [
  { value: "Start", label: "文档开头", description: "在文档最开始插入分页符" },
  { value: "End", label: "文档末尾", description: "在文档最后插入分页符" },
  { value: "Before", label: "选中内容之前", description: "在当前选中内容之前插入分页符" },
  { value: "After", label: "选中内容之后", description: "在当前选中内容之后插入分页符" },
  { value: "Replace", label: "替换选中内容", description: "用分页符替换当前选中的内容" },
];

export const InsertPageBreakDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [insertLocation, setInsertLocation] = useState<InsertLocation>("End");
  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

  const handleInsert = async () => {
    setLoading(true);
    setResult(null);

    try {
      const insertResult = await insertPageBreak(insertLocation);

      if (insertResult.success) {
        setResult({
          success: true,
          message: `成功在 ${INSERT_LOCATIONS.find((loc) => loc.value === insertLocation)?.label} 插入分页符`,
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
    setInsertLocation("End");
    setResult(null);
  };

  return (
    <div className={styles.container}>
      {/* 标题 / Header */}
      <div className={styles.header}>
        <DocumentPageBreak24Regular />
        <span className={styles.title}>插入分页符</span>
      </div>

      {/* 描述 / Description */}
      <div className={styles.description}>
        分页符用于强制在指定位置开始新的一页。选择插入位置后点击"插入分页符"按钮即可。
      </div>

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
          <li>分页符会在指定位置强制开始新页面</li>
          <li>选择"选中内容之前/之后"时，请先在文档中选中内容</li>
          <li>选择"替换选中内容"会删除选中内容并插入分页符</li>
        </ul>
      </div>

      {/* 操作按钮 / Action buttons */}
      <div className={styles.buttonGroup}>
        <Button appearance="primary" onClick={handleInsert} disabled={loading}>
          {loading ? <Spinner size="tiny" /> : "插入分页符"}
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
