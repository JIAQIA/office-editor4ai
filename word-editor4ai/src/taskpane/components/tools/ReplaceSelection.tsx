/**
 * 文件名: ReplaceSelection.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 替换选中内容的工具组件
 */

/* global console */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  makeStyles,
  tokens,
  Spinner,
  Switch,
  Label,
  Input,
  Textarea,
  Card,
  CardHeader,
  Divider,
} from "@fluentui/react-components";
import { Delete24Regular, Image24Regular } from "@fluentui/react-icons";
import { replaceSelection, replaceTextAtSelection, type TextFormat, type ImageData } from "../../../word-tools";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
    gap: "16px",
    padding: "8px",
  },
  formContainer: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  formRow: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  buttonContainer: {
    width: "100%",
    display: "flex",
    gap: "8px",
    justifyContent: "center",
    marginTop: "8px",
  },
  resultCard: {
    width: "100%",
  },
  successMessage: {
    padding: "12px",
    backgroundColor: tokens.colorPaletteGreenBackground2,
    borderRadius: tokens.borderRadiusSmall,
    color: tokens.colorPaletteGreenForeground1,
  },
  errorMessage: {
    padding: "12px",
    backgroundColor: tokens.colorPaletteRedBackground2,
    borderRadius: tokens.borderRadiusSmall,
    color: tokens.colorPaletteRedForeground1,
  },
  imageUploadContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
  },
  imagePreviewList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    maxHeight: "300px",
    overflowY: "auto",
  },
  imagePreviewItem: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusSmall,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  imagePreview: {
    width: "60px",
    height: "60px",
    objectFit: "cover",
    borderRadius: tokens.borderRadiusSmall,
  },
  imageInfo: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    fontSize: "12px",
  },
  uploadButton: {
    width: "100%",
  },
});

/**
 * 替换选中内容的工具组件
 * Tool component for replacing selected content
 */
export const ReplaceSelection: React.FC = () => {
  const styles = useStyles();

  // 状态管理 / State management
  const [loading, setLoading] = useState(false);
  const [text, setText] = useState("这是要插入的文本内容");
  const [shouldReplace, setShouldReplace] = useState(true);
  const [applyFormat, setApplyFormat] = useState(false);
  const [fontName, setFontName] = useState("Arial");
  const [fontSize, setFontSize] = useState("14");
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const [color, setColor] = useState("#000000");
  const [images, setImages] = useState<ImageData[]>([]);
  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

  /**
   * 处理图片上传 / Handle image upload
   */
  const handleImageUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const base64 = e.target?.result as string;
        const img = new Image();
        img.onload = () => {
          setImages((prev) => [
            ...prev,
            {
              base64,
              width: img.width > 400 ? 400 : img.width,
              height: img.height > 400 ? (400 * img.height) / img.width : img.height,
              altText: file.name,
            },
          ]);
        };
        img.src = base64;
      };
      reader.readAsDataURL(file);
    });

    // 重置 input 以允许重复上传同一文件 / Reset input to allow re-uploading the same file
    event.target.value = "";
  };

  /**
   * 删除图片 / Remove image
   */
  const handleRemoveImage = (index: number) => {
    setImages((prev) => prev.filter((_, i) => i !== index));
  };

  /**
   * 处理替换操作 / Handle replace operation
   */
  const handleReplace = async () => {
    if (!text.trim() && images.length === 0) {
      setResult({
        success: false,
        message: "请输入文本内容或上传图片",
      });
      return;
    }

    setLoading(true);
    setResult(null);

    try {
      // 构建格式对象 / Build format object
      const format: TextFormat | undefined = applyFormat
        ? {
            fontName,
            fontSize: parseFloat(fontSize),
            bold,
            italic,
            color,
          }
        : undefined;

      // 调用替换函数 / Call replace function
      await replaceSelection({
        text: text.trim() || undefined,
        format,
        images: images.length > 0 ? images : undefined,
        replaceSelection: shouldReplace,
      });

      setResult({
        success: true,
        message: `成功${shouldReplace ? "替换" : "插入"}内容${text.trim() ? "（文本）" : ""}${images.length > 0 ? `（${images.length}张图片）` : ""}`,
      });
    } catch (error) {
      console.error("插入内容失败:", error);
      setResult({
        success: false,
        message: `${shouldReplace ? "替换" : "插入"}失败: ${error.message}`,
      });
    } finally {
      setLoading(false);
    }
  };

  /**
   * 处理简单文本替换 / Handle simple text replace
   */
  const handleSimpleReplace = async () => {
    if (!text.trim()) {
      setResult({
        success: false,
        message: "请输入要插入的文本内容",
      });
      return;
    }

    setLoading(true);
    setResult(null);

    try {
      await replaceTextAtSelection(text);

      setResult({
        success: true,
        message: "成功替换文本内容（保持原格式）",
      });
    } catch (error) {
      console.error("替换文本失败:", error);
      setResult({
        success: false,
        message: `替换失败: ${error.message}`,
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      {/* 表单区域 / Form area */}
      <div className={styles.formContainer}>
        {/* 文本输入 / Text input */}
        <div className={styles.formRow}>
          <Label weight="semibold">文本内容</Label>
          <Textarea
            value={text}
            onChange={(_, data) => setText(data.value)}
            placeholder="输入要插入的文本..."
            rows={4}
          />
        </div>

        {/* 图片上传 / Image upload */}
        <div className={styles.formRow}>
          <Label weight="semibold">图片（可选）</Label>
          <div className={styles.imageUploadContainer}>
            <input
              type="file"
              accept="image/*"
              multiple
              onChange={handleImageUpload}
              style={{ display: "none" }}
              id="image-upload-input"
            />
            <Button
              icon={<Image24Regular />}
              appearance="secondary"
              onClick={() => document.getElementById("image-upload-input")?.click()}
              className={styles.uploadButton}
            >
              上传图片（支持多选）
            </Button>

            {images.length > 0 && (
              <div className={styles.imagePreviewList}>
                {images.map((img, index) => (
                  <div key={index} className={styles.imagePreviewItem}>
                    <img src={img.base64} alt={img.altText} className={styles.imagePreview} />
                    <div className={styles.imageInfo}>
                      <div>
                        <strong>{img.altText}</strong>
                      </div>
                      <div>
                        尺寸: {Math.round(img.width || 0)} × {Math.round(img.height || 0)} 磅
                      </div>
                    </div>
                    <Button
                      icon={<Delete24Regular />}
                      appearance="subtle"
                      onClick={() => handleRemoveImage(index)}
                      title="删除图片"
                    />
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* 替换选项 / Replace option */}
        <div className={styles.formRow}>
          <Switch
            checked={shouldReplace}
            onChange={(_, data) => setShouldReplace(data.checked)}
            label={shouldReplace ? "替换选中内容" : "在选中位置后插入"}
          />
        </div>

        <Divider />

        {/* 格式选项 / Format options */}
        <div className={styles.formRow}>
          <Switch
            checked={applyFormat}
            onChange={(_, data) => setApplyFormat(data.checked)}
            label="应用自定义格式"
          />
        </div>

        {applyFormat && (
          <>
            <div className={styles.formRow}>
              <Label>字体</Label>
              <Input
                value={fontName}
                onChange={(_, data) => setFontName(data.value)}
                placeholder="Arial"
              />
            </div>

            <div className={styles.formRow}>
              <Label>字号</Label>
              <Input
                type="number"
                value={fontSize}
                onChange={(_, data) => setFontSize(data.value)}
                placeholder="14"
              />
            </div>

            <div className={styles.formRow}>
              <Switch checked={bold} onChange={(_, data) => setBold(data.checked)} label="加粗" />
            </div>

            <div className={styles.formRow}>
              <Switch
                checked={italic}
                onChange={(_, data) => setItalic(data.checked)}
                label="斜体"
              />
            </div>

            <div className={styles.formRow}>
              <Label>颜色</Label>
              <input
                type="color"
                value={color}
                onChange={(e) => setColor(e.target.value)}
                style={{
                  width: "100%",
                  height: "32px",
                  border: `1px solid ${tokens.colorNeutralStroke1}`,
                  borderRadius: tokens.borderRadiusSmall,
                  cursor: "pointer",
                }}
              />
            </div>
          </>
        )}
      </div>

      {/* 按钮区域 / Button area */}
      <div className={styles.buttonContainer}>
        <Button appearance="primary" onClick={handleReplace} disabled={loading}>
          {loading ? <Spinner size="tiny" /> : applyFormat ? "替换并应用格式" : "替换内容"}
        </Button>
        {!images.length && (
          <Button onClick={handleSimpleReplace} disabled={loading}>
            {loading ? <Spinner size="tiny" /> : "替换文本（保持原格式）"}
          </Button>
        )}
      </div>

      {/* 结果显示 / Result display */}
      {result && (
        <Card className={styles.resultCard}>
          <CardHeader header={result.success ? "✅ 成功" : "❌ 失败"} />
          <div className={result.success ? styles.successMessage : styles.errorMessage}>
            {result.message}
          </div>
        </Card>
      )}
    </div>
  );
};
