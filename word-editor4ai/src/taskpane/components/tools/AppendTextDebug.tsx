/**
 * 文件名: AppendTextDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: appendText工具的调试组件
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
  Textarea,
  tokens,
} from "@fluentui/react-components";
import { Image24Regular, Delete24Regular } from "@fluentui/react-icons";
import { type AppendOptions, type ImageData } from "../../../word-tools";

interface AppendTextDebugProps {
  appendText?: (options: AppendOptions) => Promise<void>;
  appendContent?: (options: AppendOptions) => Promise<void>;
}

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
  formatContainer: {
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

export const AppendTextDebug: React.FC<AppendTextDebugProps> = ({ appendText, appendContent }) => {
  const styles = useStyles();
  const [text, setText] = useState("这是追加到文档末尾的文本");
  const [loading, setLoading] = useState(false);
  const [applyFormat, setApplyFormat] = useState(false);
  const [fontName, setFontName] = useState("Arial");
  const [fontSize, setFontSize] = useState("14");
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const [color, setColor] = useState("#000000");
  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);
  const [images, setImages] = useState<ImageData[]>([]);

  const handleAppendText = async () => {
    setLoading(true);
    setResult(null);

    try {
      const format = applyFormat
        ? {
            fontName,
            fontSize: parseInt(fontSize),
            bold,
            italic,
            color,
          }
        : undefined;

      if (appendContent) {
        await appendContent({
          text,
          format,
          images: images.length > 0 ? images : undefined,
        });
      } else if (appendText) {
        await appendText({
          text,
          format,
          images: images.length > 0 ? images : undefined,
        });
      } else {
        throw new Error("没有可用的追加函数");
      }

      setResult({
        success: true,
        message: `内容已成功追加${text ? "（文本）" : ""}${images.length > 0 ? `（${images.length}张图片）` : ""}`,
      });
    } catch (error) {
      console.error("追加内容失败:", error);
      setResult({
        success: false,
        message: `追加内容失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

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
  };

  const removeImage = (index: number) => {
    setImages((prev) => prev.filter((_, i) => i !== index));
  };

  return (
    <div className={styles.container}>
      <div className={styles.formRow}>
        <Label>要追加的文本内容:</Label>
        <Textarea
          value={text}
          onChange={(e) => setText(e.target.value)}
          resize="vertical"
          rows={3}
        />
      </div>

      <Switch
        label="应用文本格式"
        checked={applyFormat}
        onChange={(_, data) => setApplyFormat(data.checked)}
      />

      {applyFormat && (
        <div className={styles.formatContainer}>
          <div className={styles.formRow}>
            <Label>字体名称:</Label>
            <Input value={fontName} onChange={(e) => setFontName(e.target.value)} />
          </div>

          <div className={styles.formRow}>
            <Label>字体大小:</Label>
            <Input
              type="number"
              value={fontSize}
              onChange={(e) => setFontSize(e.target.value)}
              min="1"
              max="72"
            />
          </div>

          <div className={styles.formRow}>
            <Switch label="加粗" checked={bold} onChange={(_, data) => setBold(data.checked)} />
            <Switch label="斜体" checked={italic} onChange={(_, data) => setItalic(data.checked)} />
          </div>

          <div className={styles.formRow}>
            <Label>字体颜色:</Label>
            <input
              type="color"
              value={color}
              onChange={(e) => setColor(e.target.value)}
              style={{ width: "60px", height: "30px", padding: "2px" }}
            />
          </div>
        </div>
      )}

      {/* 图片上传区域 */}
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
                    onClick={() => removeImage(index)}
                    title="删除图片"
                  />
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className={styles.buttonGroup}>
        <Button
          appearance="primary"
          onClick={handleAppendText}
          disabled={loading || !text}
          icon={loading ? <Spinner size="tiny" /> : undefined}
        >
          {loading ? "正在追加..." : "追加文本"}
        </Button>
      </div>

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
