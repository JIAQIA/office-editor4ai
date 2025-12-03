/**
 * 文件名: InsertImageDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: insertImage工具的调试组件
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
import { Image24Regular, Delete24Regular } from "@fluentui/react-icons";
import {
  insertImage,
  type InsertImageOptions,
  type InsertLocation,
  type ImageLayoutType,
  type WrapType,
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
  imageUploadContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
  },
  imagePreview: {
    width: "100%",
    maxHeight: "200px",
    objectFit: "contain",
    borderRadius: tokens.borderRadiusSmall,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  uploadButton: {
    width: "100%",
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
});

export const InsertImageDebug: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [imageBase64, setImageBase64] = useState<string>("");
  const [imageFile, setImageFile] = useState<File | null>(null);
  
  // 基本选项 / Basic options
  const [width, setWidth] = useState<string>("300");
  const [height, setHeight] = useState<string>("200");
  const [altText, setAltText] = useState<string>("示例图片");
  const [description, setDescription] = useState<string>("");
  const [insertLocation, setInsertLocation] = useState<InsertLocation>("Replace");
  const [layoutType, setLayoutType] = useState<ImageLayoutType>("inline");
  const [keepAspectRatio, setKeepAspectRatio] = useState(true);
  const [hyperlink, setHyperlink] = useState<string>("");
  
  // 浮动图片选项 / Floating image options
  const [wrapType, setWrapType] = useState<WrapType>("Square");
  const [leftPosition, setLeftPosition] = useState<string>("0");
  const [topPosition, setTopPosition] = useState<string>("0");
  const [lockAnchor, setLockAnchor] = useState(false);
  const [allowOverlap, setAllowOverlap] = useState(false);
  
  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

  const handleImageUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setImageFile(file);
    const reader = new FileReader();
    reader.onload = (e) => {
      const base64 = e.target?.result as string;
      setImageBase64(base64);
      
      // 自动获取图片尺寸 / Auto get image dimensions
      const img = new Image();
      img.onload = () => {
        setWidth(String(Math.min(img.width, 400)));
        setHeight(String(Math.min(img.height, 400)));
      };
      img.src = base64;
    };
    reader.readAsDataURL(file);
  };

  const handleInsertImage = async () => {
    if (!imageBase64) {
      setResult({
        success: false,
        message: "请先上传图片",
      });
      return;
    }

    setLoading(true);
    setResult(null);

    try {
      const options: InsertImageOptions = {
        base64: imageBase64,
        width: width ? parseFloat(width) : undefined,
        height: height ? parseFloat(height) : undefined,
        altText: altText || undefined,
        description: description || undefined,
        insertLocation,
        layoutType,
        keepAspectRatio,
        hyperlink: hyperlink || undefined,
      };

      // 如果是浮动图片，添加浮动选项 / Add floating options if floating image
      if (layoutType === "floating") {
        options.floatingOptions = {
          wrapType,
          position: {
            left: leftPosition ? parseFloat(leftPosition) : undefined,
            top: topPosition ? parseFloat(topPosition) : undefined,
          },
          lockAnchor,
          allowOverlap,
        };
      }

      const insertResult = await insertImage(options);

      if (insertResult.success) {
        setResult({
          success: true,
          message: insertResult.imageId 
            ? `图片插入成功！标识符: ${insertResult.imageId}` 
            : "图片插入成功！",
        });
      } else {
        setResult({
          success: false,
          message: `图片插入失败: ${insertResult.error}`,
        });
      }
    } catch (error) {
      console.error("插入图片失败:", error);
      setResult({
        success: false,
        message: `插入图片失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  const clearImage = () => {
    setImageBase64("");
    setImageFile(null);
    setResult(null);
  };

  return (
    <div className={styles.container}>
      {/* 图片上传区域 / Image upload area */}
      <div className={styles.formRow}>
        <Label weight="semibold">选择图片</Label>
        <div className={styles.imageUploadContainer}>
          <input
            type="file"
            accept="image/*"
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
            {imageFile ? `已选择: ${imageFile.name}` : "上传图片"}
          </Button>

          {imageBase64 && (
            <>
              <img src={imageBase64} alt="预览" className={styles.imagePreview} />
              <Button
                icon={<Delete24Regular />}
                appearance="subtle"
                onClick={clearImage}
              >
                清除图片
              </Button>
            </>
          )}
        </div>
      </div>

      {/* 基本选项 / Basic options */}
      <div className={styles.formRow}>
        <Label weight="semibold">基本选项</Label>
        <div className={styles.optionsContainer}>
          <div className={styles.formGrid}>
            <Field label="宽度（磅）">
              <Input
                type="number"
                value={width}
                onChange={(e) => setWidth(e.target.value)}
                placeholder="300"
              />
            </Field>

            <Field label="高度（磅）">
              <Input
                type="number"
                value={height}
                onChange={(e) => setHeight(e.target.value)}
                placeholder="200"
              />
            </Field>
          </div>

          <Field label="替代文本">
            <Input
              value={altText}
              onChange={(e) => setAltText(e.target.value)}
              placeholder="图片描述"
            />
          </Field>

          <Field label="详细描述（可选）">
            <Input
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="图片的详细描述"
            />
          </Field>

          <Field label="超链接（可选）">
            <Input
              value={hyperlink}
              onChange={(e) => setHyperlink(e.target.value)}
              placeholder="https://example.com"
            />
          </Field>

          <div className={styles.formGrid}>
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

            <Field label="布局类型">
              <Dropdown
                value={layoutType}
                selectedOptions={[layoutType]}
                onOptionSelect={(_, data) => setLayoutType(data.optionValue as ImageLayoutType)}
              >
                <Option value="inline">内联</Option>
                <Option value="floating">浮动</Option>
              </Dropdown>
            </Field>
          </div>

          <Switch
            label="保持纵横比"
            checked={keepAspectRatio}
            onChange={(_, data) => setKeepAspectRatio(data.checked)}
          />
        </div>
      </div>

      {/* 浮动图片选项 / Floating image options */}
      {layoutType === "floating" && (
        <div className={styles.formRow}>
          <Label weight="semibold">浮动图片选项</Label>
          <div className={styles.optionsContainer}>
            <Field label="文字环绕方式">
              <Dropdown
                value={wrapType}
                selectedOptions={[wrapType]}
                onOptionSelect={(_, data) => setWrapType(data.optionValue as WrapType)}
              >
                <Option value="Square">四周型</Option>
                <Option value="Tight">紧密型</Option>
                <Option value="Through">穿越型</Option>
                <Option value="TopAndBottom">上下型</Option>
                <Option value="Behind">衬于文字下方</Option>
                <Option value="InFrontOf">浮于文字上方</Option>
              </Dropdown>
            </Field>

            <div className={styles.formGrid}>
              <Field label="水平位置（磅）">
                <Input
                  type="number"
                  value={leftPosition}
                  onChange={(e) => setLeftPosition(e.target.value)}
                  placeholder="0"
                />
              </Field>

              <Field label="垂直位置（磅）">
                <Input
                  type="number"
                  value={topPosition}
                  onChange={(e) => setTopPosition(e.target.value)}
                  placeholder="0"
                />
              </Field>
            </div>

            <Switch
              label="锁定锚点"
              checked={lockAnchor}
              onChange={(_, data) => setLockAnchor(data.checked)}
            />

            <Switch
              label="允许与文字重叠"
              checked={allowOverlap}
              onChange={(_, data) => setAllowOverlap(data.checked)}
            />

            <div style={{ 
              padding: "8px", 
              backgroundColor: tokens.colorPaletteYellowBackground1,
              borderRadius: tokens.borderRadiusSmall,
              fontSize: "12px",
              marginTop: "8px"
            }}>
              注意：Word JavaScript API 对浮动图片的支持有限，某些选项可能无法完全生效。
            </div>
          </div>
        </div>
      )}

      {/* 操作按钮 / Action buttons */}
      <div className={styles.buttonGroup}>
        <Button
          appearance="primary"
          onClick={handleInsertImage}
          disabled={loading || !imageBase64}
          icon={loading ? <Spinner size="tiny" /> : undefined}
        >
          {loading ? "正在插入..." : "插入图片"}
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
