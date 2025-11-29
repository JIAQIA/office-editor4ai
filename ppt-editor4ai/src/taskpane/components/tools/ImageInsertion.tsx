/**
 * 文件名: ImageInsertion.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 图片插入工具 UI 组件
 */

import * as React from "react";
import { useState, useRef } from "react";
import {
  Button,
  Field,
  Input,
  tokens,
  makeStyles,
  RadioGroup,
  Radio,
  Label,
  Card,
} from "@fluentui/react-components";
import { insertImageToSlide, readImageAsBase64, fetchImageAsBase64 } from "../../../ppt-tools";
import { Image24Regular, ArrowUpload24Regular } from "@fluentui/react-icons";

/* global HTMLInputElement */

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
interface ImageInsertionProps {}

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
  radioGroup: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  uploadCard: {
    width: "100%",
    padding: "16px",
    marginBottom: "16px",
    cursor: "pointer",
    border: `2px dashed ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    transition: "all 0.2s ease",
    ":hover": {
      border: `2px dashed ${tokens.colorBrandStroke1}`,
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  uploadCardActive: {
    border: `2px dashed ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1Selected,
  },
  uploadContent: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "8px",
  },
  uploadIcon: {
    fontSize: "32px",
    color: tokens.colorBrandForeground1,
  },
  uploadText: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground2,
  },
  previewContainer: {
    width: "100%",
    marginTop: "12px",
    marginBottom: "12px",
    display: "flex",
    justifyContent: "center",
  },
  previewImage: {
    maxWidth: "100%",
    maxHeight: "200px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
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
  hint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginBottom: "12px",
    width: "100%",
    textAlign: "center",
    lineHeight: "1.4",
  },
  hiddenInput: {
    display: "none",
  },
  fileName: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginTop: "4px",
    textAlign: "center",
  },
});

const ImageInsertion: React.FC<ImageInsertionProps> = () => {
  const styles = useStyles();

  // 图片来源类型：base64 或 url
  const [sourceType, setSourceType] = useState<"base64" | "url">("base64");

  // Base64 相关状态
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [base64Data, setBase64Data] = useState<string>("");
  const [previewUrl, setPreviewUrl] = useState<string>("");

  // URL 相关状态
  const [imageUrl, setImageUrl] = useState<string>("");

  // 位置和尺寸
  const [left, setLeft] = useState<string>("");
  const [top, setTop] = useState<string>("");
  const [width, setWidth] = useState<string>("");
  const [height, setHeight] = useState<string>("");

  // 状态
  const [isInserting, setIsInserting] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // 处理文件选择
  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // 验证文件类型
    if (!file.type.startsWith("image/")) {
      alert("请选择图片文件");
      return;
    }

    setSelectedFile(file);

    try {
      // 读取为 Base64
      const base64 = await readImageAsBase64(file);
      setBase64Data(base64);
      setPreviewUrl(base64);
    } catch (error) {
      console.error("读取图片失败:", error);
      alert("读取图片失败，请重试");
    }
  };

  // 触发文件选择
  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  // 处理插入图片
  const handleInsertImage = async () => {
    setIsInserting(true);

    try {
      // 解析位置和尺寸参数
      const leftValue = left.trim() === "" ? undefined : parseFloat(left);
      const topValue = top.trim() === "" ? undefined : parseFloat(top);
      const widthValue = width.trim() === "" ? undefined : parseFloat(width);
      const heightValue = height.trim() === "" ? undefined : parseFloat(height);

      let imageSource: string;

      if (sourceType === "base64") {
        if (!base64Data) {
          alert("请先选择图片文件");
          return;
        }
        imageSource = base64Data;
      } else {
        // URL 方式：先转换为 Base64
        if (!imageUrl.trim()) {
          alert("请输入图片 URL");
          return;
        }
        
        try {
          // 从 URL 加载图片并转换为 Base64
          imageSource = await fetchImageAsBase64(imageUrl.trim());
        } catch (error) {
          console.error("加载图片失败:", error);
          alert(`加载图片失败: ${(error as Error).message}\n\n提示：请确保 URL 可访问且支持 CORS`);
          return;
        }
      }

      // 插入图片（统一使用 base64 方式）
      const result = await insertImageToSlide({
        imageSource,
        sourceType: "base64", // 统一使用 base64，因为 URL 已经转换了
        left: leftValue,
        top: topValue,
        width: widthValue,
        height: heightValue,
      });

      alert(`图片插入成功！\nID: ${result.shapeId}\n尺寸: ${result.width.toFixed(1)} × ${result.height.toFixed(1)} 磅`);

      // 清空表单（可选）
      // resetForm();
    } catch (error) {
      console.error("插入图片失败:", error);
      alert(`插入图片失败: ${(error as Error).message}`);
    } finally {
      setIsInserting(false);
    }
  };

  // 重置表单
  const resetForm = () => {
    setSelectedFile(null);
    setBase64Data("");
    setPreviewUrl("");
    setImageUrl("");
    setLeft("");
    setTop("");
    setWidth("");
    setHeight("");
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  return (
    <div className={styles.container}>
      {/* 图片来源类型选择 */}
      <div className={styles.section}>
        <Label weight="semibold">选择图片来源</Label>
        <RadioGroup
          value={sourceType}
          onChange={(_, data) => setSourceType(data.value as "base64" | "url")}
          className={styles.radioGroup}
        >
          <Radio value="base64" label="上传本地图片（推荐）" />
          <Radio value="url" label="使用图片 URL" />
        </RadioGroup>
      </div>

      {/* Base64 上传区域 */}
      {sourceType === "base64" && (
        <div className={styles.section}>
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            onChange={handleFileSelect}
            className={styles.hiddenInput}
          />
          <Card
            className={`${styles.uploadCard} ${selectedFile ? styles.uploadCardActive : ""}`}
            onClick={handleUploadClick}
          >
            <div className={styles.uploadContent}>
              {selectedFile ? (
                <Image24Regular className={styles.uploadIcon} />
              ) : (
                <ArrowUpload24Regular className={styles.uploadIcon} />
              )}
              <div className={styles.uploadText}>
                {selectedFile ? "点击更换图片" : "点击选择图片文件"}
              </div>
              {selectedFile && <div className={styles.fileName}>{selectedFile.name}</div>}
            </div>
          </Card>

          {/* 图片预览 */}
          {previewUrl && (
            <div className={styles.previewContainer}>
              <img src={previewUrl} alt="预览" className={styles.previewImage} />
            </div>
          )}
        </div>
      )}

      {/* URL 输入区域 */}
      {sourceType === "url" && (
        <div className={styles.section}>
          <Field label="图片 URL">
            <Input
              value={imageUrl}
              onChange={(e) => setImageUrl(e.target.value)}
              placeholder="https://example.com/image.png"
            />
          </Field>
        </div>
      )}

      {/* 位置和尺寸设置 */}
      <div className={styles.section}>
        <Label weight="semibold">位置和尺寸（可选）</Label>
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
                placeholder="默认 200"
              />
            </Field>
            <Field className={styles.positionField} label="高度">
              <Input
                type="number"
                value={height}
                onChange={(e) => setHeight(e.target.value)}
                placeholder="默认 150"
              />
            </Field>
          </div>
        </div>
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
          onClick={handleInsertImage}
          disabled={isInserting || (sourceType === "base64" && !base64Data) || (sourceType === "url" && !imageUrl)}
        >
          {isInserting ? "插入中..." : "确认插入"}
        </Button>
      </div>
    </div>
  );
};

export default ImageInsertion;
