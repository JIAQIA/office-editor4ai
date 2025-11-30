/**
 * 文件名: ImageReplace.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 图片替换工具 UI 组件
 *       支持多种设置图片的方式：本地上传、URL、Base64 文本
 */

import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  Button,
  Field,
  Input,
  Textarea,
  tokens,
  makeStyles,
  RadioGroup,
  Radio,
  Label,
  Card,
  Switch,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from "@fluentui/react-components";
import { replaceImage, getImageInfo, type ImageElementInfo } from "../../../ppt-tools";
import { readImageAsBase64, fetchImageAsBase64 } from "../../../ppt-tools";
import {
  ImageEdit24Regular,
  ArrowUpload24Regular,
  Info24Regular,
  Warning24Regular,
  Checkmark24Regular,
} from "@fluentui/react-icons";

/* global HTMLInputElement, File, console */

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
interface ImageReplaceProps {}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
    padding: "0 8px",
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
  hiddenInput: {
    display: "none",
  },
  fileName: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginTop: "4px",
    textAlign: "center",
  },
  hint: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginBottom: "12px",
    width: "100%",
    textAlign: "center",
    lineHeight: "1.4",
  },
  infoCard: {
    width: "100%",
    padding: "12px",
    marginBottom: "16px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  infoRow: {
    display: "flex",
    justifyContent: "space-between",
    marginBottom: "4px",
    fontSize: tokens.fontSizeBase200,
  },
  infoLabel: {
    color: tokens.colorNeutralForeground3,
  },
  infoValue: {
    color: tokens.colorNeutralForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
  dimensionContainer: {
    display: "flex",
    gap: "12px",
    width: "100%",
    marginTop: "12px",
  },
  dimensionField: {
    flex: 1,
  },
  messageBar: {
    marginBottom: "16px",
  },
});

const ImageReplace: React.FC<ImageReplaceProps> = () => {
  const styles = useStyles();

  // 图片来源类型：upload, url, base64
  const [sourceType, setSourceType] = useState<"upload" | "url" | "base64">("upload");

  // 本地上传相关状态
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [uploadedBase64, setUploadedBase64] = useState<string>("");
  const [previewUrl, setPreviewUrl] = useState<string>("");

  // URL 相关状态
  const [imageUrl, setImageUrl] = useState<string>("");

  // Base64 文本相关状态
  const [base64Text, setBase64Text] = useState<string>("");

  // 尺寸控制
  const [keepDimensions, setKeepDimensions] = useState<boolean>(true);
  const [width, setWidth] = useState<string>("");
  const [height, setHeight] = useState<string>("");

  // 当前选中的图片信息
  const [selectedImageInfo, setSelectedImageInfo] = useState<ImageElementInfo | null>(null);

  // 状态
  const [isReplacing, setIsReplacing] = useState<boolean>(false);
  const [isLoadingInfo, setIsLoadingInfo] = useState<boolean>(false);

  // 消息提示状态
  const [message, setMessage] = useState<{
    text: string;
    intent: "success" | "error" | "warning" | "info";
  } | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // 加载当前选中的图片信息
  const loadSelectedImageInfo = async () => {
    setIsLoadingInfo(true);
    try {
      const info = await getImageInfo();
      setSelectedImageInfo(info);

      // 如果有选中的图片，设置默认尺寸
      if (info) {
        setWidth(info.width.toFixed(1));
        setHeight(info.height.toFixed(1));
      }
    } catch (error) {
      console.error("加载图片信息失败:", error);
      setSelectedImageInfo(null);
    } finally {
      setIsLoadingInfo(false);
    }
  };

  // 组件加载时获取选中的图片信息
  useEffect(() => {
    loadSelectedImageInfo();
  }, []);

  // 处理文件选择
  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // 验证文件类型
    if (!file.type.startsWith("image/")) {
      setMessage({ text: "请选择图片文件", intent: "error" });
      return;
    }

    setSelectedFile(file);

    try {
      // 读取为 Base64
      const base64 = await readImageAsBase64(file);
      setUploadedBase64(base64);
      setPreviewUrl(base64);
    } catch (error) {
      console.error("读取图片失败:", error);
      setMessage({ text: "读取图片失败，请重试", intent: "error" });
    }
  };

  // 触发文件选择
  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  // 处理 Base64 文本变化
  const handleBase64TextChange = (value: string) => {
    setBase64Text(value);

    // 尝试生成预览
    try {
      let previewData = value.trim();
      if (!previewData.startsWith("data:")) {
        // 如果不是 data URL，尝试添加前缀
        previewData = `data:image/png;base64,${previewData}`;
      }
      setPreviewUrl(previewData);
    } catch (error) {
      console.error("生成预览失败:", error);
      setPreviewUrl("");
    }
  };

  // 处理替换图片
  const handleReplaceImage = async () => {
    setIsReplacing(true);

    try {
      // 检查是否选中了图片
      if (!selectedImageInfo) {
        setMessage({ text: "请先选中一个图片元素", intent: "error" });
        await loadSelectedImageInfo();
        return;
      }

      // 获取图片数据
      let imageSource: string;

      if (sourceType === "upload") {
        if (!uploadedBase64) {
          setMessage({ text: "请先选择图片文件", intent: "error" });
          return;
        }
        imageSource = uploadedBase64;
      } else if (sourceType === "url") {
        if (!imageUrl.trim()) {
          setMessage({ text: "请输入图片 URL", intent: "error" });
          return;
        }

        try {
          // 从 URL 加载图片并转换为 Base64
          imageSource = await fetchImageAsBase64(imageUrl.trim());
        } catch (error) {
          console.error("加载图片失败:", error);
          setMessage({
            text: `加载图片失败: ${(error as Error).message}。提示：请确保 URL 可访问且支持 CORS`,
            intent: "error",
          });
          return;
        }
      } else {
        // base64 文本
        if (!base64Text.trim()) {
          setMessage({ text: "请输入 Base64 数据", intent: "error" });
          return;
        }
        imageSource = base64Text.trim();
      }

      // 解析尺寸参数
      const widthValue = width.trim() === "" ? undefined : parseFloat(width);
      const heightValue = height.trim() === "" ? undefined : parseFloat(height);

      // 替换图片
      const result = await replaceImage({
        imageSource,
        keepDimensions,
        width: widthValue,
        height: heightValue,
      });

      if (result.success) {
        setMessage({
          text: `图片替换成功！元素类型: ${result.elementType}，新元素ID: ${result.elementId}，原始尺寸: ${result.originalDimensions?.width.toFixed(1)} × ${result.originalDimensions?.height.toFixed(1)} 磅`,
          intent: "success",
        });

        // 重新加载图片信息
        await loadSelectedImageInfo();
      } else {
        setMessage({ text: `替换失败: ${result.message}`, intent: "error" });
      }
    } catch (error) {
      console.error("替换图片失败:", error);
      setMessage({ text: `替换图片失败: ${(error as Error).message}`, intent: "error" });
    } finally {
      setIsReplacing(false);
    }
  };

  // 判断是否可以替换
  const canReplace = () => {
    if (!selectedImageInfo) return false;

    if (sourceType === "upload") {
      return !!uploadedBase64;
    } else if (sourceType === "url") {
      return !!imageUrl.trim();
    } else {
      return !!base64Text.trim();
    }
  };

  // 判断选中的元素是否是图片
  const isImageElement = () => {
    if (!selectedImageInfo) return false;
    
    // Picture、Image 和 Placeholder 类型都支持
    return (
      selectedImageInfo.elementType === "Picture" ||
      selectedImageInfo.elementType === "Image" ||
      selectedImageInfo.elementType === "Placeholder"
    );
  };

  return (
    <div className={styles.container}>
      {/* 操作结果消息 */}
      {message && (
        <div className={styles.section}>
          <MessageBar intent={message.intent} className={styles.messageBar}>
            <MessageBarBody>{message.text}</MessageBarBody>
          </MessageBar>
        </div>
      )}

      {/* 提示信息 */}
      <div className={styles.section}>
        {isLoadingInfo ? (
          <MessageBar intent="info" className={styles.messageBar}>
            <MessageBarBody>
              <MessageBarTitle>正在加载...</MessageBarTitle>
              正在获取选中元素的信息
            </MessageBarBody>
          </MessageBar>
        ) : selectedImageInfo ? (
          isImageElement() ? (
            <MessageBar
              intent="success"
              className={styles.messageBar}
              icon={<Checkmark24Regular />}
            >
              <MessageBarBody>
                <MessageBarTitle>已选中图片元素</MessageBarTitle>
                类型: {selectedImageInfo.elementType}
                {selectedImageInfo.isPlaceholder && ` (${selectedImageInfo.placeholderType})`}
              </MessageBarBody>
            </MessageBar>
          ) : (
            <MessageBar intent="warning" className={styles.messageBar} icon={<Warning24Regular />}>
              <MessageBarBody>
                <MessageBarTitle>选中的元素不支持图片替换</MessageBarTitle>
                当前选中的是 {selectedImageInfo.elementType} 类型
                {selectedImageInfo.isPlaceholder && ` (${selectedImageInfo.placeholderType})`}
                。请选择图片元素或占位符
              </MessageBarBody>
            </MessageBar>
          )
        ) : (
          <MessageBar intent="info" className={styles.messageBar} icon={<Info24Regular />}>
            <MessageBarBody>
              <MessageBarTitle>使用说明</MessageBarTitle>
              请先在幻灯片中选中要替换的图片，然后选择新图片来源
            </MessageBarBody>
          </MessageBar>
        )}
      </div>

      {/* 当前选中的图片信息 */}
      {selectedImageInfo && isImageElement() && (
        <div className={styles.section}>
          <Card className={styles.infoCard}>
            <Label weight="semibold">当前图片信息</Label>
            <div className={styles.infoRow}>
              <span className={styles.infoLabel}>名称:</span>
              <span className={styles.infoValue}>{selectedImageInfo.name || "(无名称)"}</span>
            </div>
            <div className={styles.infoRow}>
              <span className={styles.infoLabel}>类型:</span>
              <span className={styles.infoValue}>
                {selectedImageInfo.elementType}
                {selectedImageInfo.isPlaceholder && ` (${selectedImageInfo.placeholderType})`}
              </span>
            </div>
            <div className={styles.infoRow}>
              <span className={styles.infoLabel}>位置:</span>
              <span className={styles.infoValue}>
                ({selectedImageInfo.left.toFixed(1)}, {selectedImageInfo.top.toFixed(1)})
              </span>
            </div>
            <div className={styles.infoRow}>
              <span className={styles.infoLabel}>尺寸:</span>
              <span className={styles.infoValue}>
                {selectedImageInfo.width.toFixed(1)} × {selectedImageInfo.height.toFixed(1)} 磅
              </span>
            </div>
          </Card>
        </div>
      )}

      {/* 刷新按钮 */}
      <div className={styles.section}>
        <Button appearance="secondary" onClick={loadSelectedImageInfo} disabled={isLoadingInfo}>
          {isLoadingInfo ? "加载中..." : "刷新选中元素"}
        </Button>
      </div>

      {/* 图片来源类型选择 */}
      <div className={styles.section}>
        <Label weight="semibold">选择新图片来源</Label>
        <RadioGroup
          value={sourceType}
          onChange={(_, data) => setSourceType(data.value as "upload" | "url" | "base64")}
          className={styles.radioGroup}
        >
          <Radio value="upload" label="上传本地图片（推荐）" />
          <Radio value="url" label="使用图片 URL" />
          <Radio value="base64" label="粘贴 Base64 数据" />
        </RadioGroup>
      </div>

      {/* 本地上传区域 */}
      {sourceType === "upload" && (
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
                <ImageEdit24Regular className={styles.uploadIcon} />
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

      {/* Base64 文本输入区域 */}
      {sourceType === "base64" && (
        <div className={styles.section}>
          <Field label="Base64 数据" hint="支持带或不带 data URL 前缀">
            <Textarea
              value={base64Text}
              onChange={(e) => handleBase64TextChange(e.target.value)}
              placeholder="data:image/png;base64,iVBORw0KGgo... 或 iVBORw0KGgo..."
              rows={6}
            />
          </Field>

          {/* 图片预览 */}
          {previewUrl && (
            <div className={styles.previewContainer}>
              <img src={previewUrl} alt="预览" className={styles.previewImage} />
            </div>
          )}
        </div>
      )}

      {/* 尺寸控制 */}
      <div className={styles.section}>
        <Field label="尺寸设置">
          <Switch
            checked={keepDimensions}
            onChange={(e) => setKeepDimensions(e.currentTarget.checked)}
            label={keepDimensions ? "保持原图片尺寸" : "自定义尺寸"}
          />
        </Field>

        {!keepDimensions && (
          <div className={styles.dimensionContainer}>
            <Field className={styles.dimensionField} label="宽度（磅）">
              <Input
                type="number"
                value={width}
                onChange={(e) => setWidth(e.target.value)}
                placeholder="200"
              />
            </Field>
            <Field className={styles.dimensionField} label="高度（磅）">
              <Input
                type="number"
                value={height}
                onChange={(e) => setHeight(e.target.value)}
                placeholder="150"
              />
            </Field>
          </div>
        )}
      </div>

      <div className={styles.hint}>
        💡 提示: 替换图片会保持原图片的位置，并根据设置保持或修改尺寸
        <br />
        支持普通图片（Picture/Image）和所有类型的占位符（Placeholder）
      </div>

      {/* 操作按钮 */}
      <div className={styles.section}>
        <Button
          appearance="primary"
          size="large"
          onClick={handleReplaceImage}
          disabled={isReplacing || !canReplace() || !isImageElement()}
        >
          {isReplacing ? "替换中..." : "确认替换"}
        </Button>
      </div>
    </div>
  );
};

export default ImageReplace;
