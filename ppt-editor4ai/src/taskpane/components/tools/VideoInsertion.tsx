/**
 * 文件名: VideoInsertion.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 视频插入工具 UI 组件
 * 
 * ⚠️ 功能状态：不可用
 * 
 * PowerPoint JavaScript API 目前不支持通过 ShapeCollection 插入媒体元素（视频/音频）。
 * 详情：https://techcommunity.microsoft.com/idea/microsoft365developerplatform/support-for-inserting-media-elements-via-powerpoint-shapecollection-addmediaelem/4404793
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
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Link,
} from "@fluentui/react-components";
import { insertVideoToSlide, readVideoAsBase64, fetchVideoAsBase64 } from "../../../ppt-tools";
import { Video24Regular, ArrowUpload24Regular, DismissRegular, ErrorCircle24Regular } from "@fluentui/react-icons";

/* global HTMLInputElement, File, console */

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
interface VideoInsertionProps {}

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
  previewVideo: {
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
  warning: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteYellowForeground1,
    marginBottom: "12px",
    width: "100%",
    textAlign: "center",
    lineHeight: "1.4",
    padding: "8px",
    backgroundColor: tokens.colorPaletteYellowBackground1,
    borderRadius: tokens.borderRadiusMedium,
  },
  unavailableNotice: {
    width: "100%",
    marginBottom: "16px",
    padding: "16px",
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderRadius: tokens.borderRadiusMedium,
    border: `2px solid ${tokens.colorPaletteRedBorder1}`,
  },
  unavailableTitle: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorPaletteRedForeground1,
    marginBottom: "12px",
  },
  unavailableContent: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    lineHeight: "1.6",
    marginBottom: "12px",
  },
  unavailableList: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    lineHeight: "1.6",
    paddingLeft: "20px",
    marginBottom: "12px",
  },
  disabledOverlay: {
    opacity: 0.5,
    pointerEvents: "none",
  },
  messageBar: {
    marginBottom: "12px",
    width: "100%",
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

const VideoInsertion: React.FC<VideoInsertionProps> = () => {
  const styles = useStyles();

  // 视频来源类型：base64 或 url
  const [sourceType, setSourceType] = useState<"base64" | "url">("base64");

  // Base64 相关状态
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [base64Data, setBase64Data] = useState<string>("");
  const [previewUrl, setPreviewUrl] = useState<string>("");

  // URL 相关状态
  const [videoUrl, setVideoUrl] = useState<string>("");

  // 位置和尺寸
  const [left, setLeft] = useState<string>("");
  const [top, setTop] = useState<string>("");
  const [width, setWidth] = useState<string>("");
  const [height, setHeight] = useState<string>("");

  // 状态
  const [isInserting, setIsInserting] = useState<boolean>(false);
  const [message, setMessage] = useState<{ type: "success" | "error" | "warning" | "info"; title: string; content: string } | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // 处理文件选择
  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // 验证文件类型
    if (!file.type.startsWith("video/")) {
      setMessage({ type: "error", title: "文件类型错误", content: "请选择视频文件" });
      return;
    }

    setSelectedFile(file);

    try {
      // 读取为 Base64
      const base64 = await readVideoAsBase64(file);
      setBase64Data(base64);
      setPreviewUrl(base64);
    } catch (error) {
      console.error("读取视频失败:", error);
      setMessage({ type: "error", title: "读取失败", content: "读取视频失败，请重试" });
    }
  };

  // 触发文件选择
  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  // 处理插入视频
  const handleInsertVideo = async () => {
    setIsInserting(true);

    try {
      // 解析位置和尺寸参数
      const leftValue = left.trim() === "" ? undefined : parseFloat(left);
      const topValue = top.trim() === "" ? undefined : parseFloat(top);
      const widthValue = width.trim() === "" ? undefined : parseFloat(width);
      const heightValue = height.trim() === "" ? undefined : parseFloat(height);

      let videoSource: string;

      if (sourceType === "base64") {
        if (!base64Data) {
          setMessage({ type: "warning", title: "未选择文件", content: "请先选择视频文件" });
          return;
        }
        videoSource = base64Data;
      } else {
        // URL 方式：先转换为 Base64
        if (!videoUrl.trim()) {
          setMessage({ type: "warning", title: "未输入 URL", content: "请输入视频 URL" });
          return;
        }
        
        try {
          // 从 URL 加载视频并转换为 Base64
          videoSource = await fetchVideoAsBase64(videoUrl.trim());
        } catch (error) {
          console.error("加载视频失败:", error);
          setMessage({ 
            type: "error", 
            title: "加载失败", 
            content: `加载视频失败: ${(error as Error).message}。提示：请确保 URL 可访问且支持 CORS` 
          });
          return;
        }
      }

      // 插入视频（videoSource 已经是 Base64 格式）
      const result = await insertVideoToSlide({
        videoSource,
        left: leftValue,
        top: topValue,
        width: widthValue,
        height: heightValue,
      });

      setMessage({ 
        type: "success", 
        title: "插入成功", 
        content: `视频已插入！尺寸: ${result.width.toFixed(1)} × ${result.height.toFixed(1)} 磅` 
      });

      // 清空表单（可选）
      // _resetForm();
    } catch (error) {
      console.error("插入视频失败:", error);
      setMessage({ 
        type: "error", 
        title: "插入失败", 
        content: `${(error as Error).message}` 
      });
    } finally {
      setIsInserting(false);
    }
  };

  // 重置表单
  const _resetForm = () => {
    setSelectedFile(null);
    setBase64Data("");
    setPreviewUrl("");
    setVideoUrl("");
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
      {/* 功能不可用通知 */}
      <div className={styles.unavailableNotice}>
        <div className={styles.unavailableTitle}>
          <ErrorCircle24Regular />
          功能不可用
        </div>
        <div className={styles.unavailableContent}>
          PowerPoint JavaScript API 目前不支持通过 ShapeCollection 插入媒体元素（视频/音频）。
          这是 Microsoft Office JavaScript API 的已知限制。
        </div>
        <div className={styles.unavailableContent}>
          <strong>官方功能请求：</strong>
          <br />
          <Link 
            href="https://techcommunity.microsoft.com/idea/microsoft365developerplatform/support-for-inserting-media-elements-via-powerpoint-shapecollection-addmediaelem/4404793"
            target="_blank"
          >
            Support for inserting media elements via PowerPoint ShapeCollection
          </Link>
        </div>
        <div className={styles.unavailableContent}>
          <strong>替代方案：</strong>
        </div>
        <ul className={styles.unavailableList}>
          <li>使用 PowerPoint 桌面版手动插入视频</li>
          <li>使用在线视频嵌入（YouTube, Microsoft Stream）</li>
          <li>等待 Microsoft 官方 API 支持</li>
        </ul>
      </div>

      {/* 消息提示 */}
      {message && (
        <MessageBar
          key={message.type + message.title}
          intent={message.type}
          className={styles.messageBar}
        >
          <MessageBarBody>
            <MessageBarTitle>{message.title}</MessageBarTitle>
            {message.content}
          </MessageBarBody>
        </MessageBar>
      )}

      {/* 禁用的表单区域 */}
      <div className={styles.disabledOverlay}>

      {/* 视频来源类型选择 */}
      <div className={styles.section}>
        <Label weight="semibold">选择视频来源</Label>
        <RadioGroup
          value={sourceType}
          onChange={(_, data) => setSourceType(data.value as "base64" | "url")}
          className={styles.radioGroup}
        >
          <Radio value="base64" label="上传本地视频（推荐）" />
          <Radio value="url" label="使用视频 URL" />
        </RadioGroup>
      </div>

      {/* Base64 上传区域 */}
      {sourceType === "base64" && (
        <div className={styles.section}>
          <input
            ref={fileInputRef}
            type="file"
            accept="video/*"
            onChange={handleFileSelect}
            className={styles.hiddenInput}
          />
          <Card
            className={`${styles.uploadCard} ${selectedFile ? styles.uploadCardActive : ""}`}
            onClick={handleUploadClick}
          >
            <div className={styles.uploadContent}>
              {selectedFile ? (
                <Video24Regular className={styles.uploadIcon} />
              ) : (
                <ArrowUpload24Regular className={styles.uploadIcon} />
              )}
              <div className={styles.uploadText}>
                {selectedFile ? "点击更换视频" : "点击选择视频文件"}
              </div>
              {selectedFile && <div className={styles.fileName}>{selectedFile.name}</div>}
            </div>
          </Card>

          {/* 视频预览 */}
          {previewUrl && (
            <div className={styles.previewContainer}>
              <video src={previewUrl} controls className={styles.previewVideo} />
            </div>
          )}
        </div>
      )}

      {/* URL 输入区域 */}
      {sourceType === "url" && (
        <div className={styles.section}>
          <Field label="视频 URL">
            <Input
              value={videoUrl}
              onChange={(e) => setVideoUrl(e.target.value)}
              placeholder="https://example.com/video.mp4"
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
                placeholder="默认 400"
              />
            </Field>
            <Field className={styles.positionField} label="高度">
              <Input
                type="number"
                value={height}
                onChange={(e) => setHeight(e.target.value)}
                placeholder="默认 300"
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
          onClick={handleInsertVideo}
          disabled={isInserting || (sourceType === "base64" && !base64Data) || (sourceType === "url" && !videoUrl)}
        >
          {isInserting ? "插入中..." : "确认插入"}
        </Button>
      </div>
      </div>
    </div>
  );
};

export default VideoInsertion;
