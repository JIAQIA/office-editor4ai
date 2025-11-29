/**
 * 文件名: SlideScreenshot.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片截图工具 UI 组件
 */

import * as React from "react";
import { useState } from "react";
import {
  Button,
  Field,
  Input,
  tokens,
  makeStyles,
  Label,
  Card,
  RadioGroup,
  Radio,
  Spinner,
} from "@fluentui/react-components";
import {
  getCurrentSlideScreenshot,
  getSlideScreenshotByPageNumber,
  getAllSlidesScreenshots,
} from "../../../ppt-tools";
import type { SlideScreenshotResult } from "../../../ppt-tools";
import { Camera24Regular, Image24Regular, Copy24Regular } from "@fluentui/react-icons";

/* global alert, console, navigator, document */

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
interface SlideScreenshotProps {}

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
  sizeContainer: {
    display: "flex",
    gap: "12px",
    width: "100%",
    marginTop: "12px",
  },
  sizeField: {
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
  previewContainer: {
    width: "100%",
    marginTop: "16px",
    marginBottom: "16px",
  },
  previewCard: {
    width: "100%",
    padding: "16px",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  previewHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "12px",
  },
  previewTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  previewInfo: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginBottom: "12px",
  },
  previewImageContainer: {
    width: "100%",
    display: "flex",
    justifyContent: "center",
    marginBottom: "12px",
  },
  previewImage: {
    maxWidth: "100%",
    maxHeight: "300px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    justifyContent: "center",
  },
  loadingContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "12px",
    padding: "24px",
  },
  loadingText: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground2,
  },
  allSlidesContainer: {
    width: "100%",
    marginTop: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  slideCard: {
    width: "100%",
    padding: "12px",
  },
  slideHeader: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "8px",
  },
  slideImage: {
    width: "100%",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
  },
});

const SlideScreenshot: React.FC<SlideScreenshotProps> = () => {
  const styles = useStyles();

  // 截图模式：current（当前）、specific（指定页码）、all（所有）
  const [mode, setMode] = useState<"current" | "specific" | "all">("current");

  // 指定页码
  const [pageNumber, setPageNumber] = useState<string>("1");

  // 尺寸设置
  const [width, setWidth] = useState<string>("");
  const [height, setHeight] = useState<string>("");

  // 截图结果
  const [screenshot, setScreenshot] = useState<SlideScreenshotResult | null>(null);
  const [allScreenshots, setAllScreenshots] = useState<SlideScreenshotResult[]>([]);

  // 状态
  const [isCapturing, setIsCapturing] = useState<boolean>(false);

  // 处理截图
  const handleCapture = async () => {
    setIsCapturing(true);
    setScreenshot(null);
    setAllScreenshots([]);

    try {
      // 解析尺寸参数
      const widthValue = width.trim() === "" ? undefined : parseInt(width);
      const heightValue = height.trim() === "" ? undefined : parseInt(height);

      if (mode === "current") {
        // 获取当前幻灯片截图
        const result = await getCurrentSlideScreenshot(widthValue, heightValue);
        setScreenshot(result);
        console.log("当前幻灯片截图获取成功:", result);
      } else if (mode === "specific") {
        // 获取指定页码的截图
        const page = parseInt(pageNumber);
        if (isNaN(page) || page < 1) {
          alert("请输入有效的页码（从 1 开始）");
          return;
        }
        const result = await getSlideScreenshotByPageNumber(page, widthValue, heightValue);
        setScreenshot(result);
        console.log(`第 ${page} 页截图获取成功:`, result);
      } else {
        // 获取所有幻灯片截图
        const results = await getAllSlidesScreenshots(widthValue, heightValue);
        setAllScreenshots(results);
        console.log(`所有幻灯片截图获取成功，共 ${results.length} 张`);
      }
    } catch (error) {
      console.error("获取截图失败:", error);
      alert(`获取截图失败: ${(error as Error).message}`);
    } finally {
      setIsCapturing(false);
    }
  };

  // 复制 Base64 到剪贴板
  const handleCopyBase64 = async (base64: string) => {
    try {
      await navigator.clipboard.writeText(base64);
      alert("Base64 数据已复制到剪贴板");
    } catch (error) {
      console.error("复制失败:", error);
      alert("复制失败，请手动复制");
    }
  };

  // 下载图片
  const handleDownload = (base64: string, slideIndex: number) => {
    const dataUrl = `data:image/png;base64,${base64}`;
    const link = document.createElement("a");
    link.href = dataUrl;
    link.download = `slide-${slideIndex + 1}.png`;
    link.click();
  };

  // 渲染单个截图预览
  const renderScreenshotPreview = (result: SlideScreenshotResult) => {
    const dataUrl = `data:image/png;base64,${result.imageBase64}`;

    return (
      <Card className={styles.previewCard}>
        <div className={styles.previewHeader}>
          <div className={styles.previewTitle}>
            <Image24Regular /> 幻灯片截图
          </div>
        </div>

        <div className={styles.previewInfo}>
          页码: {result.slideIndex + 1} | ID: {result.slideId}
          {result.width && result.height && ` | 尺寸: ${result.width}×${result.height}px`}
        </div>

        <div className={styles.previewImageContainer}>
          <img
            src={dataUrl}
            alt={`幻灯片 ${result.slideIndex + 1}`}
            className={styles.previewImage}
          />
        </div>

        <div className={styles.buttonGroup}>
          <Button
            appearance="secondary"
            icon={<Copy24Regular />}
            onClick={() => handleCopyBase64(result.imageBase64)}
          >
            复制 Base64
          </Button>
          <Button
            appearance="secondary"
            onClick={() => handleDownload(result.imageBase64, result.slideIndex)}
          >
            下载图片
          </Button>
        </div>
      </Card>
    );
  };

  return (
    <div className={styles.container}>
      {/* 截图模式选择 */}
      <div className={styles.section}>
        <Label weight="semibold">选择截图模式</Label>
        <RadioGroup
          value={mode}
          onChange={(_, data) => setMode(data.value as "current" | "specific" | "all")}
          className={styles.radioGroup}
        >
          <Radio value="current" label="当前幻灯片" />
          <Radio value="specific" label="指定页码" />
          <Radio value="all" label="所有幻灯片" />
        </RadioGroup>
      </div>

      {/* 指定页码输入 */}
      {mode === "specific" && (
        <div className={styles.section}>
          <Field label="页码（从 1 开始）">
            <Input
              type="number"
              value={pageNumber}
              onChange={(e) => setPageNumber(e.target.value)}
              placeholder="1"
              min={1}
            />
          </Field>
        </div>
      )}

      {/* 尺寸设置 */}
      <div className={styles.section}>
        <Label weight="semibold">图片尺寸（可选）</Label>
        <div className={styles.sizeContainer}>
          <Field className={styles.sizeField} label="宽度（像素）">
            <Input
              type="number"
              value={width}
              onChange={(e) => setWidth(e.target.value)}
              placeholder="自动"
            />
          </Field>
          <Field className={styles.sizeField} label="高度（像素）">
            <Input
              type="number"
              value={height}
              onChange={(e) => setHeight(e.target.value)}
              placeholder="自动"
            />
          </Field>
        </div>
      </div>

      <div className={styles.hint}>
        💡 提示: 如果不指定尺寸，将使用幻灯片的实际尺寸。
        <br />
        如果只指定宽度或高度，另一维度将自动按比例缩放。
      </div>

      {/* 截图按钮 */}
      <div className={styles.section}>
        <Button
          appearance="primary"
          size="large"
          icon={<Camera24Regular />}
          onClick={handleCapture}
          disabled={isCapturing}
        >
          {isCapturing ? "截图中..." : "开始截图"}
        </Button>
      </div>

      {/* 加载状态 */}
      {isCapturing && (
        <div className={styles.loadingContainer}>
          <Spinner size="large" />
          <div className={styles.loadingText}>
            {mode === "all" ? "正在获取所有幻灯片截图，请稍候..." : "正在获取截图..."}
          </div>
        </div>
      )}

      {/* 单个截图预览 */}
      {!isCapturing && screenshot && (
        <div className={styles.previewContainer}>{renderScreenshotPreview(screenshot)}</div>
      )}

      {/* 所有截图预览 */}
      {!isCapturing && allScreenshots.length > 0 && (
        <div className={styles.allSlidesContainer}>
          <Label weight="semibold">共 {allScreenshots.length} 张幻灯片</Label>
          {allScreenshots.map((result, index) => (
            <Card key={index} className={styles.slideCard}>
              <div className={styles.slideHeader}>幻灯片 {result.slideIndex + 1}</div>
              <div className={styles.previewImageContainer}>
                <img
                  src={`data:image/png;base64,${result.imageBase64}`}
                  alt={`幻灯片 ${result.slideIndex + 1}`}
                  className={styles.slideImage}
                />
              </div>
              <div className={styles.buttonGroup}>
                <Button size="small" onClick={() => handleCopyBase64(result.imageBase64)}>
                  复制
                </Button>
                <Button
                  size="small"
                  onClick={() => handleDownload(result.imageBase64, result.slideIndex)}
                >
                  下载
                </Button>
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
};

export default SlideScreenshot;
