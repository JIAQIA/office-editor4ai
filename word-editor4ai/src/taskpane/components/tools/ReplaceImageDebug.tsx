/**
 * 文件名: ReplaceImageDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: replaceImage 工具的调试组件
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
  Select,
  Card,
  CardHeader,
  Divider,
} from "@fluentui/react-components";
import { replaceImage, type ReplaceImageOptions, type ReplaceImageLocator } from "../../../word-tools/replaceImage";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
  },
  section: {
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
    gap: "8px",
  },
  inlineRow: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-end",
  },
  propertiesContainer: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "8px",
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    marginTop: "12px",
  },
  resultCard: {
    width: "100%",
  },
  resultMessage: {
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
  },
  success: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground1,
  },
  error: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground1,
  },
  imagePreview: {
    maxWidth: "100%",
    maxHeight: "200px",
    marginTop: "8px",
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
});

export const ReplaceImageDebug: React.FC = () => {
  const styles = useStyles();

  const [loading, setLoading] = useState(false);
  const [locatorType, setLocatorType] = useState<"selection" | "index" | "search" | "range">("selection");

  const [imageIndex, setImageIndex] = useState("0");

  const [searchAltText, setSearchAltText] = useState("");
  const [searchMinWidth, setSearchMinWidth] = useState("");
  const [searchMaxWidth, setSearchMaxWidth] = useState("");
  const [searchMinHeight, setSearchMinHeight] = useState("");
  const [searchMaxHeight, setSearchMaxHeight] = useState("");

  const [rangeType, setRangeType] = useState<"bookmark" | "heading" | "paragraph" | "section" | "contentControl">(
    "paragraph"
  );
  const [bookmarkName, setBookmarkName] = useState("myBookmark");
  const [headingText, setHeadingText] = useState("");
  const [headingLevel, setHeadingLevel] = useState("1");
  const [headingIndex, setHeadingIndex] = useState("0");
  const [paragraphStart, setParagraphStart] = useState("0");
  const [paragraphEnd, setParagraphEnd] = useState("");
  const [sectionIndex, setSectionIndex] = useState("0");
  const [controlTitle, setControlTitle] = useState("");
  const [controlTag, setControlTag] = useState("");
  const [controlIndex, setControlIndex] = useState("0");

  const [replaceAll, setReplaceAll] = useState(false);

  const [useNewImage, setUseNewImage] = useState(false);
  const [imageBase64, setImageBase64] = useState("");
  const [imagePreview, setImagePreview] = useState("");

  const [applyProperties, setApplyProperties] = useState(true);
  const [width, setWidth] = useState("200");
  const [height, setHeight] = useState("150");
  const [altText, setAltText] = useState("示例图片");
  const [hyperlink, setHyperlink] = useState("");
  const [lockAspectRatio, setLockAspectRatio] = useState(true);

  const [result, setResult] = useState<{ success: boolean; message: string; count?: number } | null>(null);

  const handleImageUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const dataUrl = e.target?.result as string;
        // 移除 Data URL 前缀，只保留纯 Base64 字符串 / Remove Data URL prefix, keep only pure Base64 string
        const base64 = dataUrl.split(',')[1];
        setImageBase64(base64);
        setImagePreview(dataUrl); // 预览仍使用完整的 Data URL / Preview still uses full Data URL
      };
      reader.readAsDataURL(file);
    }
  };

  const handleReplace = async () => {
    if (!useNewImage && !applyProperties) {
      setResult({
        success: false,
        message: "请至少选择替换图片或更新属性中的一项",
      });
      return;
    }

    if (useNewImage && !imageBase64) {
      setResult({
        success: false,
        message: "请上传图片",
      });
      return;
    }

    setLoading(true);
    setResult(null);

    try {
      let locator: ReplaceImageLocator;

      switch (locatorType) {
        case "selection":
          locator = { type: "selection" };
          break;

        case "index":
          locator = {
            type: "index",
            index: parseInt(imageIndex),
          };
          break;

        case "search":
          locator = {
            type: "search",
            searchOptions: {
              altText: searchAltText.trim() || undefined,
              minWidth: searchMinWidth ? parseFloat(searchMinWidth) : undefined,
              maxWidth: searchMaxWidth ? parseFloat(searchMaxWidth) : undefined,
              minHeight: searchMinHeight ? parseFloat(searchMinHeight) : undefined,
              maxHeight: searchMaxHeight ? parseFloat(searchMaxHeight) : undefined,
            },
          };
          break;

        case "range":
          switch (rangeType) {
            case "bookmark":
              if (!bookmarkName.trim()) {
                setResult({
                  success: false,
                  message: "请输入书签名称",
                });
                setLoading(false);
                return;
              }
              locator = {
                type: "range",
                rangeLocator: {
                  type: "bookmark",
                  name: bookmarkName.trim(),
                },
              };
              break;

            case "heading":
              locator = {
                type: "range",
                rangeLocator: {
                  type: "heading",
                  text: headingText.trim() || undefined,
                  level: headingLevel ? parseInt(headingLevel) : undefined,
                  index: headingIndex ? parseInt(headingIndex) : undefined,
                },
              };
              break;

            case "paragraph":
              locator = {
                type: "range",
                rangeLocator: {
                  type: "paragraph",
                  startIndex: parseInt(paragraphStart),
                  endIndex: paragraphEnd ? parseInt(paragraphEnd) : undefined,
                },
              };
              break;

            case "section":
              locator = {
                type: "range",
                rangeLocator: {
                  type: "section",
                  index: parseInt(sectionIndex),
                },
              };
              break;

            case "contentControl":
              locator = {
                type: "range",
                rangeLocator: {
                  type: "contentControl",
                  title: controlTitle.trim() || undefined,
                  tag: controlTag.trim() || undefined,
                  index: controlIndex ? parseInt(controlIndex) : undefined,
                },
              };
              break;
          }
          break;
      }

      const options: ReplaceImageOptions = {
        locator,
        newImageData: useNewImage ? imageBase64 : undefined,
        properties: applyProperties
          ? {
              width: width ? parseFloat(width) : undefined,
              height: height ? parseFloat(height) : undefined,
              altText: altText.trim() || undefined,
              hyperlink: hyperlink.trim() || undefined,
              lockAspectRatio,
            }
          : undefined,
        replaceAll: locatorType === "search" || locatorType === "range" ? replaceAll : undefined,
      };

      const replaceResult = await replaceImage(options);

      if (replaceResult.success) {
        setResult({
          success: true,
          message: `成功替换 ${replaceResult.count} 张图片`,
          count: replaceResult.count,
        });
      } else {
        setResult({
          success: false,
          message: replaceResult.error || "替换失败",
        });
      }
    } catch (error) {
      console.error("替换图片失败:", error);
      setResult({
        success: false,
        message: `替换失败: ${error instanceof Error ? error.message : String(error)}`,
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.section}>
        <div className={styles.formRow}>
          <Label weight="semibold">定位方式</Label>
          <Select
            value={locatorType}
            onChange={(_, data) => setLocatorType(data.value as "selection" | "index" | "search" | "range")}
          >
            <option value="selection">当前选区</option>
            <option value="index">按索引</option>
            <option value="search">搜索匹配</option>
            <option value="range">指定范围</option>
          </Select>
        </div>

        {locatorType === "index" && (
          <div className={styles.formRow}>
            <Label>图片索引</Label>
            <Input
              type="number"
              value={imageIndex}
              onChange={(_, data) => setImageIndex(data.value)}
              placeholder="0"
            />
          </div>
        )}

        {locatorType === "search" && (
          <>
            <div className={styles.formRow}>
              <Label>替代文本（可选）</Label>
              <Input
                value={searchAltText}
                onChange={(_, data) => setSearchAltText(data.value)}
                placeholder="部分匹配"
              />
            </div>
            <div className={styles.inlineRow}>
              <div style={{ flex: 1 }}>
                <Label>最小宽度（可选）</Label>
                <Input
                  type="number"
                  value={searchMinWidth}
                  onChange={(_, data) => setSearchMinWidth(data.value)}
                  placeholder="磅"
                />
              </div>
              <div style={{ flex: 1 }}>
                <Label>最大宽度（可选）</Label>
                <Input
                  type="number"
                  value={searchMaxWidth}
                  onChange={(_, data) => setSearchMaxWidth(data.value)}
                  placeholder="磅"
                />
              </div>
            </div>
            <div className={styles.inlineRow}>
              <div style={{ flex: 1 }}>
                <Label>最小高度（可选）</Label>
                <Input
                  type="number"
                  value={searchMinHeight}
                  onChange={(_, data) => setSearchMinHeight(data.value)}
                  placeholder="磅"
                />
              </div>
              <div style={{ flex: 1 }}>
                <Label>最大高度（可选）</Label>
                <Input
                  type="number"
                  value={searchMaxHeight}
                  onChange={(_, data) => setSearchMaxHeight(data.value)}
                  placeholder="磅"
                />
              </div>
            </div>
          </>
        )}

        {locatorType === "range" && (
          <>
            <div className={styles.formRow}>
              <Label>范围类型</Label>
              <Select
                value={rangeType}
                onChange={(_, data) =>
                  setRangeType(data.value as "bookmark" | "heading" | "paragraph" | "section" | "contentControl")
                }
              >
                <option value="bookmark">书签</option>
                <option value="heading">标题</option>
                <option value="paragraph">段落</option>
                <option value="section">节</option>
                <option value="contentControl">内容控件</option>
              </Select>
            </div>

            {rangeType === "bookmark" && (
              <div className={styles.formRow}>
                <Label>书签名称</Label>
                <Input value={bookmarkName} onChange={(_, data) => setBookmarkName(data.value)} />
              </div>
            )}

            {rangeType === "heading" && (
              <>
                <div className={styles.formRow}>
                  <Label>标题文本（可选）</Label>
                  <Input
                    value={headingText}
                    onChange={(_, data) => setHeadingText(data.value)}
                    placeholder="部分匹配"
                  />
                </div>
                <div className={styles.inlineRow}>
                  <div style={{ flex: 1 }}>
                    <Label>标题级别（可选）</Label>
                    <Input
                      type="number"
                      value={headingLevel}
                      onChange={(_, data) => setHeadingLevel(data.value)}
                      placeholder="1-9"
                    />
                  </div>
                  <div style={{ flex: 1 }}>
                    <Label>索引（可选）</Label>
                    <Input
                      type="number"
                      value={headingIndex}
                      onChange={(_, data) => setHeadingIndex(data.value)}
                      placeholder="0"
                    />
                  </div>
                </div>
              </>
            )}

            {rangeType === "paragraph" && (
              <div className={styles.inlineRow}>
                <div style={{ flex: 1 }}>
                  <Label>起始段落</Label>
                  <Input
                    type="number"
                    value={paragraphStart}
                    onChange={(_, data) => setParagraphStart(data.value)}
                    placeholder="0"
                  />
                </div>
                <div style={{ flex: 1 }}>
                  <Label>结束段落（可选）</Label>
                  <Input
                    type="number"
                    value={paragraphEnd}
                    onChange={(_, data) => setParagraphEnd(data.value)}
                    placeholder="留空则单段"
                  />
                </div>
              </div>
            )}

            {rangeType === "section" && (
              <div className={styles.formRow}>
                <Label>节索引</Label>
                <Input
                  type="number"
                  value={sectionIndex}
                  onChange={(_, data) => setSectionIndex(data.value)}
                  placeholder="0"
                />
              </div>
            )}

            {rangeType === "contentControl" && (
              <>
                <div className={styles.formRow}>
                  <Label>控件标题（可选）</Label>
                  <Input value={controlTitle} onChange={(_, data) => setControlTitle(data.value)} />
                </div>
                <div className={styles.formRow}>
                  <Label>控件标签（可选）</Label>
                  <Input value={controlTag} onChange={(_, data) => setControlTag(data.value)} />
                </div>
                <div className={styles.formRow}>
                  <Label>控件索引（可选）</Label>
                  <Input
                    type="number"
                    value={controlIndex}
                    onChange={(_, data) => setControlIndex(data.value)}
                    placeholder="0"
                  />
                </div>
              </>
            )}
          </>
        )}

        {(locatorType === "search" || locatorType === "range") && (
          <div className={styles.formRow}>
            <Switch checked={replaceAll} onChange={(_, data) => setReplaceAll(data.checked)} label="替换所有匹配项" />
          </div>
        )}
      </div>

      <Divider />

      <div className={styles.section}>
        <div className={styles.formRow}>
          <Switch checked={useNewImage} onChange={(_, data) => setUseNewImage(data.checked)} label="替换图片内容" />
        </div>

        {useNewImage && (
          <>
            <div className={styles.formRow}>
              <Label>上传图片</Label>
              <input
                type="file"
                accept="image/*"
                onChange={handleImageUpload}
                style={{
                  padding: "8px",
                  border: `1px solid ${tokens.colorNeutralStroke1}`,
                  borderRadius: tokens.borderRadiusSmall,
                  cursor: "pointer",
                }}
              />
            </div>
            {imagePreview && <img src={imagePreview} alt="预览" className={styles.imagePreview} />}
          </>
        )}
      </div>

      <Divider />

      <div className={styles.section}>
        <div className={styles.formRow}>
          <Switch
            checked={applyProperties}
            onChange={(_, data) => setApplyProperties(data.checked)}
            label="更新图片属性"
          />
        </div>

        {applyProperties && (
          <div className={styles.propertiesContainer}>
            <div className={styles.inlineRow}>
              <div style={{ flex: 1 }}>
                <Label>宽度（磅）</Label>
                <Input type="number" value={width} onChange={(_, data) => setWidth(data.value)} placeholder="200" />
              </div>
              <div style={{ flex: 1 }}>
                <Label>高度（磅）</Label>
                <Input type="number" value={height} onChange={(_, data) => setHeight(data.value)} placeholder="150" />
              </div>
            </div>

            <div className={styles.formRow}>
              <Label>替代文本</Label>
              <Input value={altText} onChange={(_, data) => setAltText(data.value)} placeholder="图片描述" />
            </div>

            <div className={styles.formRow}>
              <Label>超链接（可选）</Label>
              <Input
                value={hyperlink}
                onChange={(_, data) => setHyperlink(data.value)}
                placeholder="https://example.com"
              />
            </div>

            <div className={styles.formRow}>
              <Switch
                checked={lockAspectRatio}
                onChange={(_, data) => setLockAspectRatio(data.checked)}
                label="锁定纵横比"
              />
            </div>
          </div>
        )}
      </div>

      <div className={styles.buttonGroup}>
        <Button appearance="primary" onClick={handleReplace} disabled={loading}>
          {loading ? <Spinner size="tiny" /> : "执行替换"}
        </Button>
      </div>

      {result && (
        <Card className={styles.resultCard}>
          <CardHeader header={result.success ? "✅ 成功" : "❌ 失败"} />
          <div className={`${styles.resultMessage} ${result.success ? styles.success : styles.error}`}>
            {result.message}
          </div>
        </Card>
      )}
    </div>
  );
};
