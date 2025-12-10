/**
 * 文件名: ReplaceTextDebug.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/10
 * 最后修改日期: 2025/12/10
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: replaceText 工具的调试组件
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
import { replaceText, type ReplaceTextOptions, type ReplaceTextLocator } from "../../../word-tools";

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
  formatContainer: {
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
});

export const ReplaceTextDebug: React.FC = () => {
  const styles = useStyles();

  const [loading, setLoading] = useState(false);
  const [locatorType, setLocatorType] = useState<"selection" | "search" | "range">("selection");
  const [newText, setNewText] = useState("新文本内容");

  const [searchText, setSearchText] = useState("旧文本");
  const [matchCase, setMatchCase] = useState(false);
  const [matchWholeWord, setMatchWholeWord] = useState(false);
  const [replaceAll, setReplaceAll] = useState(false);

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

  const [applyFormat, setApplyFormat] = useState(false);
  const [fontName, setFontName] = useState("Arial");
  const [fontSize, setFontSize] = useState("14");
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const [color, setColor] = useState("#000000");

  const [result, setResult] = useState<{ success: boolean; message: string; count?: number } | null>(null);

  const handleReplace = async () => {
    if (!newText.trim()) {
      setResult({
        success: false,
        message: "请输入新文本内容",
      });
      return;
    }

    setLoading(true);
    setResult(null);

    try {
      let locator: ReplaceTextLocator;

      switch (locatorType) {
        case "selection":
          locator = { type: "selection" };
          break;

        case "search":
          if (!searchText.trim()) {
            setResult({
              success: false,
              message: "请输入搜索文本",
            });
            setLoading(false);
            return;
          }
          locator = {
            type: "search",
            searchText: searchText.trim(),
            searchOptions: {
              matchCase,
              matchWholeWord,
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

      const options: ReplaceTextOptions = {
        locator,
        newText: newText.trim(),
        format: applyFormat
          ? {
              fontName,
              fontSize: parseFloat(fontSize),
              bold,
              italic,
              color,
            }
          : undefined,
        replaceAll: locatorType === "search" ? replaceAll : undefined,
      };

      const replaceResult = await replaceText(options);

      if (replaceResult.success) {
        setResult({
          success: true,
          message: `成功替换 ${replaceResult.count} 处文本`,
          count: replaceResult.count,
        });
      } else {
        setResult({
          success: false,
          message: replaceResult.error || "替换失败",
        });
      }
    } catch (error) {
      console.error("替换文本失败:", error);
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
            onChange={(_, data) => setLocatorType(data.value as "selection" | "search" | "range")}
          >
            <option value="selection">当前选区</option>
            <option value="search">搜索匹配</option>
            <option value="range">指定范围</option>
          </Select>
        </div>

        {locatorType === "search" && (
          <>
            <div className={styles.formRow}>
              <Label>搜索文本</Label>
              <Input value={searchText} onChange={(_, data) => setSearchText(data.value)} placeholder="要查找的文本" />
            </div>
            <div className={styles.formRow}>
              <Switch checked={matchCase} onChange={(_, data) => setMatchCase(data.checked)} label="区分大小写" />
            </div>
            <div className={styles.formRow}>
              <Switch
                checked={matchWholeWord}
                onChange={(_, data) => setMatchWholeWord(data.checked)}
                label="全字匹配"
              />
            </div>
            <div className={styles.formRow}>
              <Switch checked={replaceAll} onChange={(_, data) => setReplaceAll(data.checked)} label="替换所有匹配项" />
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
      </div>

      <Divider />

      <div className={styles.section}>
        <div className={styles.formRow}>
          <Label weight="semibold">新文本内容</Label>
          <Textarea value={newText} onChange={(_, data) => setNewText(data.value)} placeholder="输入新文本..." rows={4} />
        </div>
      </div>

      <Divider />

      <div className={styles.section}>
        <div className={styles.formRow}>
          <Switch checked={applyFormat} onChange={(_, data) => setApplyFormat(data.checked)} label="应用自定义格式" />
        </div>

        {applyFormat && (
          <div className={styles.formatContainer}>
            <div className={styles.formRow}>
              <Label>字体</Label>
              <Input value={fontName} onChange={(_, data) => setFontName(data.value)} placeholder="Arial" />
            </div>

            <div className={styles.formRow}>
              <Label>字号</Label>
              <Input type="number" value={fontSize} onChange={(_, data) => setFontSize(data.value)} placeholder="14" />
            </div>

            <div className={styles.formRow}>
              <Switch checked={bold} onChange={(_, data) => setBold(data.checked)} label="加粗" />
            </div>

            <div className={styles.formRow}>
              <Switch checked={italic} onChange={(_, data) => setItalic(data.checked)} label="斜体" />
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
