/**
 * 文件名: Comments.tsx
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 获取批注内容的工具组件
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
  Card,
  CardHeader,
  Divider,
  Input,
  Badge,
} from "@fluentui/react-components";
import { getComments, type CommentInfo, type GetCommentsOptions } from "../../../word-tools";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
    gap: "16px",
    padding: "8px",
  },
  optionsContainer: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    marginBottom: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  optionRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  button: {
    width: "100%",
    marginTop: "8px",
  },
  resultContainer: {
    width: "100%",
    marginTop: "16px",
  },
  resultCard: {
    marginBottom: "12px",
    width: "100%",
  },
  cardContent: {
    padding: "12px",
  },
  commentHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "8px",
  },
  commentIcon: {
    fontSize: "24px",
  },
  commentTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    flex: 1,
  },
  metadataGrid: {
    display: "grid",
    gridTemplateColumns: "auto 1fr",
    gap: "8px",
    marginBottom: "12px",
  },
  metadataLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
  },
  metadataValue: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    wordBreak: "break-word",
  },
  commentContent: {
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    marginBottom: "8px",
  },
  associatedText: {
    padding: "8px",
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    marginBottom: "8px",
    borderLeft: `3px solid ${tokens.colorBrandBackground}`,
  },
  replyItem: {
    padding: "8px",
    marginBottom: "8px",
    backgroundColor: tokens.colorNeutralBackground4,
    borderRadius: tokens.borderRadiusSmall,
    borderLeft: `3px solid ${tokens.colorPaletteRedBorder1}`,
  },
  replyContent: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    marginBottom: "4px",
  },
  replyMeta: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
  },
  emptyState: {
    textAlign: "center",
    padding: "24px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
  },
  errorState: {
    textAlign: "center",
    padding: "24px",
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase300,
  },
  jsonOutput: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    fontFamily: "monospace",
    color: tokens.colorNeutralForeground2,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    overflowX: "auto",
    maxHeight: "400px",
    overflowY: "auto",
  },
});

const Comments: React.FC = () => {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [comments, setComments] = useState<CommentInfo[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  // 选项状态 / Option states
  const [includeResolved, setIncludeResolved] = useState(true);
  const [includeReplies, setIncludeReplies] = useState(true);
  const [includeAssociatedText, setIncludeAssociatedText] = useState(true);
  const [detailedMetadata, setDetailedMetadata] = useState(false);
  const [maxTextLength, setMaxTextLength] = useState<string>("");

  /**
   * 获取批注内容
   * Get comments content
   */
  const handleGetComments = async () => {
    setLoading(true);
    setError(null);
    setComments(null);

    try {
      const options: GetCommentsOptions = {
        includeResolved,
        includeReplies,
        includeAssociatedText,
        detailedMetadata,
        maxTextLength: maxTextLength ? parseInt(maxTextLength, 10) : undefined,
      };

      console.log("获取批注内容，选项:", options);
      const result = await getComments(options);
      console.log("获取到的批注:", result);
      setComments(result);
    } catch (err) {
      console.error("获取批注内容失败:", err);
      setError(err instanceof Error ? err.message : "未知错误");
    } finally {
      setLoading(false);
    }
  };

  /**
   * 格式化日期
   * Format date
   */
  const formatDate = (date?: Date): string => {
    if (!date) return "";
    return new Date(date).toLocaleString("zh-CN");
  };

  /**
   * 渲染批注卡片
   * Render comment card
   */
  const renderCommentCard = (comment: CommentInfo, index: number) => {
    return (
      <Card key={comment.id} className={styles.resultCard}>
        <CardHeader
          header={
            <div className={styles.commentHeader}>
              <span className={styles.commentIcon}>💬</span>
              <span className={styles.commentTitle}>批注 {index + 1}</span>
              {comment.resolved !== undefined && (
                <Badge appearance={comment.resolved ? "filled" : "outline"} color="success">
                  {comment.resolved ? "已解决" : "未解决"}
                </Badge>
              )}
            </div>
          }
        />
        <div className={styles.cardContent}>
          {/* 元数据信息 / Metadata information */}
          {detailedMetadata && (
            <>
              <div className={styles.metadataGrid}>
                <span className={styles.metadataLabel}>ID:</span>
                <span className={styles.metadataValue}>{comment.id}</span>

                {comment.authorName && (
                  <>
                    <span className={styles.metadataLabel}>作者:</span>
                    <span className={styles.metadataValue}>{comment.authorName}</span>
                  </>
                )}

                {comment.authorEmail && (
                  <>
                    <span className={styles.metadataLabel}>邮箱:</span>
                    <span className={styles.metadataValue}>{comment.authorEmail}</span>
                  </>
                )}

                {comment.creationDate && (
                  <>
                    <span className={styles.metadataLabel}>创建时间:</span>
                    <span className={styles.metadataValue}>{formatDate(comment.creationDate)}</span>
                  </>
                )}
              </div>
              <Divider />
            </>
          )}

          {/* 批注内容 / Comment content */}
          <Label weight="semibold">批注内容:</Label>
          <div className={styles.commentContent}>{comment.content}</div>

          {/* 关联文本 / Associated text */}
          {includeAssociatedText && comment.associatedText && (
            <>
              <Label weight="semibold">关联文本:</Label>
              <div className={styles.associatedText}>{comment.associatedText}</div>
              
              {/* 位置信息和元数据 / Location info and metadata */}
              {comment.rangeLocation && (
                <div className={styles.metadataGrid} style={{ marginTop: "8px" }}>
                  {comment.rangeLocation.textHash && (
                    <>
                      <span className={styles.metadataLabel}>文本哈希:</span>
                      <span className={styles.metadataValue}>{comment.rangeLocation.textHash}</span>
                    </>
                  )}
                  {comment.rangeLocation.textLength !== undefined && (
                    <>
                      <span className={styles.metadataLabel}>文本长度:</span>
                      <span className={styles.metadataValue}>{comment.rangeLocation.textLength} 字符</span>
                    </>
                  )}
                  {comment.rangeLocation.paragraphIndex !== undefined && (
                    <>
                      <span className={styles.metadataLabel}>段落:</span>
                      <span className={styles.metadataValue}>第 {comment.rangeLocation.paragraphIndex + 1} 段</span>
                    </>
                  )}
                  {comment.rangeLocation.style && (
                    <>
                      <span className={styles.metadataLabel}>样式:</span>
                      <span className={styles.metadataValue}>{comment.rangeLocation.style}</span>
                    </>
                  )}
                  {comment.rangeLocation.isListItem && (
                    <>
                      <span className={styles.metadataLabel}>列表项:</span>
                      <span className={styles.metadataValue}>
                        是{comment.rangeLocation.listLevel !== undefined ? ` (级别 ${comment.rangeLocation.listLevel})` : ""}
                      </span>
                    </>
                  )}
                  {comment.rangeLocation.font && (
                    <>
                      <span className={styles.metadataLabel}>字体:</span>
                      <span className={styles.metadataValue}>
                        {comment.rangeLocation.font}
                        {comment.rangeLocation.fontSize ? ` (${comment.rangeLocation.fontSize}pt)` : ""}
                      </span>
                    </>
                  )}
                  {(comment.rangeLocation.isBold || comment.rangeLocation.isItalic || comment.rangeLocation.isUnderlined) && (
                    <>
                      <span className={styles.metadataLabel}>格式:</span>
                      <span className={styles.metadataValue}>
                        {[
                          comment.rangeLocation.isBold && "粗体",
                          comment.rangeLocation.isItalic && "斜体",
                          comment.rangeLocation.isUnderlined && "下划线",
                        ]
                          .filter(Boolean)
                          .join(", ")}
                      </span>
                    </>
                  )}
                  {comment.rangeLocation.highlightColor && comment.rangeLocation.highlightColor !== "None" && (
                    <>
                      <span className={styles.metadataLabel}>高亮:</span>
                      <span className={styles.metadataValue}>{comment.rangeLocation.highlightColor}</span>
                    </>
                  )}
                </div>
              )}
            </>
          )}

          {/* 批注回复 / Comment replies */}
          {includeReplies && comment.replies && comment.replies.length > 0 && (
            <>
              <Label weight="semibold">回复 ({comment.replies.length} 条):</Label>
              {comment.replies.map((reply) => (
                <div key={reply.id} className={styles.replyItem}>
                  <div className={styles.replyContent}>💬 {reply.content}</div>
                  {detailedMetadata && (
                    <div className={styles.replyMeta}>
                      {reply.authorName && `作者: ${reply.authorName}`}
                      {reply.authorEmail && ` (${reply.authorEmail})`}
                      {reply.creationDate && ` | ${formatDate(reply.creationDate)}`}
                    </div>
                  )}
                </div>
              ))}
            </>
          )}
        </div>
      </Card>
    );
  };

  return (
    <div className={styles.container}>
      {/* 选项配置 / Options configuration */}
      <div className={styles.optionsContainer}>
        <Label weight="semibold">获取选项</Label>

        <div className={styles.optionRow}>
          <Switch
            checked={includeResolved}
            onChange={(_, data) => setIncludeResolved(data.checked)}
            label="包含已解决的批注"
          />
        </div>

        <div className={styles.optionRow}>
          <Switch
            checked={includeReplies}
            onChange={(_, data) => setIncludeReplies(data.checked)}
            label="包含批注回复"
          />
        </div>

        <div className={styles.optionRow}>
          <Switch
            checked={includeAssociatedText}
            onChange={(_, data) => setIncludeAssociatedText(data.checked)}
            label="包含关联文本"
          />
        </div>

        <div className={styles.optionRow}>
          <Switch
            checked={detailedMetadata}
            onChange={(_, data) => setDetailedMetadata(data.checked)}
            label="详细元数据"
          />
        </div>

        <div className={styles.optionRow}>
          <Label>最大文本长度 (可选):</Label>
          <Input
            type="number"
            value={maxTextLength}
            onChange={(_, data) => setMaxTextLength(data.value)}
            placeholder="不限制"
          />
        </div>
      </div>

      {/* 获取按钮 / Get button */}
      <Button
        appearance="primary"
        className={styles.button}
        onClick={handleGetComments}
        disabled={loading}
      >
        {loading ? <Spinner size="tiny" /> : "获取批注内容"}
      </Button>

      {/* 结果展示 / Result display */}
      {error && <div className={styles.errorState}>错误: {error}</div>}

      {!loading && !error && comments !== null && (
        <div className={styles.resultContainer}>
          {comments.length === 0 ? (
            <div className={styles.emptyState}>未找到批注</div>
          ) : (
            <>
              <Label weight="semibold">找到 {comments.length} 条批注:</Label>
              {comments.map((comment, index) => renderCommentCard(comment, index))}

              {/* JSON 输出 / JSON output */}
              <Card className={styles.resultCard}>
                <CardHeader header={<Label weight="semibold">JSON 输出</Label>} />
                <div className={styles.cardContent}>
                  <div className={styles.jsonOutput}>{JSON.stringify(comments, null, 2)}</div>
                </div>
              </Card>
            </>
          )}
        </div>
      )}
    </div>
  );
};

export default Comments;
