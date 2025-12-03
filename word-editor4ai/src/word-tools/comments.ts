/**
 * 文件名: comments.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 获取批注内容的工具核心逻辑，与 Word API 交互
 */

/* global Word, console */

import type { CommentInfo, CommentReplyInfo, GetCommentsOptions } from "./types";

/**
 * 生成文本的简单哈希值 / Generate simple hash for text
 * 用于识别重复的引用文本 / Used to identify duplicate referenced text
 */
function simpleHash(text: string): string {
  let hash = 0;
  for (let i = 0; i < text.length; i++) {
    const char = text.charCodeAt(i);
    hash = (hash << 5) - hash + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return Math.abs(hash).toString(16);
}

/**
 * 获取批注内容
 * Get comments content
 *
 * @param options - 获取选项 / Get options
 * @returns Promise<CommentInfo[]> 批注信息列表 / Comment information list
 *
 * @remarks
 * 此函数按以下优先级获取批注：
 * 1. 如果用户有选择，优先返回选择范围内的批注
 * 2. 如果没有选择，返回整个文档的批注
 * 3. 可以选择是否包含已解决的批注
 *
 * This function gets comments in the following priority:
 * 1. If user has a selection, return comments in the selection
 * 2. If no selection, return comments in the entire document
 * 3. Can choose whether to include resolved comments
 *
 * @example
 * ```typescript
 * // 获取批注内容
 * const comments = await getComments({
 *   includeResolved: true,
 *   includeReplies: true,
 *   includeAssociatedText: true,
 *   detailedMetadata: true
 * });
 *
 * console.log(`找到 ${comments.length} 条批注`);
 * comments.forEach(comment => {
 *   console.log(`批注: ${comment.content}, 作者: ${comment.authorName}`);
 * });
 * ```
 */
export async function getComments(options: GetCommentsOptions = {}): Promise<CommentInfo[]> {
  const {
    includeResolved = true,
    includeReplies = true,
    includeAssociatedText = true,
    detailedMetadata = false,
    maxTextLength,
  } = options;

  try {
    return await Word.run(async (context) => {
      // 获取当前选中的范围 / Get current selection range
      const selection = context.document.getSelection();
      // eslint-disable-next-line office-addins/no-navigational-load
      selection.load(["text", "isEmpty"]);
      await context.sync();

      let comments: Word.CommentCollection;
      let rangeType: "selection" | "document";

      // 判断是否有选择（通过 isEmpty 判断）/ Check if there is a selection (by isEmpty)
      if (!selection.isEmpty) {
        // 获取选择范围内的批注 / Get comments in selection
        try {
          comments = selection.getComments();
          rangeType = "selection";
          console.log("获取选择范围内的批注 / Getting comments in selection");
        } catch (error) {
          // 如果获取选择范围批注失败，回退到获取文档所有批注 / If getting selection comments fails, fallback to all comments
          console.warn(
            "获取选择范围批注失败，将获取文档所有批注 / Failed to get selection comments, will get all comments:",
            error
          );
          comments = context.document.body.getComments();
          rangeType = "document";
        }
      } else {
        // 没有选择，获取文档所有批注 / No selection, get all comments in document
        comments = context.document.body.getComments();
        rangeType = "document";
        console.log("获取文档所有批注 / Getting all comments in document");
      }

      // 加载批注集合 / Load comments collection
      comments.load("items");
      await context.sync();

      if (comments.items.length === 0) {
        console.log(
          `在${rangeType === "selection" ? "选择范围" : "文档"}内未找到批注 / No comments found in ${rangeType === "selection" ? "selection" : "document"}`
        );
        return [];
      }

      console.log(
        `在${rangeType === "selection" ? "选择范围" : "文档"}内找到 ${comments.items.length} 条批注 / Found ${comments.items.length} comments in ${rangeType === "selection" ? "selection" : "document"}`
      );

      // 批量加载批注的基本属性 / Batch load basic properties of comments
      for (const comment of comments.items) {
        comment.load("id,content,resolved");
        if (detailedMetadata) {
          comment.load("authorName,authorEmail,createdDate");
        }
        if (includeAssociatedText) {
          comment.load("contentRange");
        }
        if (includeReplies) {
          comment.load("replies");
        }
      }
      await context.sync();

      // 加载批注关联的文本 / Load associated text of comments
      if (includeAssociatedText) {
        for (const comment of comments.items) {
          try {
            const contentRange = comment.contentRange;
            if (contentRange) {
              contentRange.load("text");
            }
          } catch (error) {
            console.warn(
              `加载批注 ${comment.id} 的关联文本失败 / Failed to load associated text of comment ${comment.id}:`,
              error
            );
          }
        }
        await context.sync();
      }

      // 加载批注回复 / Load comment replies
      const commentsWithReplies: Word.Comment[] = [];
      if (includeReplies) {
        for (const comment of comments.items) {
          try {
            const replies = comment.replies;
            replies.load("items");
            commentsWithReplies.push(comment);
          } catch (error) {
            console.warn(
              `加载批注 ${comment.id} 的回复失败 / Failed to load replies of comment ${comment.id}:`,
              error
            );
          }
        }

        if (commentsWithReplies.length > 0) {
          try {
            await context.sync();
          } catch (error) {
            console.warn(
              `同步批注回复失败，将跳过所有回复详情 / Failed to sync comment replies, will skip all reply details:`,
              error
            );
            commentsWithReplies.length = 0; // 清空数组 / Clear array
          }
        }

        // 加载回复的详细属性 / Load detailed reply properties
        if (commentsWithReplies.length > 0) {
          for (const comment of commentsWithReplies) {
            try {
              const replies = comment.replies;
              for (const reply of replies.items) {
                reply.load("id,content");
                if (detailedMetadata) {
                  reply.load("authorName,authorEmail,createdDate");
                }
              }
            } catch (error) {
              console.warn(
                `加载批注 ${comment.id} 回复详细属性失败 / Failed to load detailed reply properties for comment ${comment.id}:`,
                error
              );
            }
          }

          try {
            await context.sync();
          } catch (error) {
            console.warn(`同步回复详细属性失败 / Failed to sync reply details:`, error);
          }
        }
      }

      // 构建批注信息列表 / Build comment information list
      const commentInfoList: CommentInfo[] = [];
      const rangeMap = new Map<string, Word.Range>();

      // 第一步：构建基本批注信息并批量加载范围 / Step 1: Build basic comment info and batch load ranges
      for (const comment of comments.items) {
        try {
          // 根据选项过滤已解决的批注 / Filter resolved comments based on options
          if (!includeResolved && comment.resolved) {
            continue;
          }

          let content = comment.content;
          if (maxTextLength && content.length > maxTextLength) {
            content = content.substring(0, maxTextLength) + "...";
          }

          const commentInfo: CommentInfo = {
            id: comment.id,
            content: content,
            resolved: comment.resolved,
          };

          // 添加详细元数据 / Add detailed metadata
          if (detailedMetadata) {
            commentInfo.authorName = comment.authorName;
            commentInfo.authorEmail = comment.authorEmail;
            commentInfo.creationDate = comment.creationDate;
          }

          // 批量加载关联文本范围 / Batch load associated text ranges
          if (includeAssociatedText) {
            try {
              const range = comment.getRange();
              // eslint-disable-next-line office-addins/no-navigational-load
              range.load([
                "text",
                "style",
                "parentBody",
                "font",
                "hyperlink",
                "isEmpty",
              ]);
              // 加载字体属性 / Load font properties
              range.font.load([
                "name",
                "size",
                "bold",
                "italic",
                "underline",
                "highlightColor",
              ]);
              rangeMap.set(comment.id, range);
            } catch (error) {
              console.warn(
                `加载批注 ${comment.id} 的范围失败 / Failed to load range for comment ${comment.id}:`,
                error
              );
            }
          }

          commentInfoList.push(commentInfo);
        } catch (error) {
          console.warn(`处理批注失败 / Failed to process comment:`, error);
        }
      }

      // 第二步：同步所有批注范围的加载 / Step 2: Sync all comment range loads
      if (includeAssociatedText && rangeMap.size > 0) {
        try {
          await context.sync();
        } catch (error) {
          console.warn(`同步批注范围失败 / Failed to sync comment ranges:`, error);
        }
      }

      // 第三步：处理关联文本和位置信息 / Step 3: Process associated text and location info
      for (const commentInfo of commentInfoList) {
        try {
          if (includeAssociatedText) {
            const range = rangeMap.get(commentInfo.id);
            if (range && range.text !== undefined) {
              let associatedText = range.text;
              if (maxTextLength && associatedText.length > maxTextLength) {
                associatedText = associatedText.substring(0, maxTextLength) + "...";
              }
              commentInfo.associatedText = associatedText;

              // 添加位置信息和元数据 / Add location info and metadata
              commentInfo.rangeLocation = {
                style: range.style,
                textHash: simpleHash(range.text),
                textLength: range.text.length,
              };

              // 添加字体和格式化信息 / Add font and formatting info
              try {
                if (range.font) {
                  commentInfo.rangeLocation.font = range.font.name;
                  commentInfo.rangeLocation.fontSize = range.font.size;
                  commentInfo.rangeLocation.isBold = range.font.bold;
                  commentInfo.rangeLocation.isItalic = range.font.italic;
                  commentInfo.rangeLocation.isUnderlined =
                    range.font.underline !== "None" && range.font.underline !== undefined;
                  commentInfo.rangeLocation.highlightColor = range.font.highlightColor;
                }
              } catch (fontError) {
                console.warn(
                  `获取批注 ${commentInfo.id} 的字体信息失败 / Failed to get font info for comment ${commentInfo.id}:`,
                  fontError
                );
              }

              // 注意：列表项信息（isListItem, listLevel）需要通过 Paragraph 对象获取
              // 这需要额外的 sync 操作，会影响性能，因此暂时不获取
              // Note: List item info (isListItem, listLevel) requires accessing Paragraph object
              // This needs extra sync operations which affects performance, so skipped for now
              // TODO: 如果需要列表信息，考虑在批量加载时一次性获取所有段落的列表属性
              // TODO: If list info is needed, consider batch loading all paragraph list properties at once

              // 注意：获取段落索引需要多次 sync，这是一个性能权衡
              // 如果需要段落索引，建议在外部批量处理所有批注
              // Note: Getting paragraph index requires multiple syncs, this is a performance tradeoff
              // If paragraph index is needed, recommend batch processing all comments externally
              // 暂时跳过段落索引以避免性能问题 / Skip paragraph index for now to avoid performance issues
              // TODO: 实现批量获取段落索引的优化方案 / TODO: Implement optimized batch paragraph index retrieval
            }
          }

          // 获取批注回复 / Get comment replies
          if (includeReplies) {
            const comment = comments.items.find((c) => c.id === commentInfo.id);
            if (comment && commentsWithReplies.includes(comment)) {
              try {
                const replies = comment.replies;
                const replyInfoList: CommentReplyInfo[] = [];

                for (const reply of replies.items) {
                  let replyContent = reply.content;
                  if (maxTextLength && replyContent.length > maxTextLength) {
                    replyContent = replyContent.substring(0, maxTextLength) + "...";
                  }

                  const replyInfo: CommentReplyInfo = {
                    id: reply.id,
                    content: replyContent,
                  };

                  if (detailedMetadata) {
                    replyInfo.authorName = reply.authorName;
                    replyInfo.authorEmail = reply.authorEmail;
                    replyInfo.creationDate = reply.creationDate;
                  }

                  replyInfoList.push(replyInfo);
                }

                commentInfo.replies = replyInfoList;
              } catch (error) {
                console.warn(
                  `获取批注 ${commentInfo.id} 的回复详情失败 / Failed to get reply details for comment ${commentInfo.id}:`,
                  error
                );
              }
            }
          }
        } catch (error) {
          console.warn(`处理批注详情失败 / Failed to process comment details:`, error);
        }
      }

      return commentInfoList;
    });
  } catch (error) {
    console.error("获取批注内容失败 / Failed to get comments:", error);
    throw error;
  }
}

/**
 * 根据文本哈希识别重复的批注引用 / Identify duplicate comment references by text hash
 * 
 * @param comments - 批注列表 / Comment list
 * @returns 重复引用的分组 / Grouped duplicate references
 * 
 * @example
 * ```typescript
 * const comments = await getComments({ includeAssociatedText: true });
 * const duplicates = findDuplicateReferences(comments);
 * 
 * duplicates.forEach(group => {
 *   console.log(`文本 "${group.text}" 被引用了 ${group.comments.length} 次`);
 *   console.log(`批注 ID: ${group.comments.map(c => c.id).join(', ')}`);
 * });
 * ```
 */
export function findDuplicateReferences(
  comments: CommentInfo[]
): Array<{ textHash: string; text: string; count: number; comments: CommentInfo[] }> {
  const hashMap = new Map<string, CommentInfo[]>();

  // 按文本哈希分组 / Group by text hash
  for (const comment of comments) {
    const hash = comment.rangeLocation?.textHash;
    if (hash && comment.associatedText) {
      if (!hashMap.has(hash)) {
        hashMap.set(hash, []);
      }
      hashMap.get(hash)!.push(comment);
    }
  }

  // 只返回有重复的组 / Only return groups with duplicates
  const duplicates: Array<{ textHash: string; text: string; count: number; comments: CommentInfo[] }> = [];
  
  hashMap.forEach((commentList, hash) => {
    if (commentList.length > 1) {
      duplicates.push({
        textHash: hash,
        text: commentList[0].associatedText || "",
        count: commentList.length,
        comments: commentList,
      });
    }
  });

  return duplicates;
}
