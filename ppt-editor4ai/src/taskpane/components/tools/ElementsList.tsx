/**
 * 文件名: ElementsList.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components
 * 描述: 元素列表工具，用于获取并显示当前幻灯片中的所有元素
 */

/* global console */

import * as React from "react";
import { useState } from "react";
import { Button, makeStyles, tokens, Spinner, Input, Label } from "@fluentui/react-components";
import { getSlideElements, type SlideElement } from "../../../ppt-tools";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    width: "100%",
    gap: "16px",
  },
  inputContainer: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginBottom: "8px",
  },
  inputField: {
    width: "100%",
  },
  buttonContainer: {
    width: "100%",
    display: "flex",
    justifyContent: "center",
    marginBottom: "8px",
  },
  emptyState: {
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
    padding: "32px 16px",
  },
  elementsList: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  elementCard: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    padding: "12px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  elementHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
  },
  elementType: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorBrandForeground1,
  },
  elementId: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    fontFamily: "monospace",
  },
  elementDetails: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "8px",
    fontSize: tokens.fontSizeBase200,
  },
  detailItem: {
    display: "flex",
    flexDirection: "column",
  },
  detailLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    marginBottom: "2px",
  },
  detailValue: {
    color: tokens.colorNeutralForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
  elementName: {
    marginTop: "8px",
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    fontStyle: "italic",
  },
  elementText: {
    marginTop: "8px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground1,
    wordBreak: "break-word",
    lineHeight: "1.4",
  },
  textLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    marginBottom: "4px",
    fontWeight: tokens.fontWeightSemibold,
  },
  errorMessage: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase300,
    padding: "16px",
    textAlign: "center",
  },
});

const ElementsList: React.FC = () => {
  const styles = useStyles();
  const [elements, setElements] = useState<SlideElement[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [slideNumber, setSlideNumber] = useState<string>("");

  const fetchElements = async () => {
    setLoading(true);
    setError(null);
    
    try {
      // 解析页码输入
      const pageNum = slideNumber.trim() === "" ? undefined : parseInt(slideNumber, 10);
      
      // 验证页码
      if (pageNum !== undefined && (isNaN(pageNum) || pageNum < 1)) {
        setError("页码必须是大于0的整数");
        setLoading(false);
        return;
      }
      
      const elementsList = await getSlideElements({ 
        slideNumber: pageNum,
        includeText: true 
      });
      
      setElements(elementsList);
      
      // 如果返回空数组，显示提示
      if (elementsList.length === 0 && pageNum !== undefined) {
        setError(`页码 ${pageNum} 不存在或没有元素`);
      }
    } catch (err) {
      console.error("获取元素列表失败:", err);
      setError(err instanceof Error ? err.message : "获取元素列表失败");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.inputContainer}>
        <Label htmlFor="slideNumber">
          页码（可选，不填则使用当前页）
        </Label>
        <Input
          id="slideNumber"
          type="number"
          min="1"
          placeholder="请输入页码，从1开始"
          value={slideNumber}
          onChange={(e) => setSlideNumber(e.target.value)}
          className={styles.inputField}
          disabled={loading}
        />
      </div>
      
      <div className={styles.buttonContainer}>
        <Button 
          appearance="primary" 
          size="large" 
          onClick={fetchElements}
          disabled={loading}
        >
          {loading ? <Spinner size="tiny" /> : "获取元素列表"}
        </Button>
      </div>

      {error && (
        <div className={styles.errorMessage}>
          ❌ {error}
        </div>
      )}

      {!loading && !error && elements.length === 0 && (
        <div className={styles.emptyState}>
          输入页码（可选）并点击按钮获取元素列表
        </div>
      )}

      {elements.length > 0 && (
        <div className={styles.elementsList}>
          {elements.map((element, index) => (
            <div key={element.id} className={styles.elementCard}>
              <div className={styles.elementHeader}>
                <span className={styles.elementType}>
                  {element.type}
                </span>
                <span className={styles.elementId}>
                  #{index + 1}
                </span>
              </div>
              
              <div className={styles.elementDetails}>
                <div className={styles.detailItem}>
                  <span className={styles.detailLabel}>X 坐标</span>
                  <span className={styles.detailValue}>{element.left}</span>
                </div>
                <div className={styles.detailItem}>
                  <span className={styles.detailLabel}>Y 坐标</span>
                  <span className={styles.detailValue}>{element.top}</span>
                </div>
                <div className={styles.detailItem}>
                  <span className={styles.detailLabel}>宽度</span>
                  <span className={styles.detailValue}>{element.width}</span>
                </div>
                <div className={styles.detailItem}>
                  <span className={styles.detailLabel}>高度</span>
                  <span className={styles.detailValue}>{element.height}</span>
                </div>
              </div>
              
              {element.name && (
                <div className={styles.elementName}>
                  名称: {element.name}
                </div>
              )}
              
              {element.text && (
                <div className={styles.elementText}>
                  <div className={styles.textLabel}>文本内容:</div>
                  {element.text.length > 10 
                    ? `${element.text.substring(0, 10)}...` 
                    : element.text
                  }
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default ElementsList;
