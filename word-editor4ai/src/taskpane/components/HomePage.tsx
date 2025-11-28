/**
 * 文件名: HomePage.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @fluentui/react-components, @fluentui/react-icons
 * 描述: 首页组件，展示欢迎信息和插件介绍
 */

import * as React from "react";
import { Image, tokens, makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "20px 16px",
    minHeight: "100vh",
    minWidth: "280px", // 确保内容区有最小宽度
  },
  header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    paddingBottom: "16px",
    paddingTop: "20px",
    backgroundColor: tokens.colorNeutralBackground3,
    width: "100%",
  },
  message: {
    fontSize: tokens.fontSizeHero800,
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground1,
    marginTop: "10px",
  },
  mainContent: {
    width: "100%",
    maxWidth: "480px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    marginTop: "24px",
    padding: "0 8px",
  },
  description: {
    fontSize: tokens.fontSizeBase400,
    color: tokens.colorNeutralForeground2,
    fontWeight: tokens.fontWeightRegular,
    textAlign: "center",
    marginBottom: "20px",
    lineHeight: "1.5",
  },
  featureList: {
    width: "100%",
    listStyle: "none",
    padding: 0,
    margin: 0,
  },
  featureItem: {
    display: "flex",
    alignItems: "center",
    padding: "12px 16px",
    marginBottom: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow4,
  },
  icon: {
    marginRight: "12px",
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
  },
  featureText: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    lineHeight: "1.4",
  },
});

const HomePage: React.FC = () => {
  const styles = useStyles();

  return (
    <div className={styles.container}>
      <section className={styles.header}>
        <Image width="60" height="60" src="assets/logo-filled.png" alt="Word Editor for AI" />
        <h1 className={styles.message}>欢迎</h1>
      </section>

      <div className={styles.mainContent}>
        <p className={styles.description}>
          欢迎使用 Word Editor for AI！<br />
          这是一个专为 AI Agent 设计的 Word 编辑工具包，提供强大的文档编辑能力。<br />
        </p>

        <ul className={styles.featureList}>
          <li className={styles.featureItem}>
            <i className={styles.icon}>
              <Ribbon24Regular />
            </i>
            <span className={styles.featureText}>
              与 Office 深度集成，实现更多功能
            </span>
          </li>
          <li className={styles.featureItem}>
            <i className={styles.icon}>
              <LockOpen24Regular />
            </i>
            <span className={styles.featureText}>
              解锁强大的编辑功能
            </span>
          </li>
          <li className={styles.featureItem}>
            <i className={styles.icon}>
              <DesignIdeas24Regular />
            </i>
            <span className={styles.featureText}>
              像专业人士一样创建和可视化
            </span>
          </li>
        </ul>
      </div>
    </div>
  );
};

export default HomePage;
