import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, HTMLElement */

// Webpack HMR 类型定义 / Webpack HMR type definitions
interface WebpackHotModule {
  hot?: {
    accept(path: string, callback: () => void): void;
  };
}

declare const module: WebpackHotModule;

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App />
    </FluentProvider>
  );
});

// Webpack 热模块替换 (HMR) 配置 / Webpack Hot Module Replacement (HMR) configuration
if (module.hot) {
  module.hot.accept("./components/App", () => {
    // 动态导入更新后的组件 / Dynamically import updated component
    import("./components/App").then((module) => {
      const NextApp = module.default;
      root?.render(
        <FluentProvider theme={webLightTheme}>
          <NextApp />
        </FluentProvider>
      );
    });
  });
}
