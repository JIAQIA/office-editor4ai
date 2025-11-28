import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const title = "Contoso Task Pane Add-in";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

// Webpack 热模块替换 (HMR) 配置 / Webpack Hot Module Replacement (HMR) configuration
// eslint-disable-next-line @typescript-eslint/no-explicit-any
if ((module as any).hot) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (module as any).hot.accept("./components/App", () => {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
