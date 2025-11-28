/**
 * 文件名: toolsConfig.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 工具配置文件，定义所有可用工具的元数据和组件
 */

import * as React from "react";
import TextInsertion from "./TextInsertion";
import ElementsList from "./ElementsList";

export interface ToolConfig {
  id: string;
  title: string;
  subtitle: string;
  icon?: string;
  component: React.ReactNode;
}

/**
 * 所有可用工具的配置
 */
export const toolsConfig: Record<string, ToolConfig> = {
  "text-insertion": {
    id: "text-insertion",
    title: "文本插入工具",
    subtitle: "在幻灯片中插入文本框，支持自定义位置",
    component: <TextInsertion />,
  },
  "elements-list": {
    id: "elements-list",
    title: "元素列表",
    subtitle: "获取当前幻灯片中所有元素的信息",
    component: <ElementsList />,
  },
};

/**
 * 获取工具配置
 */
export const getToolConfig = (toolId: string): ToolConfig | undefined => {
  return toolsConfig[toolId];
};

/**
 * 获取所有工具ID列表
 */
export const getAllToolIds = (): string[] => {
  return Object.keys(toolsConfig);
};
