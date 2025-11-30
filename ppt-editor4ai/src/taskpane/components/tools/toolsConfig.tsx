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
import SlideLayoutInfo from "./SlideLayoutInfo";
import SlideLayouts from "./SlideLayouts";
import ImageInsertion from "./ImageInsertion";
import VideoInsertion from "./VideoInsertion";
import SlideScreenshot from "./SlideScreenshot";
import ShapeInsertion from "./ShapeInsertion";
import TableInsertion from "./TableInsertion";
import { ElementDeletion } from "./ElementDeletion";

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
  "slide-layout-info": {
    id: "slide-layout-info",
    title: "页面布局信息",
    subtitle: "获取页面完整布局、尺寸和元素详细信息，支持导出 JSON 用于 AutoLayout 计算",
    component: <SlideLayoutInfo />,
  },
  "slide-layouts": {
    id: "slide-layouts",
    title: "布局模板列表",
    subtitle: "获取可用的幻灯片布局模板，支持使用指定模板创建新幻灯片",
    component: <SlideLayouts />,
  },
  "image-insertion": {
    id: "image-insertion",
    title: "图片插入工具",
    subtitle: "在幻灯片中插入图片，支持本地上传或 URL",
    component: <ImageInsertion />,
  },
  "slide-screenshot": {
    id: "slide-screenshot",
    title: "幻灯片截图工具",
    subtitle: "获取指定幻灯片的截图，支持导出 PNG 格式用于 AutoLayout 分析",
    component: <SlideScreenshot />,
  },
  "video-insertion": {
    id: "video-insertion",
    title: "视频插入工具",
    subtitle: "在幻灯片中插入视频，支持本地上传或 URL（实验性功能）",
    component: <VideoInsertion />,
  },
  "shape-insertion": {
    id: "shape-insertion",
    title: "形状插入工具",
    subtitle: "在幻灯片中插入各种几何形状，支持自定义样式和文本",
    component: <ShapeInsertion />,
  },
  "table-insertion": {
    id: "table-insertion",
    title: "表格插入工具",
    subtitle: "在幻灯片中插入表格，支持自定义行列数、样式和数据填充",
    component: <TableInsertion />,
  },
  "element-deletion": {
    id: "element-deletion",
    title: "元素删除工具",
    subtitle: "删除幻灯片中的元素，支持通过ID、名称或索引选择元素（调试工具）",
    component: <ElementDeletion />,
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
