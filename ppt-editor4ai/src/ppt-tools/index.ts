/**
 * 文件名: index.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: PPT 工具集导出文件
 */

// 文本插入工具
export { insertText, insertTextToSlide } from "./textInsertion";
export type { TextInsertionOptions } from "./textInsertion";

// 元素列表工具
export {
  getSlideElements,
  getCurrentSlideElements,
  getSlideElementsByPageNumber,
} from "./elementsList";
export type { SlideElement, GetElementsOptions } from "./elementsList";

// 页面布局信息工具
export {
  getSlideLayoutInfo,
  getCurrentSlideLayoutInfo,
  getSlideDimensions,
  getPresentationDimensions,
} from "./slideLayoutInfo";
export type {
  SlideLayoutInfo,
  EnhancedElement,
  SlideDimensions,
  RelativePosition,
  TextInfo,
  ImageInfo,
  FillInfo,
  LayoutInfo,
  BackgroundInfo,
  GetLayoutInfoOptions,
} from "./slideLayoutInfo";
