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

// 幻灯片布局模板工具
export {
  getAvailableSlideLayouts,
  createSlideWithLayout,
  getLayoutDescription,
} from "./slideLayouts";
export type { SlideLayoutTemplate, GetSlideLayoutsOptions } from "./slideLayouts";

// 图片插入工具
export {
  insertImage,
  insertImageToSlide,
  readImageAsBase64,
  fetchImageAsBase64,
} from "./imageInsertion";
export type { ImageInsertionOptions, ImageInsertionResult } from "./imageInsertion";

// 幻灯片截图工具
export {
  getSlideScreenshot,
  getCurrentSlideScreenshot,
  getSlideScreenshotByPageNumber,
  getAllSlidesScreenshots,
} from "./slideScreenshot";
export type { SlideScreenshotOptions, SlideScreenshotResult } from "./slideScreenshot";

// 视频插入工具
export {
  insertVideo,
  insertVideoToSlide,
  readVideoAsBase64,
  fetchVideoAsBase64,
} from "./videoInsertion";
export type { VideoInsertionOptions, VideoInsertionResult } from "./videoInsertion";

// 形状插入工具
export { insertShape, insertShapeToSlide, COMMON_SHAPES } from "./shapeInsertion";
export type { ShapeInsertionOptions, ShapeInsertionResult, ShapeType } from "./shapeInsertion";

// 表格插入工具
export { insertTable, insertTableToSlide, TABLE_TEMPLATES } from "./tableInsertion";
export type { TableInsertionOptions, TableInsertionResult } from "./tableInsertion";

// 元素删除工具
export {
  deleteElement,
  deleteElementById,
  deleteElementByName,
  deleteElementByIndex,
  deleteElementsByIds,
} from "./elementDeletion";
export type { DeleteElementOptions, DeleteElementResult } from "./elementDeletion";

// 幻灯片删除工具
export { deleteSlides, deleteCurrentSlide, deleteSlidesByNumbers } from "./slideDeletion";
export type { DeleteSlideOptions, DeleteSlideResult } from "./slideDeletion";

// 文本框更新工具
export { updateTextBox, updateTextBoxes, getTextBoxStyle } from "./textUpdate";
export type { TextUpdateOptions, TextUpdateResult } from "./textUpdate";
