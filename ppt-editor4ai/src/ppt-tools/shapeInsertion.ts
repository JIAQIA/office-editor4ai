/**
 * 文件名: shapeInsertion.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 形状插入工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint, console */

import { getSlideDimensions } from "./slideLayoutInfo";

/**
 * PowerPoint 支持的形状类型
 * 参考: https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapetype
 */
export type ShapeType =
  | "rectangle"
  | "roundRectangle"
  | "ellipse"
  | "diamond"
  | "triangle"
  | "rightTriangle"
  | "parallelogram"
  | "trapezoid"
  | "hexagon"
  | "octagon"
  | "plus"
  | "star5"
  | "star6"
  | "star7"
  | "star8"
  | "star10"
  | "star12"
  | "star16"
  | "star24"
  | "star32"
  | "round1Rectangle"
  | "round2SameRectangle"
  | "round2DiagonalRectangle"
  | "snipRoundRectangle"
  | "snip1Rectangle"
  | "snip2SameRectangle"
  | "snip2DiagonalRectangle"
  | "plaque"
  | "ellipseRibbon"
  | "ellipseRibbon2"
  | "leftRightRibbon"
  | "verticalScroll"
  | "horizontalScroll"
  | "wave"
  | "doubleWave"
  | "leftRightArrow"
  | "upDownArrow"
  | "leftUpArrow"
  | "bentUpArrow"
  | "bentArrow"
  | "stripedRightArrow"
  | "notchedRightArrow"
  | "homePlate"
  | "chevron"
  | "rightArrowCallout"
  | "downArrowCallout"
  | "leftArrowCallout"
  | "upArrowCallout"
  | "leftRightArrowCallout"
  | "quadArrowCallout"
  | "circularArrow"
  | "mathPlus"
  | "mathMinus"
  | "mathMultiply"
  | "mathDivide"
  | "mathEqual"
  | "mathNotEqual"
  | "cornerTabs"
  | "squareTabs"
  | "plaqueTabs"
  | "chartX"
  | "chartStar"
  | "chartPlus";

/**
 * 形状插入选项
 */
export interface ShapeInsertionOptions {
  /** 形状类型 */
  shapeType: ShapeType;
  /** X 坐标（可选，单位：磅） */
  left?: number;
  /** Y 坐标（可选，单位：磅） */
  top?: number;
  /** 宽度（可选，单位：磅，默认 100） */
  width?: number;
  /** 高度（可选，单位：磅，默认 100） */
  height?: number;
  /** 填充颜色（可选，默认蓝色） */
  fillColor?: string;
  /** 边框颜色（可选，默认深蓝色） */
  lineColor?: string;
  /** 边框粗细（可选，单位：磅，默认 2） */
  lineWeight?: number;
  /** 形状内文本（可选） */
  text?: string;
}

/**
 * 形状插入结果
 */
export interface ShapeInsertionResult {
  /** 插入的形状 ID */
  shapeId: string;
  /** 形状类型 */
  shapeType: string;
  /** 实际宽度 */
  width: number;
  /** 实际高度 */
  height: number;
  /** 实际 X 坐标 */
  left: number;
  /** 实际 Y 坐标 */
  top: number;
}

/**
 * 将自定义形状类型映射到 PowerPoint API 的 GeometricShapeType
 */
function mapShapeType(shapeType: ShapeType): PowerPoint.GeometricShapeType {
  // PowerPoint API 使用枚举类型
  const typeMap: Record<string, PowerPoint.GeometricShapeType> = {
    rectangle: PowerPoint.GeometricShapeType.rectangle,
    roundRectangle: PowerPoint.GeometricShapeType.roundRectangle,
    ellipse: PowerPoint.GeometricShapeType.ellipse,
    diamond: PowerPoint.GeometricShapeType.diamond,
    triangle: PowerPoint.GeometricShapeType.triangle,
    rightTriangle: PowerPoint.GeometricShapeType.rightTriangle,
    parallelogram: PowerPoint.GeometricShapeType.parallelogram,
    trapezoid: PowerPoint.GeometricShapeType.trapezoid,
    hexagon: PowerPoint.GeometricShapeType.hexagon,
    octagon: PowerPoint.GeometricShapeType.octagon,
    plus: PowerPoint.GeometricShapeType.plus,
    star5: PowerPoint.GeometricShapeType.star5,
    star6: PowerPoint.GeometricShapeType.star6,
    star7: PowerPoint.GeometricShapeType.star7,
    star8: PowerPoint.GeometricShapeType.star8,
    star10: PowerPoint.GeometricShapeType.star10,
    star12: PowerPoint.GeometricShapeType.star12,
    star16: PowerPoint.GeometricShapeType.star16,
    star24: PowerPoint.GeometricShapeType.star24,
    star32: PowerPoint.GeometricShapeType.star32,
    round1Rectangle: PowerPoint.GeometricShapeType.round1Rectangle,
    round2SameRectangle: PowerPoint.GeometricShapeType.round2SameRectangle,
    round2DiagonalRectangle: PowerPoint.GeometricShapeType.round2DiagonalRectangle,
    snipRoundRectangle: PowerPoint.GeometricShapeType.snipRoundRectangle,
    snip1Rectangle: PowerPoint.GeometricShapeType.snip1Rectangle,
    snip2SameRectangle: PowerPoint.GeometricShapeType.snip2SameRectangle,
    snip2DiagonalRectangle: PowerPoint.GeometricShapeType.snip2DiagonalRectangle,
    plaque: PowerPoint.GeometricShapeType.plaque,
    ellipseRibbon: PowerPoint.GeometricShapeType.ellipseRibbon,
    ellipseRibbon2: PowerPoint.GeometricShapeType.ellipseRibbon2,
    leftRightRibbon: PowerPoint.GeometricShapeType.leftRightRibbon,
    verticalScroll: PowerPoint.GeometricShapeType.verticalScroll,
    horizontalScroll: PowerPoint.GeometricShapeType.horizontalScroll,
    wave: PowerPoint.GeometricShapeType.wave,
    doubleWave: PowerPoint.GeometricShapeType.doubleWave,
    leftRightArrow: PowerPoint.GeometricShapeType.leftRightArrow,
    upDownArrow: PowerPoint.GeometricShapeType.upDownArrow,
    leftUpArrow: PowerPoint.GeometricShapeType.leftUpArrow,
    bentUpArrow: PowerPoint.GeometricShapeType.bentUpArrow,
    bentArrow: PowerPoint.GeometricShapeType.bentArrow,
    stripedRightArrow: PowerPoint.GeometricShapeType.stripedRightArrow,
    notchedRightArrow: PowerPoint.GeometricShapeType.notchedRightArrow,
    homePlate: PowerPoint.GeometricShapeType.homePlate,
    chevron: PowerPoint.GeometricShapeType.chevron,
    rightArrowCallout: PowerPoint.GeometricShapeType.rightArrowCallout,
    downArrowCallout: PowerPoint.GeometricShapeType.downArrowCallout,
    leftArrowCallout: PowerPoint.GeometricShapeType.leftArrowCallout,
    upArrowCallout: PowerPoint.GeometricShapeType.upArrowCallout,
    leftRightArrowCallout: PowerPoint.GeometricShapeType.leftRightArrowCallout,
    quadArrowCallout: PowerPoint.GeometricShapeType.quadArrowCallout,
    circularArrow: PowerPoint.GeometricShapeType.circularArrow,
    mathPlus: PowerPoint.GeometricShapeType.mathPlus,
    mathMinus: PowerPoint.GeometricShapeType.mathMinus,
    mathMultiply: PowerPoint.GeometricShapeType.mathMultiply,
    mathDivide: PowerPoint.GeometricShapeType.mathDivide,
    mathEqual: PowerPoint.GeometricShapeType.mathEqual,
    mathNotEqual: PowerPoint.GeometricShapeType.mathNotEqual,
    cornerTabs: PowerPoint.GeometricShapeType.cornerTabs,
    squareTabs: PowerPoint.GeometricShapeType.squareTabs,
    plaqueTabs: PowerPoint.GeometricShapeType.plaqueTabs,
    chartX: PowerPoint.GeometricShapeType.chartX,
    chartStar: PowerPoint.GeometricShapeType.chartStar,
    chartPlus: PowerPoint.GeometricShapeType.chartPlus,
  };

  return typeMap[shapeType] || PowerPoint.GeometricShapeType.rectangle;
}

/**
 * 插入形状到幻灯片
 *
 * @param options 形状插入选项
 * @returns Promise<ShapeInsertionResult> 插入结果
 *
 * @example
 * ```typescript
 * // 插入一个矩形
 * const result = await insertShapeToSlide({
 *   shapeType: "rectangle",
 *   left: 100,
 *   top: 100,
 *   width: 200,
 *   height: 100,
 *   fillColor: "#4472C4",
 *   lineColor: "#2E5090",
 *   lineWeight: 2,
 *   text: "示例文本"
 * });
 *
 * // 插入一个星形（使用默认位置和尺寸）
 * const star = await insertShapeToSlide({
 *   shapeType: "star5",
 *   fillColor: "#FFD700"
 * });
 * ```
 */
export async function insertShapeToSlide(
  options: ShapeInsertionOptions
): Promise<ShapeInsertionResult> {
  const {
    shapeType,
    left,
    top,
    width = 100,
    height = 100,
    fillColor = "#4472C4",
    lineColor = "#2E5090",
    lineWeight = 2,
    text,
  } = options;

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);

      // 映射形状类型
      const geometricShapeType = mapShapeType(shapeType);

      // 计算位置（如果未指定，则居中）
      let actualLeft = left;
      let actualTop = top;

      if (actualLeft === undefined || actualTop === undefined) {
        // 获取幻灯片尺寸以计算居中位置
        const dimensions = await getSlideDimensions();
        const slideWidth = dimensions.width;
        const slideHeight = dimensions.height;

        actualLeft = actualLeft ?? (slideWidth - width) / 2;
        actualTop = actualTop ?? (slideHeight - height) / 2;
      }

      // 添加几何形状
      const shape = slide.shapes.addGeometricShape(geometricShapeType, {
        left: actualLeft,
        top: actualTop,
        width,
        height,
      });

      // 设置填充颜色
      shape.fill.setSolidColor(fillColor);

      // 设置边框样式
      shape.lineFormat.color = lineColor;
      shape.lineFormat.weight = lineWeight;

      // 如果提供了文本，添加文本
      if (text) {
        const textFrame = shape.textFrame;
        textFrame.textRange.text = text;
      }

      // 加载属性以返回结果
      shape.load("id,type,width,height,left,top");
      await context.sync();

      return {
        shapeId: shape.id,
        shapeType: shape.type,
        width: shape.width,
        height: shape.height,
        left: shape.left,
        top: shape.top,
      };
    });
  } catch (error) {
    console.error("插入形状失败:", error);
    throw error;
  }
}

/**
 * 简化版本：插入形状（兼容旧接口）
 *
 * @param shapeType 形状类型
 * @param left X 坐标（可选）
 * @param top Y 坐标（可选）
 * @param width 宽度（可选）
 * @param height 高度（可选）
 * @returns Promise<ShapeInsertionResult> 插入结果
 */
export async function insertShape(
  shapeType: ShapeType,
  left?: number,
  top?: number,
  width?: number,
  height?: number
): Promise<ShapeInsertionResult> {
  return insertShapeToSlide({ shapeType, left, top, width, height });
}

/**
 * 获取常用形状列表（用于 UI 展示）
 */
export const COMMON_SHAPES: Array<{ type: ShapeType; label: string; category: string }> = [
  // 基础形状
  { type: "rectangle", label: "矩形", category: "基础形状" },
  { type: "roundRectangle", label: "圆角矩形", category: "基础形状" },
  { type: "ellipse", label: "椭圆", category: "基础形状" },
  { type: "diamond", label: "菱形", category: "基础形状" },
  { type: "triangle", label: "三角形", category: "基础形状" },
  { type: "rightTriangle", label: "直角三角形", category: "基础形状" },
  { type: "parallelogram", label: "平行四边形", category: "基础形状" },
  { type: "trapezoid", label: "梯形", category: "基础形状" },
  { type: "hexagon", label: "六边形", category: "基础形状" },
  { type: "octagon", label: "八边形", category: "基础形状" },
  { type: "plus", label: "加号", category: "基础形状" },

  // 星形
  { type: "star5", label: "五角星", category: "星形" },
  { type: "star6", label: "六角星", category: "星形" },
  { type: "star8", label: "八角星", category: "星形" },
  { type: "star10", label: "十角星", category: "星形" },
  { type: "star12", label: "十二角星", category: "星形" },

  // 箭头
  { type: "leftRightArrow", label: "左右箭头", category: "箭头" },
  { type: "upDownArrow", label: "上下箭头", category: "箭头" },
  { type: "leftUpArrow", label: "左上箭头", category: "箭头" },
  { type: "bentUpArrow", label: "向上弯曲箭头", category: "箭头" },
  { type: "bentArrow", label: "弯曲箭头", category: "箭头" },
  { type: "stripedRightArrow", label: "条纹右箭头", category: "箭头" },
  { type: "notchedRightArrow", label: "缺口右箭头", category: "箭头" },
  { type: "homePlate", label: "五边形箭头", category: "箭头" },
  { type: "chevron", label: "V形箭头", category: "箭头" },
  { type: "circularArrow", label: "环形箭头", category: "箭头" },

  // 标注
  { type: "rightArrowCallout", label: "右箭头标注", category: "标注" },
  { type: "downArrowCallout", label: "下箭头标注", category: "标注" },
  { type: "leftArrowCallout", label: "左箭头标注", category: "标注" },
  { type: "upArrowCallout", label: "上箭头标注", category: "标注" },
  { type: "leftRightArrowCallout", label: "左右箭头标注", category: "标注" },
  { type: "quadArrowCallout", label: "四向箭头标注", category: "标注" },

  // 装饰形状
  { type: "plaque", label: "铭牌", category: "装饰形状" },
  { type: "ellipseRibbon", label: "椭圆缎带", category: "装饰形状" },
  { type: "ellipseRibbon2", label: "椭圆缎带2", category: "装饰形状" },
  { type: "leftRightRibbon", label: "左右缎带", category: "装饰形状" },
  { type: "verticalScroll", label: "垂直卷轴", category: "装饰形状" },
  { type: "horizontalScroll", label: "水平卷轴", category: "装饰形状" },
  { type: "wave", label: "波浪", category: "装饰形状" },
  { type: "doubleWave", label: "双波浪", category: "装饰形状" },

  // 数学符号
  { type: "mathPlus", label: "数学加号", category: "数学符号" },
  { type: "mathMinus", label: "数学减号", category: "数学符号" },
  { type: "mathMultiply", label: "数学乘号", category: "数学符号" },
  { type: "mathDivide", label: "数学除号", category: "数学符号" },
  { type: "mathEqual", label: "数学等号", category: "数学符号" },
  { type: "mathNotEqual", label: "数学不等号", category: "数学符号" },
];
