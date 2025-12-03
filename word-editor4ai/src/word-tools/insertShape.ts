/**
 * 文件名: insertShape.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 插入形状工具核心逻辑
 */

/* global Word, console */

import type { InsertLocation } from "./types";

// 重新导出以保持向后兼容性 / Re-export for backward compatibility
export type { InsertLocation };

/**
 * 形状选项 / Shape Options
 */
export interface ShapeOptions {
  /** 形状宽度（磅），默认为 100 / Shape width in points, default 100 */
  width?: number;
  /** 形状高度（磅），默认为 100 / Shape height in points, default 100 */
  height?: number;
  /** 形状名称 / Shape name */
  name?: string;
  /** 是否锁定纵横比，默认为 false / Lock aspect ratio, default false */
  lockAspectRatio?: boolean;
  /** 是否可见，默认为 true / Visible, default true */
  visible?: boolean;
  /** 左边距（磅）/ Left position in points */
  left?: number;
  /** 上边距（磅）/ Top position in points */
  top?: number;
  /** 旋转角度（度）/ Rotation in degrees */
  rotation?: number;
  /** 填充颜色（十六进制格式，如 "#FF0000"）/ Fill color (hex format, e.g., "#FF0000") */
  fillColor?: string;
  /**
   * 线条颜色（十六进制格式）- 注意：当前 Word JavaScript API 不支持设置线条属性
   * Line color (hex format) - Note: Currently not supported by Word JavaScript API
   * @deprecated Word.Shape does not expose line formatting properties in the current API
   */
  lineColor?: string;
  /**
   * 线条宽度（磅）- 注意：当前 Word JavaScript API 不支持设置线条属性
   * Line weight in points - Note: Currently not supported by Word JavaScript API
   * @deprecated Word.Shape does not expose line formatting properties in the current API
   */
  lineWeight?: number;
  /**
   * 线条样式 - 注意：当前 Word JavaScript API 不支持设置线条属性
   * Line style - Note: Currently not supported by Word JavaScript API
   * @deprecated Word.Shape does not expose line formatting properties in the current API
   */
  lineStyle?: string;
  /** 形状文本内容（仅适用于支持文本的形状）/ Shape text content (only for shapes that support text) */
  text?: string;
}

/**
 * 插入形状结果 / Insert Shape Result
 */
export interface InsertShapeResult {
  /** 是否成功 / Success */
  success: boolean;
  /** 形状标识符 / Shape identifier */
  shapeId?: string;
  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;
}

/**
 * Word 支持的形状类型 / Word Supported Shape Types
 *
 * @remarks
 * Word JavaScript API 支持的形状类型基于 Word.GeometricShapeType 枚举
 * Word JavaScript API supported shape types based on Word.GeometricShapeType enum
 */
export type WordShapeType =
  | "Rectangle"
  | "RoundRectangle"
  | "Ellipse"
  | "Diamond"
  | "Triangle"
  | "RightTriangle"
  | "Parallelogram"
  | "Trapezoid"
  | "Hexagon"
  | "Octagon"
  | "Plus"
  | "Star"
  | "Arrow"
  | "HomePlate"
  | "Cube"
  | "Balloon"
  | "Seal"
  | "Arc"
  | "Line"
  | "Plaque"
  | "Can"
  | "Donut"
  | "TextBox"
  | "BlockArc"
  | "DoubleWave"
  | "Wave"
  | "VerticalScroll"
  | "HorizontalScroll"
  | "CircularArrow"
  | "UturnArrow"
  | "CurvedRightArrow"
  | "CurvedLeftArrow"
  | "CurvedUpArrow"
  | "CurvedDownArrow"
  | "CloudCallout"
  | "EllipseRibbon"
  | "EllipseRibbon2"
  | "FlowChartProcess"
  | "FlowChartDecision"
  | "FlowChartInputOutput"
  | "FlowChartPredefinedProcess"
  | "FlowChartInternalStorage"
  | "FlowChartDocument"
  | "FlowChartMultidocument"
  | "FlowChartTerminator"
  | "FlowChartPreparation"
  | "FlowChartManualInput"
  | "FlowChartManualOperation"
  | "FlowChartConnector"
  | "FlowChartPunchedCard"
  | "FlowChartPunchedTape"
  | "FlowChartSummingJunction"
  | "FlowChartOr"
  | "FlowChartCollate"
  | "FlowChartSort"
  | "FlowChartExtract"
  | "FlowChartMerge"
  | "FlowChartOfflineStorage"
  | "FlowChartOnlineStorage"
  | "FlowChartMagneticTape"
  | "FlowChartMagneticDisk"
  | "FlowChartMagneticDrum"
  | "FlowChartDisplay"
  | "FlowChartDelay"
  | "FlowChartAlternateProcess"
  | "FlowChartOffpageConnector"
  | "LeftRightRibbon"
  | "Chevron"
  | "PentagonRight"
  | "ChevronRight"
  | "LeftRightArrow"
  | "LeftRightArrowCallout"
  | "LeftRightUpArrow"
  | "LeftUpArrow"
  | "BentUpArrow"
  | "BentArrow"
  | "StripedRightArrow"
  | "NotchedRightArrow"
  | "Pentagon"
  | "QuadArrow"
  | "LeftArrow"
  | "RightArrow"
  | "UpArrow"
  | "DownArrow"
  | "LeftArrowCallout"
  | "RightArrowCallout"
  | "UpArrowCallout"
  | "DownArrowCallout"
  | "QuadArrowCallout"
  | "Bevel"
  | "LeftBracket"
  | "RightBracket"
  | "LeftBrace"
  | "RightBrace"
  | "LeftUpArrowCallout"
  | "BentUpArrowCallout"
  | "BentArrowCallout"
  | "Seal24"
  | "Seal16"
  | "Seal32"
  | "WedgeRectCallout"
  | "WedgeRRectCallout"
  | "WedgeEllipseCallout"
  | "FoldedCorner"
  | "UpDownArrow"
  | "UpDownArrowCallout"
  | "ExplosionOne"
  | "ExplosionTwo"
  | "LightningBolt"
  | "Heart"
  | "PictureFrame"
  | "LeftCircularArrow"
  | "LeftRightCircularArrow"
  | "SwooshArrow"
  | "Sun"
  | "Moon"
  | "BracketPair"
  | "BracePair"
  | "Seal4"
  | "ActionButtonBlank"
  | "ActionButtonHome"
  | "ActionButtonHelp"
  | "ActionButtonInformation"
  | "ActionButtonForwardNext"
  | "ActionButtonBackPrevious"
  | "ActionButtonEnd"
  | "ActionButtonBeginning"
  | "ActionButtonReturn"
  | "ActionButtonDocument"
  | "ActionButtonSound"
  | "ActionButtonMovie"
  | "Gear6"
  | "Gear9"
  | "Funnel"
  | "MathPlus"
  | "MathMinus"
  | "MathMultiply"
  | "MathDivide"
  | "MathEqual"
  | "MathNotEqual"
  | "CornerTabs"
  | "SquareTabs"
  | "PlaqueTabs"
  | "ChartX"
  | "ChartStar"
  | "ChartPlus";

/**
 * 在文档中插入形状
 * Insert shape in document
 *
 * @param shapeType - 形状类型 / Shape type
 * @param location - 插入位置 / Insert location
 * @param options - 形状选项 / Shape options
 * @returns Promise<InsertShapeResult> 插入结果 / Insert result
 *
 * @remarks
 * 注意：Word JavaScript API 对形状的支持
 * - 形状通过 insertGeometricShape 方法创建，返回 Word.Shape 对象
 * - 支持的形状类型基于 Word.GeometricShapeType 枚举
 * - 插入位置基于当前选择或文档范围
 * - 当前仅支持填充颜色，不支持线条样式（Word.Shape 不提供 line 属性）
 * - 某些高级属性（如精确定位、线条样式）可能需要 OOXML
 *
 * Note: Word JavaScript API support for shapes
 * - Shapes are created through insertGeometricShape method, returns Word.Shape object
 * - Supported shape types based on Word.GeometricShapeType enum
 * - Insert location is based on current selection or document range
 * - Currently only fill color is supported, line styles are not supported (Word.Shape does not expose line property)
 * - Some advanced properties (like precise positioning, line styles) may require OOXML
 *
 * @example
 * ```typescript
 * // 插入简单矩形
 * await insertShape("Rectangle", "End");
 *
 * // 插入带样式的圆形
 * await insertShape("Ellipse", "End", {
 *   width: 150,
 *   height: 150,
 *   name: "MyCircle",
 *   fillColor: "#FF0000"
 * });
 *
 * // 插入带文本的形状
 * await insertShape("RoundRectangle", "End", {
 *   width: 200,
 *   height: 100,
 *   text: "Hello World",
 *   fillColor: "#0078D4"
 * });
 * ```
 */
export async function insertShape(
  shapeType: string,
  location: InsertLocation = "End",
  options: ShapeOptions = {}
): Promise<InsertShapeResult> {
  const {
    width = 100,
    height = 100,
    name,
    lockAspectRatio = false,
    visible = true,
    left,
    top,
    rotation,
    fillColor,
    lineColor,
    lineWeight,
    lineStyle,
    text,
  } = options;

  // 验证参数 / Validate parameters
  if (!shapeType) {
    return {
      success: false,
      error: "必须提供形状类型 / Shape type is required",
    };
  }

  try {
    let shapeId: string | undefined;

    await Word.run(async (context) => {
      // 获取插入范围 / Get insert range
      let insertRange: Word.Range;
      const selection = context.document.getSelection();

      switch (location) {
        case "Start":
          insertRange = context.document.body.getRange("Start");
          break;
        case "End":
          insertRange = context.document.body.getRange("End");
          break;
        case "Before":
          insertRange = selection;
          break;
        case "After":
          insertRange = selection;
          break;
        case "Replace":
          insertRange = selection;
          break;
        default:
          insertRange = context.document.body.getRange("End");
      }

      // 在范围处插入形状
      // Insert shape at range
      // 注意：Word JavaScript API 通过 insertGeometricShape 创建形状，返回 Word.Shape 对象
      // Note: Word JavaScript API creates shapes through insertGeometricShape, returns Word.Shape object
      const insertShapeOptions: Word.InsertShapeOptions = {
        width,
        height,
      };

      // 添加位置参数（如果提供）/ Add position parameters (if provided)
      if (left !== undefined) {
        insertShapeOptions.left = left;
      }
      if (top !== undefined) {
        insertShapeOptions.top = top;
      }

      // 将形状类型转换为 Word.GeometricShapeType
      // Convert shape type to Word.GeometricShapeType
      const shape = insertRange.insertGeometricShape(
        shapeType as Word.GeometricShapeType,
        insertShapeOptions
      );

      // 设置形状属性 / Set shape properties
      if (name) {
        shape.name = name;
      }
      if (lockAspectRatio !== undefined) {
        shape.lockAspectRatio = lockAspectRatio;
      }
      if (visible !== undefined) {
        shape.visible = visible;
      }
      if (rotation !== undefined) {
        shape.rotation = rotation;
      }

      // 应用填充样式 / Apply fill styles
      // 注意：Word JavaScript API 当前不支持通过 Shape.line 属性设置线条格式
      // Note: Word JavaScript API currently does not support setting line formatting via Shape.line property
      try {
        if (fillColor) {
          const fill = shape.fill;
          fill.setSolidColor(fillColor);
        }

        // 线条样式暂不支持 / Line styles are not currently supported
        if (lineColor || lineWeight !== undefined || lineStyle) {
          console.warn(
            "线条样式设置暂不支持：Word.Shape 类当前不提供 line 或 lineFormat 属性 / " +
              "Line style settings are not supported: Word.Shape class does not currently expose line or lineFormat properties"
          );
        }
      } catch (error) {
        console.warn("应用形状样式时出错 / Error applying shape styles:", error);
      }

      // 添加文本内容（如果提供）/ Add text content (if provided)
      if (text) {
        try {
          // 获取形状的文本范围 / Get shape text range
          // Word.Shape 对象的 body 属性包含文本内容（仅适用于文本框和几何形状）
          // Word.Shape object's body property contains text content (only applies to text boxes and geometric shapes)
          const shapeBody = shape.body;
          const textRange = shapeBody.getRange("Whole");
          textRange.insertText(text, "Replace");
        } catch (error) {
          console.warn("添加形状文本时出错 / Error adding shape text:", error);
        }
      }

      // 加载形状 ID / Load shape ID
      shape.load("id");
      await context.sync();

      shapeId = `shape-${shape.id}`;
    });

    return {
      success: true,
      shapeId,
    };
  } catch (error) {
    console.error("插入形状失败 / Insert shape failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 批量插入形状 / Batch Insert Shapes
 *
 * @param shapes - 形状列表 / Shape list
 * @returns Promise<InsertShapeResult[]> 插入结果列表 / Insert result list
 */
export async function insertShapes(
  shapes: Array<{
    shapeType: string;
    location: InsertLocation;
    options?: ShapeOptions;
  }>
): Promise<InsertShapeResult[]> {
  const results: InsertShapeResult[] = [];

  for (const shapeData of shapes) {
    const result = await insertShape(shapeData.shapeType, shapeData.location, shapeData.options);
    results.push(result);
  }

  return results;
}
