/**
 * 文件名: tableInsertion.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 表格插入工具核心逻辑，与 Office API 交互
 */

/* global PowerPoint, console */

import { getSlideDimensions } from "./slideLayoutInfo";

/**
 * 表格插入选项
 */
export interface TableInsertionOptions {
  /** 行数 */
  rowCount: number;
  /** 列数 */
  columnCount: number;
  /** X 坐标（可选，单位：磅） */
  left?: number;
  /** Y 坐标（可选，单位：磅） */
  top?: number;
  /** 宽度（可选，单位：磅，默认 400） */
  width?: number;
  /** 高度（可选，单位：磅，默认根据行数计算） */
  height?: number;
  /** 表格数据（可选，二维数组） */
  values?: string[][];
  /** 是否显示表头（可选，默认 true） */
  showHeader?: boolean;
  /** 表头背景色（可选，默认蓝色） */
  headerColor?: string;
  /** 表格边框颜色（可选，默认灰色） */
  borderColor?: string;
}

/**
 * 表格插入结果
 */
export interface TableInsertionResult {
  /** 插入的表格形状 ID */
  shapeId: string;
  /** 行数 */
  rowCount: number;
  /** 列数 */
  columnCount: number;
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
 * 插入表格到幻灯片
 *
 * @param options 表格插入选项
 * @returns Promise<TableInsertionResult> 插入结果
 *
 * @example
 * ```typescript
 * // 插入一个 3x4 的空表格
 * const result = await insertTableToSlide({
 *   rowCount: 3,
 *   columnCount: 4
 * });
 *
 * // 插入带数据的表格
 * const tableWithData = await insertTableToSlide({
 *   rowCount: 3,
 *   columnCount: 3,
 *   values: [
 *     ["姓名", "年龄", "城市"],
 *     ["张三", "25", "北京"],
 *     ["李四", "30", "上海"]
 *   ],
 *   showHeader: true,
 *   headerColor: "#4472C4"
 * });
 * ```
 */
export async function insertTableToSlide(
  options: TableInsertionOptions
): Promise<TableInsertionResult> {
  const {
    rowCount,
    columnCount,
    left,
    top,
    width = 400,
    height,
    values,
    showHeader = true,
    headerColor = "#4472C4",
    borderColor = "#D0D0D0",
  } = options;

  // 验证参数
  if (rowCount <= 0 || columnCount <= 0) {
    throw new Error("行数和列数必须大于 0");
  }

  if (rowCount > 100 || columnCount > 50) {
    throw new Error("表格过大：行数不能超过 100，列数不能超过 50");
  }

  // 验证数据维度
  if (values) {
    if (values.length !== rowCount) {
      throw new Error(`数据行数 (${values.length}) 与指定行数 (${rowCount}) 不匹配`);
    }
    for (let i = 0; i < values.length; i++) {
      if (values[i].length !== columnCount) {
        throw new Error(
          `第 ${i + 1} 行数据列数 (${values[i].length}) 与指定列数 (${columnCount}) 不匹配`
        );
      }
    }
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);

      // 计算默认高度（如果未指定）
      const defaultRowHeight = 30;
      const actualHeight = height ?? rowCount * defaultRowHeight;

      // 计算位置（如果未指定，则居中）
      let actualLeft = left;
      let actualTop = top;

      if (actualLeft === undefined || actualTop === undefined) {
        const dimensions = await getSlideDimensions();
        const slideWidth = dimensions.width;
        const slideHeight = dimensions.height;

        actualLeft = actualLeft ?? (slideWidth - width) / 2;
        actualTop = actualTop ?? (slideHeight - actualHeight) / 2;
      }

      // 创建表格
      const tableShape = slide.shapes.addTable(rowCount, columnCount, {
        left: actualLeft,
        top: actualTop,
        width,
        height: actualHeight,
        values: values,
      });

      // 获取表格对象（使用 getTable() 方法）
      const table = tableShape.getTable();

      // 设置表格样式
      // 使用 getCellOrNullObject 方法访问单元格
      for (let i = 0; i < rowCount; i++) {
        for (let j = 0; j < columnCount; j++) {
          const cell = table.getCellOrNullObject(i, j);

          // 设置边框（通过 borders 属性访问）
          cell.borders.bottom.color = borderColor;
          cell.borders.top.color = borderColor;
          cell.borders.left.color = borderColor;
          cell.borders.right.color = borderColor;

          // 如果是表头行且启用表头样式
          if (i === 0 && showHeader) {
            cell.fill.setSolidColor(headerColor);
            // 设置表头文字样式
            cell.font.color = "#FFFFFF";
            cell.font.bold = true;
          }
        }
      }

      await context.sync();

      // 加载属性以返回结果
      tableShape.load("id,width,height,left,top");
      await context.sync();

      return {
        shapeId: tableShape.id,
        rowCount,
        columnCount,
        width: tableShape.width,
        height: tableShape.height,
        left: tableShape.left,
        top: tableShape.top,
      };
    });
  } catch (error) {
    console.error("插入表格失败:", error);
    throw error;
  }
}

/**
 * 简化版本：插入表格（兼容旧接口）
 *
 * @param rowCount 行数
 * @param columnCount 列数
 * @param left X 坐标（可选）
 * @param top Y 坐标（可选）
 * @param width 宽度（可选）
 * @param height 高度（可选）
 * @returns Promise<TableInsertionResult> 插入结果
 */
export async function insertTable(
  rowCount: number,
  columnCount: number,
  left?: number,
  top?: number,
  width?: number,
  height?: number
): Promise<TableInsertionResult> {
  return insertTableToSlide({ rowCount, columnCount, left, top, width, height });
}

/**
 * 常用表格模板
 */
export const TABLE_TEMPLATES = [
  {
    id: "simple-2x3",
    name: "简单表格 (2行3列)",
    rowCount: 2,
    columnCount: 3,
    description: "适合简单数据展示",
  },
  {
    id: "simple-3x3",
    name: "方形表格 (3行3列)",
    rowCount: 3,
    columnCount: 3,
    description: "适合对比数据",
  },
  {
    id: "list-5x2",
    name: "列表表格 (5行2列)",
    rowCount: 5,
    columnCount: 2,
    description: "适合列表展示",
  },
  {
    id: "data-4x5",
    name: "数据表格 (4行5列)",
    rowCount: 4,
    columnCount: 5,
    description: "适合数据分析",
  },
  {
    id: "schedule-7x5",
    name: "日程表格 (7行5列)",
    rowCount: 7,
    columnCount: 5,
    description: "适合周计划或日程",
  },
];
