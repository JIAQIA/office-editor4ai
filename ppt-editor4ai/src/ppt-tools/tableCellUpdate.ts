/**
 * 文件名: tableCellUpdate.ts
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 最后修改日期: 2025/12/1
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 表格单元格编辑工具核心逻辑 - 通过坐标修改单元格内容
 * Description: Table cell update tool core logic - modify cell content by coordinates
 */

/* global PowerPoint, console */

/**
 * 单元格更新选项
 * Cell update options
 */
export interface CellUpdateOptions {
  /** 行索引（从 0 开始）Row index (0-based) */
  rowIndex: number;
  /** 列索引（从 0 开始）Column index (0-based) */
  columnIndex: number;
  /** 新的文本内容 New text content */
  text: string;
}

/**
 * 批量单元格更新选项
 * Batch cell update options
 */
export interface BatchCellUpdateOptions {
  /** 单元格更新列表 List of cell updates */
  cells: CellUpdateOptions[];
}

/**
 * 表格定位选项
 * Table location options
 */
export interface TableLocationOptions {
  /** 表格形状 ID（可选）Table shape ID (optional) */
  shapeId?: string;
  /** 表格索引（可选，默认第一个表格）Table index (optional, defaults to first table) */
  tableIndex?: number;
}

/**
 * 单元格更新结果
 * Cell update result
 */
export interface CellUpdateResult {
  /** 是否成功 Success status */
  success: boolean;
  /** 更新的单元格数量 Number of cells updated */
  cellsUpdated: number;
  /** 表格行数 Table row count */
  rowCount: number;
  /** 表格列数 Table column count */
  columnCount: number;
  /** 错误信息（如果有）Error message (if any) */
  error?: string;
}

/**
 * 更新单个表格单元格内容
 * Update single table cell content
 *
 * @param cellOptions 单元格更新选项 Cell update options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<CellUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 更新第 2 行第 3 列的单元格
 * // Update cell at row 2, column 3
 * const result = await updateTableCell(
 *   { rowIndex: 1, columnIndex: 2, text: "新内容" },
 *   { shapeId: "shape123" }
 * );
 * ```
 */
export async function updateTableCell(
  cellOptions: CellUpdateOptions,
  tableLocation?: TableLocationOptions
): Promise<CellUpdateResult> {
  const { rowIndex, columnIndex, text } = cellOptions;

  // 验证参数
  // Validate parameters
  if (rowIndex < 0 || columnIndex < 0) {
    throw new Error("行索引和列索引必须大于等于 0 / Row and column indices must be >= 0");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("shapes");
      await context.sync();

      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      // 查找表格
      // Find table
      let tableShape: PowerPoint.Shape | undefined;

      if (tableLocation?.shapeId) {
        // 通过 ID 查找
        // Find by ID
        tableShape = shapes.items.find((shape) => shape.id === tableLocation.shapeId);
        if (!tableShape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
      } else {
        // 查找所有表格
        // Find all tables
        const tables = shapes.items.filter((shape) => shape.type === PowerPoint.ShapeType.table);
        if (tables.length === 0) {
          throw new Error("当前幻灯片没有表格 / No tables found in current slide");
        }

        const tableIndex = tableLocation?.tableIndex ?? 0;
        if (tableIndex >= tables.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围，当前幻灯片只有 ${tables.length} 个表格 / Table index ${tableIndex} out of range, only ${tables.length} tables found`
          );
        }

        tableShape = tables[tableIndex];
      }

      // 获取表格对象
      // Get table object
      const table = tableShape.getTable();
      table.load("rowCount, columnCount");
      await context.sync();

      // 验证坐标范围
      // Validate coordinate range
      if (rowIndex >= table.rowCount) {
        throw new Error(
          `行索引 ${rowIndex} 超出范围，表格只有 ${table.rowCount} 行 / Row index ${rowIndex} out of range, table has ${table.rowCount} rows`
        );
      }
      if (columnIndex >= table.columnCount) {
        throw new Error(
          `列索引 ${columnIndex} 超出范围，表格只有 ${table.columnCount} 列 / Column index ${columnIndex} out of range, table has ${table.columnCount} columns`
        );
      }

      // 更新单元格
      // Update cell
      const cell = table.getCellOrNullObject(rowIndex, columnIndex);
      cell.text = text;
      await context.sync();

      return {
        success: true,
        cellsUpdated: 1,
        rowCount: table.rowCount,
        columnCount: table.columnCount,
      };
    });
  } catch (error) {
    console.error("更新表格单元格失败 / Failed to update table cell:", error);
    return {
      success: false,
      cellsUpdated: 0,
      rowCount: 0,
      columnCount: 0,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 批量更新多个表格单元格内容
 * Batch update multiple table cells
 *
 * @param batchOptions 批量更新选项 Batch update options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<CellUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 批量更新多个单元格
 * // Batch update multiple cells
 * const result = await updateTableCellsBatch({
 *   cells: [
 *     { rowIndex: 0, columnIndex: 0, text: "标题1" },
 *     { rowIndex: 0, columnIndex: 1, text: "标题2" },
 *     { rowIndex: 1, columnIndex: 0, text: "数据1" },
 *     { rowIndex: 1, columnIndex: 1, text: "数据2" }
 *   ]
 * });
 * ```
 */
export async function updateTableCellsBatch(
  batchOptions: BatchCellUpdateOptions,
  tableLocation?: TableLocationOptions
): Promise<CellUpdateResult> {
  const { cells } = batchOptions;

  if (!cells || cells.length === 0) {
    throw new Error("单元格列表不能为空 / Cell list cannot be empty");
  }

  // 验证所有坐标
  // Validate all coordinates
  for (const cell of cells) {
    if (cell.rowIndex < 0 || cell.columnIndex < 0) {
      throw new Error(
        `单元格 (${cell.rowIndex}, ${cell.columnIndex}) 坐标无效 / Cell coordinates (${cell.rowIndex}, ${cell.columnIndex}) are invalid`
      );
    }
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("shapes");
      await context.sync();

      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      // 查找表格
      // Find table
      let tableShape: PowerPoint.Shape | undefined;

      if (tableLocation?.shapeId) {
        tableShape = shapes.items.find((shape) => shape.id === tableLocation.shapeId);
        if (!tableShape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
      } else {
        const tables = shapes.items.filter((shape) => shape.type === PowerPoint.ShapeType.table);
        if (tables.length === 0) {
          throw new Error("当前幻灯片没有表格 / No tables found in current slide");
        }

        const tableIndex = tableLocation?.tableIndex ?? 0;
        if (tableIndex >= tables.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`
          );
        }

        tableShape = tables[tableIndex];
      }

      // 获取表格对象
      // Get table object
      const table = tableShape.getTable();
      table.load("rowCount, columnCount");
      await context.sync();

      // 批量更新单元格
      // Batch update cells
      let updatedCount = 0;
      for (const cellOption of cells) {
        const { rowIndex, columnIndex, text } = cellOption;

        // 验证坐标
        // Validate coordinates
        if (rowIndex >= table.rowCount || columnIndex >= table.columnCount) {
          console.warn(
            `跳过无效坐标 (${rowIndex}, ${columnIndex}) / Skipping invalid coordinates (${rowIndex}, ${columnIndex})`
          );
          continue;
        }

        // 更新单元格
        // Update cell
        const cell = table.getCellOrNullObject(rowIndex, columnIndex);
        cell.text = text;
        updatedCount++;
      }

      await context.sync();

      return {
        success: true,
        cellsUpdated: updatedCount,
        rowCount: table.rowCount,
        columnCount: table.columnCount,
      };
    });
  } catch (error) {
    console.error("批量更新表格单元格失败 / Failed to batch update table cells:", error);
    return {
      success: false,
      cellsUpdated: 0,
      rowCount: 0,
      columnCount: 0,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 获取表格单元格内容
 * Get table cell content
 *
 * @param rowIndex 行索引 Row index
 * @param columnIndex 列索引 Column index
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<string> 单元格文本内容 Cell text content
 */
export async function getTableCellContent(
  rowIndex: number,
  columnIndex: number,
  tableLocation?: TableLocationOptions
): Promise<string> {
  if (rowIndex < 0 || columnIndex < 0) {
    throw new Error("行索引和列索引必须大于等于 0 / Row and column indices must be >= 0");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("shapes");
      await context.sync();

      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      // 查找表格
      // Find table
      let tableShape: PowerPoint.Shape | undefined;

      if (tableLocation?.shapeId) {
        tableShape = shapes.items.find((shape) => shape.id === tableLocation.shapeId);
        if (!tableShape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
      } else {
        const tables = shapes.items.filter((shape) => shape.type === PowerPoint.ShapeType.table);
        if (tables.length === 0) {
          throw new Error("当前幻灯片没有表格 / No tables found in current slide");
        }

        const tableIndex = tableLocation?.tableIndex ?? 0;
        tableShape = tables[tableIndex];
      }

      // 获取表格对象
      // Get table object
      const table = tableShape.getTable();
      table.load("rowCount, columnCount");
      await context.sync();

      // 验证坐标
      // Validate coordinates
      if (rowIndex >= table.rowCount || columnIndex >= table.columnCount) {
        throw new Error(
          `坐标 (${rowIndex}, ${columnIndex}) 超出范围 / Coordinates (${rowIndex}, ${columnIndex}) out of range`
        );
      }

      // 获取单元格内容
      // Get cell content
      const cell = table.getCellOrNullObject(rowIndex, columnIndex);
      cell.load("text");
      await context.sync();

      return cell.text;
    });
  } catch (error) {
    console.error("获取表格单元格内容失败 / Failed to get table cell content:", error);
    throw error;
  }
}
