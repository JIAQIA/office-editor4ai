/**
 * 文件名: tableRowColumnUpdate.ts
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 最后修改日期: 2025/12/1
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 表格行/列批量编辑工具核心逻辑 - 批量修改整行或整列
 * Description: Table row/column batch update tool - batch modify entire row or column
 */

/* global PowerPoint, console */

import { TableLocationOptions, CellUpdateResult } from "./tableCellUpdate";

/**
 * 行更新选项
 * Row update options
 */
export interface RowUpdateOptions {
  /** 行索引（从 0 开始）Row index (0-based) */
  rowIndex: number;
  /** 行数据（数组长度应等于列数）Row data (array length should equal column count) */
  values: string[];
  /** 是否跳过空值（可选，默认 false）Skip empty values (optional, default false) */
  skipEmpty?: boolean;
}

/**
 * 列更新选项
 * Column update options
 */
export interface ColumnUpdateOptions {
  /** 列索引（从 0 开始）Column index (0-based) */
  columnIndex: number;
  /** 列数据（数组长度应等于行数）Column data (array length should equal row count) */
  values: string[];
  /** 是否跳过空值（可选，默认 false）Skip empty values (optional, default false) */
  skipEmpty?: boolean;
}

/**
 * 批量行更新选项
 * Batch row update options
 */
export interface BatchRowUpdateOptions {
  /** 行更新列表 List of row updates */
  rows: RowUpdateOptions[];
}

/**
 * 批量列更新选项
 * Batch column update options
 */
export interface BatchColumnUpdateOptions {
  /** 列更新列表 List of column updates */
  columns: ColumnUpdateOptions[];
}

/**
 * 更新表格整行内容
 * Update entire table row
 *
 * @param rowOptions 行更新选项 Row update options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<CellUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 更新第 2 行的所有单元格
 * // Update all cells in row 2
 * const result = await updateTableRow(
 *   { rowIndex: 1, values: ["数据1", "数据2", "数据3"] },
 *   { shapeId: "shape123" }
 * );
 * ```
 */
export async function updateTableRow(
  rowOptions: RowUpdateOptions,
  tableLocation?: TableLocationOptions
): Promise<CellUpdateResult> {
  const { rowIndex, values, skipEmpty = false } = rowOptions;

  // 验证参数
  // Validate parameters
  if (rowIndex < 0) {
    throw new Error("行索引必须大于等于 0 / Row index must be >= 0");
  }

  if (!values || values.length === 0) {
    throw new Error("行数据不能为空 / Row data cannot be empty");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // eslint-disable-next-line office-addins/no-navigational-load
      slide.shapes.load("items");
      await context.sync();

      // 查找表格
      // Find table
      let tableShape: PowerPoint.Shape | undefined;

      if (tableLocation?.shapeId) {
        tableShape = slide.shapes.items.find((shape) => shape.id === tableLocation.shapeId);
        if (!tableShape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
      } else {
        const tables = slide.shapes.items.filter(
          (shape) => shape.type === PowerPoint.ShapeType.table
        );
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

      // 验证行索引
      // Validate row index
      if (rowIndex >= table.rowCount) {
        throw new Error(
          `行索引 ${rowIndex} 超出范围，表格只有 ${table.rowCount} 行 / Row index ${rowIndex} out of range, table has ${table.rowCount} rows`
        );
      }

      // 验证数据长度
      // Validate data length
      if (values.length > table.columnCount) {
        console.warn(
          `行数据长度 (${values.length}) 超过表格列数 (${table.columnCount})，将截断多余数据 / Row data length (${values.length}) exceeds table column count (${table.columnCount}), extra data will be truncated`
        );
      }

      // 更新行中的所有单元格
      // Update all cells in the row
      let updatedCount = 0;
      const maxColumns = Math.min(values.length, table.columnCount);

      for (let colIndex = 0; colIndex < maxColumns; colIndex++) {
        const value = values[colIndex];

        // 如果启用跳过空值且当前值为空，则跳过
        // Skip if skipEmpty is enabled and current value is empty
        if (skipEmpty && (!value || value.trim() === "")) {
          continue;
        }

        const cell = table.getCellOrNullObject(rowIndex, colIndex);
        cell.text = value;
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
    console.error("更新表格行失败 / Failed to update table row:", error);
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
 * 更新表格整列内容
 * Update entire table column
 *
 * @param columnOptions 列更新选项 Column update options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<CellUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 更新第 3 列的所有单元格
 * // Update all cells in column 3
 * const result = await updateTableColumn(
 *   { columnIndex: 2, values: ["标题", "数据1", "数据2"] },
 *   { tableIndex: 0 }
 * );
 * ```
 */
export async function updateTableColumn(
  columnOptions: ColumnUpdateOptions,
  tableLocation?: TableLocationOptions
): Promise<CellUpdateResult> {
  const { columnIndex, values, skipEmpty = false } = columnOptions;

  // 验证参数
  // Validate parameters
  if (columnIndex < 0) {
    throw new Error("列索引必须大于等于 0 / Column index must be >= 0");
  }

  if (!values || values.length === 0) {
    throw new Error("列数据不能为空 / Column data cannot be empty");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // eslint-disable-next-line office-addins/no-navigational-load
      slide.shapes.load("items");
      await context.sync();

      // 查找表格
      // Find table
      let tableShape: PowerPoint.Shape | undefined;

      if (tableLocation?.shapeId) {
        tableShape = slide.shapes.items.find((shape) => shape.id === tableLocation.shapeId);
        if (!tableShape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
      } else {
        const tables = slide.shapes.items.filter(
          (shape) => shape.type === PowerPoint.ShapeType.table
        );
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

      // 验证列索引
      // Validate column index
      if (columnIndex >= table.columnCount) {
        throw new Error(
          `列索引 ${columnIndex} 超出范围，表格只有 ${table.columnCount} 列 / Column index ${columnIndex} out of range, table has ${table.columnCount} columns`
        );
      }

      // 验证数据长度
      // Validate data length
      if (values.length > table.rowCount) {
        console.warn(
          `列数据长度 (${values.length}) 超过表格行数 (${table.rowCount})，将截断多余数据 / Column data length (${values.length}) exceeds table row count (${table.rowCount}), extra data will be truncated`
        );
      }

      // 更新列中的所有单元格
      // Update all cells in the column
      let updatedCount = 0;
      const maxRows = Math.min(values.length, table.rowCount);

      for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
        const value = values[rowIndex];

        // 如果启用跳过空值且当前值为空，则跳过
        // Skip if skipEmpty is enabled and current value is empty
        if (skipEmpty && (!value || value.trim() === "")) {
          continue;
        }

        const cell = table.getCellOrNullObject(rowIndex, columnIndex);
        cell.text = value;
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
    console.error("更新表格列失败 / Failed to update table column:", error);
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
 * 批量更新多行
 * Batch update multiple rows
 *
 * @param batchOptions 批量行更新选项 Batch row update options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<CellUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 批量更新多行
 * // Batch update multiple rows
 * const result = await updateTableRowsBatch({
 *   rows: [
 *     { rowIndex: 0, values: ["标题1", "标题2", "标题3"] },
 *     { rowIndex: 1, values: ["数据1", "数据2", "数据3"] },
 *     { rowIndex: 2, values: ["数据4", "数据5", "数据6"] }
 *   ]
 * });
 * ```
 */
export async function updateTableRowsBatch(
  batchOptions: BatchRowUpdateOptions,
  tableLocation?: TableLocationOptions
): Promise<CellUpdateResult> {
  const { rows } = batchOptions;

  if (!rows || rows.length === 0) {
    throw new Error("行列表不能为空 / Row list cannot be empty");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // eslint-disable-next-line office-addins/no-navigational-load
      slide.shapes.load("items");
      await context.sync();

      // 查找表格
      // Find table
      let tableShape: PowerPoint.Shape | undefined;

      if (tableLocation?.shapeId) {
        tableShape = slide.shapes.items.find((shape) => shape.id === tableLocation.shapeId);
        if (!tableShape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
      } else {
        const tables = slide.shapes.items.filter(
          (shape) => shape.type === PowerPoint.ShapeType.table
        );
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

      // 批量更新行
      // Batch update rows
      let totalUpdated = 0;

      for (const rowOption of rows) {
        const { rowIndex, values, skipEmpty = false } = rowOption;

        // 验证行索引
        // Validate row index
        if (rowIndex < 0 || rowIndex >= table.rowCount) {
          console.warn(`跳过无效行索引 ${rowIndex} / Skipping invalid row index ${rowIndex}`);
          continue;
        }

        // 更新行
        // Update row
        const maxColumns = Math.min(values.length, table.columnCount);
        for (let colIndex = 0; colIndex < maxColumns; colIndex++) {
          const value = values[colIndex];

          if (skipEmpty && (!value || value.trim() === "")) {
            continue;
          }

          const cell = table.getCellOrNullObject(rowIndex, colIndex);
          cell.text = value;
          totalUpdated++;
        }
      }

      await context.sync();

      return {
        success: true,
        cellsUpdated: totalUpdated,
        rowCount: table.rowCount,
        columnCount: table.columnCount,
      };
    });
  } catch (error) {
    console.error("批量更新表格行失败 / Failed to batch update table rows:", error);
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
 * 批量更新多列
 * Batch update multiple columns
 *
 * @param batchOptions 批量列更新选项 Batch column update options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<CellUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 批量更新多列
 * // Batch update multiple columns
 * const result = await updateTableColumnsBatch({
 *   columns: [
 *     { columnIndex: 0, values: ["姓名", "张三", "李四"] },
 *     { columnIndex: 1, values: ["年龄", "25", "30"] }
 *   ]
 * });
 * ```
 */
export async function updateTableColumnsBatch(
  batchOptions: BatchColumnUpdateOptions,
  tableLocation?: TableLocationOptions
): Promise<CellUpdateResult> {
  const { columns } = batchOptions;

  if (!columns || columns.length === 0) {
    throw new Error("列列表不能为空 / Column list cannot be empty");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // eslint-disable-next-line office-addins/no-navigational-load
      slide.shapes.load("items");
      await context.sync();

      // 查找表格
      // Find table
      let tableShape: PowerPoint.Shape | undefined;

      if (tableLocation?.shapeId) {
        tableShape = slide.shapes.items.find((shape) => shape.id === tableLocation.shapeId);
        if (!tableShape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
      } else {
        const tables = slide.shapes.items.filter(
          (shape) => shape.type === PowerPoint.ShapeType.table
        );
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

      // 批量更新列
      // Batch update columns
      let totalUpdated = 0;

      for (const columnOption of columns) {
        const { columnIndex, values, skipEmpty = false } = columnOption;

        // 验证列索引
        // Validate column index
        if (columnIndex < 0 || columnIndex >= table.columnCount) {
          console.warn(
            `跳过无效列索引 ${columnIndex} / Skipping invalid column index ${columnIndex}`
          );
          continue;
        }

        // 更新列
        // Update column
        const maxRows = Math.min(values.length, table.rowCount);
        for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
          const value = values[rowIndex];

          if (skipEmpty && (!value || value.trim() === "")) {
            continue;
          }

          const cell = table.getCellOrNullObject(rowIndex, columnIndex);
          cell.text = value;
          totalUpdated++;
        }
      }

      await context.sync();

      return {
        success: true,
        cellsUpdated: totalUpdated,
        rowCount: table.rowCount,
        columnCount: table.columnCount,
      };
    });
  } catch (error) {
    console.error("批量更新表格列失败 / Failed to batch update table columns:", error);
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
 * 获取表格整行内容
 * Get entire table row content
 *
 * @param rowIndex 行索引 Row index
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<string[]> 行数据 Row data
 */
export async function getTableRowContent(
  rowIndex: number,
  tableLocation?: TableLocationOptions
): Promise<string[]> {
  if (rowIndex < 0) {
    throw new Error("行索引必须大于等于 0 / Row index must be >= 0");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // eslint-disable-next-line office-addins/no-navigational-load
      slide.shapes.load("items");
      await context.sync();

      // 查找表格
      // Find table
      let tableShape: PowerPoint.Shape | undefined;

      if (tableLocation?.shapeId) {
        tableShape = slide.shapes.items.find((shape) => shape.id === tableLocation.shapeId);
        if (!tableShape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
      } else {
        const tables = slide.shapes.items.filter(
          (shape) => shape.type === PowerPoint.ShapeType.table
        );
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

      // 验证行索引
      // Validate row index
      if (rowIndex >= table.rowCount) {
        throw new Error(`行索引 ${rowIndex} 超出范围 / Row index ${rowIndex} out of range`);
      }

      // 获取行数据
      // Get row data
      const cells: PowerPoint.TableCell[] = [];
      for (let colIndex = 0; colIndex < table.columnCount; colIndex++) {
        const cell = table.getCellOrNullObject(rowIndex, colIndex);
        cell.load("text");
        cells.push(cell);
      }

      await context.sync();

      // 填充实际数据
      // Fill actual data
      const rowData: string[] = cells.map((cell) => cell.text);

      return rowData;
    });
  } catch (error) {
    console.error("获取表格行内容失败 / Failed to get table row content:", error);
    throw error;
  }
}

/**
 * 获取表格整列内容
 * Get entire table column content
 *
 * @param columnIndex 列索引 Column index
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<string[]> 列数据 Column data
 */
export async function getTableColumnContent(
  columnIndex: number,
  tableLocation?: TableLocationOptions
): Promise<string[]> {
  if (columnIndex < 0) {
    throw new Error("列索引必须大于等于 0 / Column index must be >= 0");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      // eslint-disable-next-line office-addins/no-navigational-load
      slide.shapes.load("items");
      await context.sync();

      // 查找表格
      // Find table
      let tableShape: PowerPoint.Shape | undefined;

      if (tableLocation?.shapeId) {
        tableShape = slide.shapes.items.find((shape) => shape.id === tableLocation.shapeId);
        if (!tableShape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
      } else {
        const tables = slide.shapes.items.filter(
          (shape) => shape.type === PowerPoint.ShapeType.table
        );
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

      // 验证列索引
      // Validate column index
      if (columnIndex >= table.columnCount) {
        throw new Error(
          `列索引 ${columnIndex} 超出范围 / Column index ${columnIndex} out of range`
        );
      }

      // 获取列数据
      // Get column data
      const cells: PowerPoint.TableCell[] = [];
      for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
        const cell = table.getCellOrNullObject(rowIndex, columnIndex);
        cell.load("text");
        cells.push(cell);
      }

      await context.sync();

      // 填充实际数据
      // Fill actual data
      const columnData: string[] = cells.map((cell) => cell.text);

      return columnData;
    });
  } catch (error) {
    console.error("获取表格列内容失败 / Failed to get table column content:", error);
    throw error;
  }
}
