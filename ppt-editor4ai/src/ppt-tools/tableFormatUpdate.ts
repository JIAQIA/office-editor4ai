/**
 * 文件名: tableFormatUpdate.ts
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 最后修改日期: 2025/12/1
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 表格格式更新工具核心逻辑 - 修改表格/单元格的格式属性
 * Description: Table format update tool core logic - modify table/cell format properties
 */

/* global PowerPoint, console */

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
 * 单元格格式选项
 * Cell format options
 */
export interface CellFormatOptions {
  /** 行索引（从 0 开始）Row index (0-based) */
  rowIndex: number;
  /** 列索引（从 0 开始）Column index (0-based) */
  columnIndex: number;
  /** 背景色（可选，格式: #RRGGBB）Background color (optional, format: #RRGGBB) */
  backgroundColor?: string;
  /** 字体名称（可选）Font name (optional) */
  fontName?: string;
  /** 字体大小（可选，单位: 磅）Font size (optional, in points) */
  fontSize?: number;
  /** 字体颜色（可选，格式: #RRGGBB）Font color (optional, format: #RRGGBB) */
  fontColor?: string;
  /** 字体加粗（可选）Font bold (optional) */
  fontBold?: boolean;
  /** 字体斜体（可选）Font italic (optional) */
  fontItalic?: boolean;
  /** 字体下划线（可选）Font underline (optional) */
  fontUnderline?: boolean;
  /** 边框宽度（可选，单位: 磅）Border width (optional, in points) */
  borderWidth?: number;
  /** 边框颜色（可选，格式: #RRGGBB）Border color (optional, format: #RRGGBB) */
  borderColor?: string;
  /** 水平对齐方式（可选）Horizontal alignment (optional) */
  horizontalAlignment?: "Left" | "Center" | "Right" | "Justify";
  /** 垂直对齐方式（可选）Vertical alignment (optional) */
  verticalAlignment?: "Top" | "Middle" | "Bottom";
}

/**
 * 批量单元格格式更新选项
 * Batch cell format update options
 */
export interface BatchCellFormatOptions {
  /** 单元格格式更新列表 List of cell format updates */
  cells: CellFormatOptions[];
}

/**
 * 行格式选项
 * Row format options
 */
export interface RowFormatOptions {
  /** 行索引（从 0 开始）Row index (0-based) */
  rowIndex: number;
  /** 行高（可选，单位: 磅）Row height (optional, in points) */
  height?: number;
  /** 背景色（可选，格式: #RRGGBB）Background color (optional, format: #RRGGBB) */
  backgroundColor?: string;
  /** 字体名称（可选）Font name (optional) */
  fontName?: string;
  /** 字体大小（可选，单位: 磅）Font size (optional, in points) */
  fontSize?: number;
  /** 字体颜色（可选，格式: #RRGGBB）Font color (optional, format: #RRGGBB) */
  fontColor?: string;
  /** 字体加粗（可选）Font bold (optional) */
  fontBold?: boolean;
  /** 字体斜体（可选）Font italic (optional) */
  fontItalic?: boolean;
}

/**
 * 列格式选项
 * Column format options
 */
export interface ColumnFormatOptions {
  /** 列索引（从 0 开始）Column index (0-based) */
  columnIndex: number;
  /** 列宽（可选，单位: 磅）Column width (optional, in points) */
  width?: number;
  /** 背景色（可选，格式: #RRGGBB）Background color (optional, format: #RRGGBB) */
  backgroundColor?: string;
  /** 字体名称（可选）Font name (optional) */
  fontName?: string;
  /** 字体大小（可选，单位: 磅）Font size (optional, in points) */
  fontSize?: number;
  /** 字体颜色（可选，格式: #RRGGBB）Font color (optional, format: #RRGGBB) */
  fontColor?: string;
  /** 字体加粗（可选）Font bold (optional) */
  fontBold?: boolean;
  /** 字体斜体（可选）Font italic (optional) */
  fontItalic?: boolean;
}

/**
 * 格式更新结果
 * Format update result
 */
export interface FormatUpdateResult {
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
 * 更新单个单元格格式
 * Update single cell format
 *
 * @param formatOptions 单元格格式选项 Cell format options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<FormatUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 更新第 2 行第 3 列的单元格格式
 * // Update cell format at row 2, column 3
 * const result = await updateCellFormat(
 *   {
 *     rowIndex: 1,
 *     columnIndex: 2,
 *     backgroundColor: "#FF0000",
 *     fontSize: 14,
 *     fontBold: true
 *   },
 *   { shapeId: "shape123" }
 * );
 * ```
 */
export async function updateCellFormat(
  formatOptions: CellFormatOptions,
  tableLocation?: TableLocationOptions
): Promise<FormatUpdateResult> {
  const { rowIndex, columnIndex } = formatOptions;

  // 验证参数 Validate parameters
  if (rowIndex < 0 || columnIndex < 0) {
    throw new Error("行索引和列索引必须大于等于 0 / Row and column indices must be >= 0");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      // 查找表格 Find table
      let table: PowerPoint.Table | null = null;

      if (tableLocation?.shapeId) {
        // 通过形状 ID 查找 Find by shape ID
        const shape = shapes.items.find((s) => s.id === tableLocation.shapeId);
        if (!shape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
        shape.load("type");
        await context.sync();

        if (shape.type !== "Table") {
          throw new Error(
            `形状 ${tableLocation.shapeId} 不是表格类型 / Shape ${tableLocation.shapeId} is not a table`
          );
        }
        table = shape.getTable();
      } else {
        // 通过索引查找 Find by index
        const tableIndex = tableLocation?.tableIndex ?? 0;
        let currentTableIndex = 0;

        for (const shape of shapes.items) {
          shape.load("type");
          await context.sync();

          if (shape.type === "Table") {
            if (currentTableIndex === tableIndex) {
              table = shape.getTable();
              break;
            }
            currentTableIndex++;
          }
        }

        if (!table) {
          throw new Error(
            `未找到索引为 ${tableIndex} 的表格 / Table with index ${tableIndex} not found`
          );
        }
      }

      // 加载表格信息 Load table info
      table.load("rowCount, columnCount");
      await context.sync();

      // 验证行列索引 Validate row and column indices
      if (rowIndex >= table.rowCount) {
        throw new Error(
          `行索引 ${rowIndex} 超出范围（表格有 ${table.rowCount} 行）/ Row index ${rowIndex} out of range (table has ${table.rowCount} rows)`
        );
      }
      if (columnIndex >= table.columnCount) {
        throw new Error(
          `列索引 ${columnIndex} 超出范围（表格有 ${table.columnCount} 列）/ Column index ${columnIndex} out of range (table has ${table.columnCount} columns)`
        );
      }

      // 获取单元格 Get cell
      const cell = table.getCellOrNullObject(rowIndex, columnIndex);
      await context.sync();

      if (cell.isNullObject) {
        throw new Error(
          `单元格 (${rowIndex}, ${columnIndex}) 不存在或是合并单元格的一部分 / Cell (${rowIndex}, ${columnIndex}) does not exist or is part of a merged cell`
        );
      }

      // 应用格式 Apply formats
      // 背景色 Background color
      if (formatOptions.backgroundColor) {
        cell.fill.setSolidColor(formatOptions.backgroundColor);
      }

      // 字体格式 Font formats
      if (formatOptions.fontName !== undefined) {
        cell.font.name = formatOptions.fontName;
      }
      if (formatOptions.fontSize !== undefined) {
        cell.font.size = formatOptions.fontSize;
      }
      if (formatOptions.fontColor !== undefined) {
        cell.font.color = formatOptions.fontColor;
      }
      if (formatOptions.fontBold !== undefined) {
        cell.font.bold = formatOptions.fontBold;
      }
      if (formatOptions.fontItalic !== undefined) {
        cell.font.italic = formatOptions.fontItalic;
      }
      if (formatOptions.fontUnderline !== undefined) {
        cell.font.underline = formatOptions.fontUnderline
          ? PowerPoint.ShapeFontUnderlineStyle.single
          : PowerPoint.ShapeFontUnderlineStyle.none;
      }

      // 对齐方式 Alignment
      if (formatOptions.horizontalAlignment !== undefined) {
        cell.horizontalAlignment =
          PowerPoint.ParagraphHorizontalAlignment[formatOptions.horizontalAlignment];
      }
      if (formatOptions.verticalAlignment !== undefined) {
        cell.verticalAlignment = PowerPoint.TextVerticalAlignment[formatOptions.verticalAlignment];
      }

      await context.sync();

      return {
        success: true,
        cellsUpdated: 1,
        rowCount: table.rowCount,
        columnCount: table.columnCount,
      };
    });
  } catch (error) {
    console.error("更新单元格格式失败 / Failed to update cell format:", error);
    return {
      success: false,
      cellsUpdated: 0,
      rowCount: 0,
      columnCount: 0,
      error: error instanceof Error ? error.message : "未知错误 / Unknown error",
    };
  }
}

/**
 * 批量更新单元格格式
 * Batch update cell formats
 *
 * @param batchOptions 批量格式更新选项 Batch format update options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<FormatUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 批量更新多个单元格格式
 * // Batch update multiple cell formats
 * const result = await updateCellFormatsBatch(
 *   {
 *     cells: [
 *       { rowIndex: 0, columnIndex: 0, backgroundColor: "#FF0000", fontBold: true },
 *       { rowIndex: 0, columnIndex: 1, backgroundColor: "#00FF00", fontSize: 16 },
 *       { rowIndex: 1, columnIndex: 0, fontColor: "#0000FF" }
 *     ]
 *   },
 *   { tableIndex: 0 }
 * );
 * ```
 */
export async function updateCellFormatsBatch(
  batchOptions: BatchCellFormatOptions,
  tableLocation?: TableLocationOptions
): Promise<FormatUpdateResult> {
  const { cells } = batchOptions;

  if (!cells || cells.length === 0) {
    throw new Error("单元格列表不能为空 / Cell list cannot be empty");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      // 查找表格 Find table
      let table: PowerPoint.Table | null = null;

      if (tableLocation?.shapeId) {
        const shape = shapes.items.find((s) => s.id === tableLocation.shapeId);
        if (!shape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
        shape.load("type");
        await context.sync();

        if (shape.type !== "Table") {
          throw new Error(
            `形状 ${tableLocation.shapeId} 不是表格类型 / Shape ${tableLocation.shapeId} is not a table`
          );
        }
        table = shape.getTable();
      } else {
        const tableIndex = tableLocation?.tableIndex ?? 0;
        let currentTableIndex = 0;

        for (const shape of shapes.items) {
          shape.load("type");
          await context.sync();

          if (shape.type === "Table") {
            if (currentTableIndex === tableIndex) {
              table = shape.getTable();
              break;
            }
            currentTableIndex++;
          }
        }

        if (!table) {
          throw new Error(
            `未找到索引为 ${tableIndex} 的表格 / Table with index ${tableIndex} not found`
          );
        }
      }

      table.load("rowCount, columnCount");
      await context.sync();

      let updatedCount = 0;

      // 批量更新单元格格式 Batch update cell formats
      for (const cellFormat of cells) {
        const { rowIndex, columnIndex } = cellFormat;

        // 验证索引 Validate indices
        if (
          rowIndex < 0 ||
          rowIndex >= table.rowCount ||
          columnIndex < 0 ||
          columnIndex >= table.columnCount
        ) {
          console.warn(
            `跳过无效单元格 (${rowIndex}, ${columnIndex}) / Skipping invalid cell (${rowIndex}, ${columnIndex})`
          );
          continue;
        }

        const cell = table.getCellOrNullObject(rowIndex, columnIndex);
        await context.sync();

        if (cell.isNullObject) {
          console.warn(
            `跳过合并单元格 (${rowIndex}, ${columnIndex}) / Skipping merged cell (${rowIndex}, ${columnIndex})`
          );
          continue;
        }

        // 应用格式 Apply formats
        if (cellFormat.backgroundColor) {
          cell.fill.setSolidColor(cellFormat.backgroundColor);
        }

        if (cellFormat.fontName !== undefined) {
          cell.font.name = cellFormat.fontName;
        }
        if (cellFormat.fontSize !== undefined) {
          cell.font.size = cellFormat.fontSize;
        }
        if (cellFormat.fontColor !== undefined) {
          cell.font.color = cellFormat.fontColor;
        }
        if (cellFormat.fontBold !== undefined) {
          cell.font.bold = cellFormat.fontBold;
        }
        if (cellFormat.fontItalic !== undefined) {
          cell.font.italic = cellFormat.fontItalic;
        }
        if (cellFormat.fontUnderline !== undefined) {
          cell.font.underline = cellFormat.fontUnderline
            ? PowerPoint.ShapeFontUnderlineStyle.single
            : PowerPoint.ShapeFontUnderlineStyle.none;
        }

        if (cellFormat.horizontalAlignment !== undefined) {
          cell.horizontalAlignment =
            PowerPoint.ParagraphHorizontalAlignment[cellFormat.horizontalAlignment];
        }
        if (cellFormat.verticalAlignment !== undefined) {
          cell.verticalAlignment = PowerPoint.TextVerticalAlignment[cellFormat.verticalAlignment];
        }

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
    console.error("批量更新单元格格式失败 / Failed to batch update cell formats:", error);
    return {
      success: false,
      cellsUpdated: 0,
      rowCount: 0,
      columnCount: 0,
      error: error instanceof Error ? error.message : "未知错误 / Unknown error",
    };
  }
}

/**
 * 更新整行格式
 * Update entire row format
 *
 * @param formatOptions 行格式选项 Row format options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<FormatUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 更新第一行的格式
 * // Update format of first row
 * const result = await updateRowFormat(
 *   {
 *     rowIndex: 0,
 *     backgroundColor: "#CCCCCC",
 *     fontBold: true,
 *     fontSize: 14
 *   },
 *   { tableIndex: 0 }
 * );
 * ```
 */
export async function updateRowFormat(
  formatOptions: RowFormatOptions,
  tableLocation?: TableLocationOptions
): Promise<FormatUpdateResult> {
  const { rowIndex } = formatOptions;

  if (rowIndex < 0) {
    throw new Error("行索引必须大于等于 0 / Row index must be >= 0");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      // 查找表格 Find table
      let table: PowerPoint.Table | null = null;

      if (tableLocation?.shapeId) {
        const shape = shapes.items.find((s) => s.id === tableLocation.shapeId);
        if (!shape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
        shape.load("type");
        await context.sync();

        if (shape.type !== "Table") {
          throw new Error(
            `形状 ${tableLocation.shapeId} 不是表格类型 / Shape ${tableLocation.shapeId} is not a table`
          );
        }
        table = shape.getTable();
      } else {
        const tableIndex = tableLocation?.tableIndex ?? 0;
        let currentTableIndex = 0;

        for (const shape of shapes.items) {
          shape.load("type");
          await context.sync();

          if (shape.type === "Table") {
            if (currentTableIndex === tableIndex) {
              table = shape.getTable();
              break;
            }
            currentTableIndex++;
          }
        }

        if (!table) {
          throw new Error(
            `未找到索引为 ${tableIndex} 的表格 / Table with index ${tableIndex} not found`
          );
        }
      }

      table.load("rowCount, columnCount, rows");
      await context.sync();

      if (rowIndex >= table.rowCount) {
        throw new Error(
          `行索引 ${rowIndex} 超出范围（表格有 ${table.rowCount} 行）/ Row index ${rowIndex} out of range (table has ${table.rowCount} rows)`
        );
      }

      const row = table.rows.getItemAt(rowIndex);

      // 设置行高 Set row height
      if (formatOptions.height !== undefined) {
        row.height = formatOptions.height;
      }

      // 更新行中所有单元格的格式 Update format of all cells in the row
      for (let colIndex = 0; colIndex < table.columnCount; colIndex++) {
        const cell = table.getCellOrNullObject(rowIndex, colIndex);
        await context.sync();

        if (cell.isNullObject) {
          continue; // 跳过合并单元格 Skip merged cells
        }

        if (formatOptions.backgroundColor) {
          cell.fill.setSolidColor(formatOptions.backgroundColor);
        }

        if (formatOptions.fontName !== undefined) {
          cell.font.name = formatOptions.fontName;
        }
        if (formatOptions.fontSize !== undefined) {
          cell.font.size = formatOptions.fontSize;
        }
        if (formatOptions.fontColor !== undefined) {
          cell.font.color = formatOptions.fontColor;
        }
        if (formatOptions.fontBold !== undefined) {
          cell.font.bold = formatOptions.fontBold;
        }
        if (formatOptions.fontItalic !== undefined) {
          cell.font.italic = formatOptions.fontItalic;
        }
      }

      await context.sync();

      return {
        success: true,
        cellsUpdated: table.columnCount,
        rowCount: table.rowCount,
        columnCount: table.columnCount,
      };
    });
  } catch (error) {
    console.error("更新行格式失败 / Failed to update row format:", error);
    return {
      success: false,
      cellsUpdated: 0,
      rowCount: 0,
      columnCount: 0,
      error: error instanceof Error ? error.message : "未知错误 / Unknown error",
    };
  }
}

/**
 * 更新整列格式
 * Update entire column format
 *
 * @param formatOptions 列格式选项 Column format options
 * @param tableLocation 表格定位选项 Table location options
 * @returns Promise<FormatUpdateResult> 更新结果 Update result
 *
 * @example
 * ```typescript
 * // 更新第一列的格式
 * // Update format of first column
 * const result = await updateColumnFormat(
 *   {
 *     columnIndex: 0,
 *     backgroundColor: "#FFFFCC",
 *     fontColor: "#FF0000",
 *     width: 100
 *   },
 *   { tableIndex: 0 }
 * );
 * ```
 */
export async function updateColumnFormat(
  formatOptions: ColumnFormatOptions,
  tableLocation?: TableLocationOptions
): Promise<FormatUpdateResult> {
  const { columnIndex } = formatOptions;

  if (columnIndex < 0) {
    throw new Error("列索引必须大于等于 0 / Column index must be >= 0");
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      // 查找表格 Find table
      let table: PowerPoint.Table | null = null;

      if (tableLocation?.shapeId) {
        const shape = shapes.items.find((s) => s.id === tableLocation.shapeId);
        if (!shape) {
          throw new Error(
            `未找到 ID 为 ${tableLocation.shapeId} 的形状 / Shape with ID ${tableLocation.shapeId} not found`
          );
        }
        shape.load("type");
        await context.sync();

        if (shape.type !== "Table") {
          throw new Error(
            `形状 ${tableLocation.shapeId} 不是表格类型 / Shape ${tableLocation.shapeId} is not a table`
          );
        }
        table = shape.getTable();
      } else {
        const tableIndex = tableLocation?.tableIndex ?? 0;
        let currentTableIndex = 0;

        for (const shape of shapes.items) {
          shape.load("type");
          await context.sync();

          if (shape.type === "Table") {
            if (currentTableIndex === tableIndex) {
              table = shape.getTable();
              break;
            }
            currentTableIndex++;
          }
        }

        if (!table) {
          throw new Error(
            `未找到索引为 ${tableIndex} 的表格 / Table with index ${tableIndex} not found`
          );
        }
      }

      table.load("rowCount, columnCount, columns");
      await context.sync();

      if (columnIndex >= table.columnCount) {
        throw new Error(
          `列索引 ${columnIndex} 超出范围（表格有 ${table.columnCount} 列）/ Column index ${columnIndex} out of range (table has ${table.columnCount} columns)`
        );
      }

      const column = table.columns.getItemAt(columnIndex);

      // 设置列宽 Set column width
      if (formatOptions.width !== undefined) {
        column.width = formatOptions.width;
      }

      // 更新列中所有单元格的格式 Update format of all cells in the column
      for (let rowIdx = 0; rowIdx < table.rowCount; rowIdx++) {
        const cell = table.getCellOrNullObject(rowIdx, columnIndex);
        await context.sync();

        if (cell.isNullObject) {
          continue; // 跳过合并单元格 Skip merged cells
        }

        if (formatOptions.backgroundColor) {
          cell.fill.setSolidColor(formatOptions.backgroundColor);
        }

        if (formatOptions.fontName !== undefined) {
          cell.font.name = formatOptions.fontName;
        }
        if (formatOptions.fontSize !== undefined) {
          cell.font.size = formatOptions.fontSize;
        }
        if (formatOptions.fontColor !== undefined) {
          cell.font.color = formatOptions.fontColor;
        }
        if (formatOptions.fontBold !== undefined) {
          cell.font.bold = formatOptions.fontBold;
        }
        if (formatOptions.fontItalic !== undefined) {
          cell.font.italic = formatOptions.fontItalic;
        }
      }

      await context.sync();

      return {
        success: true,
        cellsUpdated: table.rowCount,
        rowCount: table.rowCount,
        columnCount: table.columnCount,
      };
    });
  } catch (error) {
    console.error("更新列格式失败 / Failed to update column format:", error);
    return {
      success: false,
      cellsUpdated: 0,
      rowCount: 0,
      columnCount: 0,
      error: error instanceof Error ? error.message : "未知错误 / Unknown error",
    };
  }
}
