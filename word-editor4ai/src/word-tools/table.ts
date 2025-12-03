/**
 * 文件名: insertTable.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 表格工具核心逻辑（插入、修改、查询、删除表格）
 */

/* global Word, console */

import type { InsertLocation } from "./types";

// 重新导出以保持向后兼容性 / Re-export for backward compatibility
export type { InsertLocation };

/**
 * 表格对齐方式 / Table Alignment
 */
export type TableAlignment = "Left" | "Centered" | "Right" | "Justified";

/**
 * 单元格对齐方式 / Cell Alignment
 */
export type CellAlignment = "Left" | "Centered" | "Right" | "Justified" | "Distributed";

/**
 * 单元格垂直对齐方式 / Cell Vertical Alignment
 */
export type CellVerticalAlignment = "Top" | "Center" | "Bottom";

/**
 * 表格边框样式 / Table Border Style
 */
export type BorderStyle = "Single" | "Dotted" | "Dashed" | "Double" | "None";

/**
 * 表格样式类型（使用 Word API 的内置样式）/ Table Style Type (using Word API built-in styles)
 */
export type TableStyleType = Word.BuiltInStyleName;

/**
 * 单元格格式选项 / Cell Format Options
 */
export interface CellFormatOptions {
  /** 水平对齐方式 / Horizontal alignment */
  alignment?: CellAlignment;
  /** 垂直对齐方式 / Vertical alignment */
  verticalAlignment?: CellVerticalAlignment;
  /** 背景色 / Background color */
  backgroundColor?: string;
  /** 字体名称 / Font name */
  fontName?: string;
  /** 字体大小（磅）/ Font size in points */
  fontSize?: number;
  /** 是否加粗 / Bold */
  bold?: boolean;
  /** 是否斜体 / Italic */
  italic?: boolean;
  /** 字体颜色 / Font color */
  fontColor?: string;
  /** 单元格内边距（磅）/ Cell padding in points */
  cellPadding?: number;
}

/**
 * 表格边框选项 / Table Border Options
 */
export interface TableBorderOptions {
  /** 边框样式 / Border style */
  style?: BorderStyle;
  /** 边框宽度（磅）/ Border width in points */
  width?: number;
  /** 边框颜色 / Border color */
  color?: string;
  /** 是否显示内部边框 / Show inside borders */
  insideBorders?: boolean;
  /** 是否显示外部边框 / Show outside borders */
  outsideBorders?: boolean;
}

/**
 * 表格样式选项 / Table Style Options
 */
export interface TableStyleOptions {
  /** 预设样式类型 / Preset style type */
  styleType?: TableStyleType;
  /** 是否显示首行 / Show first row */
  firstRow?: boolean;
  /** 是否显示末行 / Show last row */
  lastRow?: boolean;
  /** 是否显示首列 / Show first column */
  firstColumn?: boolean;
  /** 是否显示末列 / Show last column */
  lastColumn?: boolean;
  /** 是否使用条纹行 / Use banded rows */
  bandedRows?: boolean;
  /** 是否使用条纹列 / Use banded columns */
  bandedColumns?: boolean;
}

/**
 * 插入表格选项 / Insert Table Options
 */
export interface InsertTableOptions {
  /** 行数（必需）/ Row count (required) */
  rows: number;
  /** 列数（必需）/ Column count (required) */
  cols: number;
  /** 插入位置，默认为 "End" / Insert location, default "End" */
  insertLocation?: InsertLocation;
  /** 表格数据（可选，按行列顺序填充）/ Table data (optional, fill by row and column order) */
  data?: string[][];
  /** 表头数据（可选，第一行数据）/ Header data (optional, first row data) */
  headerRow?: string[];
  /** 列宽（磅），可以是单个值或数组 / Column widths in points, can be single value or array */
  columnWidths?: number | number[];
  /** 表格对齐方式，默认为 "Left" / Table alignment, default "Left" */
  alignment?: TableAlignment;
  /** 表格样式选项 / Table style options */
  styleOptions?: TableStyleOptions;
  /** 表格边框选项 / Table border options */
  borderOptions?: TableBorderOptions;
  /** 表头单元格格式 / Header cell format */
  headerFormat?: CellFormatOptions;
  /** 数据单元格格式 / Data cell format */
  dataFormat?: CellFormatOptions;
  /** 表格标题（用于标识）/ Table title (for identification) */
  title?: string;
  /** 表格描述 / Table description */
  description?: string;
}

/**
 * 更新表格选项 / Update Table Options
 */
export interface UpdateTableOptions {
  /** 表格索引（从0开始，可选）。如果为空，尝试获取当前选中的表格 / Table index (0-based, optional). If empty, try to get the currently selected table */
  tableIndex?: number;
  /** 新的表格数据（可选）/ New table data (optional) */
  data?: string[][];
  /** 表格样式选项 / Table style options */
  styleOptions?: TableStyleOptions;
  /** 表格边框选项 / Table border options */
  borderOptions?: TableBorderOptions;
  /** 列宽（磅）/ Column widths in points */
  columnWidths?: number | number[];
  /** 表格对齐方式 / Table alignment */
  alignment?: TableAlignment;
}

/**
 * 更新单元格选项 / Update Cell Options
 */
export interface UpdateCellOptions {
  /** 表格索引（从0开始）/ Table index (0-based) */
  tableIndex: number;
  /** 行索引（从0开始）/ Row index (0-based) */
  rowIndex: number;
  /** 列索引（从0开始）/ Column index (0-based) */
  columnIndex: number;
  /** 新的单元格内容 / New cell content */
  content?: string;
  /** 单元格格式 / Cell format */
  format?: CellFormatOptions;
}

/**
 * 表格信息 / Table Info
 */
export interface TableInfo {
  /** 表格索引 / Table index */
  index: number;
  /** 行数 / Row count */
  rowCount: number;
  /** 列数 / Column count */
  columnCount: number;
  /** 表格数据 / Table data */
  data?: string[][];
  /** 表格样式 / Table style */
  style?: string;
  /** 表格对齐方式 / Table alignment */
  alignment?: string;
  /** 表格宽度（磅）/ Table width in points */
  width?: number;
}

/**
 * 插入表格结果 / Insert Table Result
 */
export interface InsertTableResult {
  /** 是否成功 / Success */
  success: boolean;
  /** 表格索引（如果成功）/ Table index (if successful) */
  tableIndex?: number;
  /** 错误信息（如果失败）/ Error message (if failed) */
  error?: string;
}

/**
 * 在文档中插入表格
 * Insert table in document
 */
export async function insertTable(options: InsertTableOptions): Promise<InsertTableResult> {
  const {
    rows,
    cols,
    insertLocation = "End",
    data,
    headerRow,
    columnWidths,
    alignment = "Left",
    styleOptions,
    borderOptions,
    headerFormat,
    dataFormat,
    title,
  } = options;

  // 验证参数 / Validate parameters
  if (rows < 1 || cols < 1) {
    return {
      success: false,
      error: "行数和列数必须大于0 / Row and column count must be greater than 0",
    };
  }

  try {
    let tableIndex: number | undefined;

    await Word.run(async (context) => {
      // 获取插入范围 / Get insert range
      let insertRange: Word.Range;
      const selection = context.document.getSelection();

      switch (insertLocation) {
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

      // 插入表格 / Insert table
      // Word API 的 insertTable 只支持 "Before" 和 "After"，其他位置通过 Range 来控制
      // Word API's insertTable only supports "Before" and "After", other positions are controlled by Range
      const apiInsertLocation:
        | Word.InsertLocation.before
        | Word.InsertLocation.after
        | "Before"
        | "After" = insertLocation === "Start" || insertLocation === "Before" ? "Before" : "After";
      // insertTable 的第四个参数需要 string[][]（二维数组）
      // The 4th parameter of insertTable requires string[][] (2D array)
      const initialValues = headerRow ? [headerRow] : data?.[0] ? [data[0]] : undefined;
      const table = insertRange.insertTable(rows, cols, apiInsertLocation, initialValues);

      // 设置表格对齐方式 / Set table alignment
      table.alignment = alignment as Word.Alignment;

      // 填充表格数据 / Fill table data
      if (data && data.length > 0) {
        const startRow = headerRow ? 1 : 0; // 如果有表头，从第二行开始填充 / If has header, start from second row
        for (let i = 0; i < data.length && startRow + i < rows; i++) {
          const rowData = data[i];
          for (let j = 0; j < rowData.length && j < cols; j++) {
            const cell = table.getCell(startRow + i, j);
            cell.value = rowData[j] || "";
          }
        }
      }

      // 设置列宽 / Set column widths
      if (columnWidths !== undefined) {
        table.columns.load("items");
        await context.sync();

        if (typeof columnWidths === "number") {
          // 所有列使用相同宽度 / All columns use same width
          for (let j = 0; j < cols; j++) {
            table.columns.items[j].width = columnWidths;
          }
        } else if (Array.isArray(columnWidths)) {
          // 每列使用不同宽度 / Each column uses different width
          for (let j = 0; j < Math.min(columnWidths.length, cols); j++) {
            table.columns.items[j].width = columnWidths[j];
          }
        }
      }

      // 应用表格样式 / Apply table style
      if (styleOptions) {
        if (styleOptions.styleType) {
          // 应用预设样式 / Apply preset style
          table.styleBuiltIn = styleOptions.styleType;
        }
        if (styleOptions.firstRow !== undefined) {
          table.styleFirstColumn = styleOptions.firstRow;
        }
        if (styleOptions.lastRow !== undefined) {
          table.styleTotalRow = styleOptions.lastRow;
        }
        if (styleOptions.firstColumn !== undefined) {
          table.styleFirstColumn = styleOptions.firstColumn;
        }
        if (styleOptions.bandedRows !== undefined) {
          table.styleBandedRows = styleOptions.bandedRows;
        }
        if (styleOptions.bandedColumns !== undefined) {
          table.styleBandedColumns = styleOptions.bandedColumns;
        }
      }

      // 应用边框样式 / Apply border style
      if (borderOptions) {
        const borders = table.getBorder(Word.BorderLocation.all);
        if (borderOptions.style) {
          borders.type = borderOptions.style as Word.BorderType;
        }
        if (borderOptions.width !== undefined) {
          borders.width = borderOptions.width;
        }
        if (borderOptions.color) {
          borders.color = borderOptions.color;
        }
      }

      // 应用表头格式 / Apply header format
      if (headerFormat && (headerRow || data?.[0])) {
        const headerRowObj = table.rows.getFirst();
        // eslint-disable-next-line office-addins/no-navigational-load
        headerRowObj.load("cells");
        await context.sync();
        await applyCellFormat(headerRowObj.cells, headerFormat, context);
      }

      // 应用数据格式 / Apply data format
      if (dataFormat) {
        table.rows.load("items");
        await context.sync();

        const startRow = headerRow || data?.[0] ? 1 : 0;
        for (let i = startRow; i < rows; i++) {
          const row = table.rows.items[i];
          await applyCellFormat(row.cells, dataFormat, context);
        }
      }

      // 设置表格标题和描述（使用表格的第一个单元格的注释或自定义属性）
      // Set table title and description (using comment or custom property of first cell)
      if (title) {
        // 使用 Word API 的 title 属性而不是直接修改 values
        // Use Word API's title property instead of modifying values directly
        table.title = title;
        // Note: Word JavaScript API 对表格元数据的支持有限
        // Word JavaScript API has limited support for table metadata
      }

      await context.sync();

      // 获取表格索引 / Get table index
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      tableIndex = tables.items.findIndex((t) => t === table);
    });

    return {
      success: true,
      tableIndex,
    };
  } catch (error) {
    console.error("插入表格失败 / Insert table failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 应用单元格格式
 * Apply cell format
 */
async function applyCellFormat(
  cells: Word.TableCellCollection,
  format: CellFormatOptions,
  context: Word.RequestContext
): Promise<void> {
  cells.load("items");
  await context.sync();

  for (const cell of cells.items) {
    if (format.alignment) {
      cell.horizontalAlignment = format.alignment as Word.Alignment;
    }
    if (format.verticalAlignment) {
      cell.verticalAlignment = format.verticalAlignment as Word.VerticalAlignment;
    }
    if (format.backgroundColor) {
      cell.shadingColor = format.backgroundColor;
    }

    // 应用字体格式 / Apply font format
    cell.body.load("font");
    const font = cell.body.font;
    if (format.fontName) {
      font.name = format.fontName;
    }
    if (format.fontSize !== undefined) {
      font.size = format.fontSize;
    }
    if (format.bold !== undefined) {
      font.bold = format.bold;
    }
    if (format.italic !== undefined) {
      font.italic = format.italic;
    }
    if (format.fontColor) {
      font.color = format.fontColor;
    }
  }
}

/**
 * 更新表格
 * Update table
 */
export async function updateTable(options: UpdateTableOptions): Promise<InsertTableResult> {
  const { tableIndex, data, styleOptions, borderOptions, columnWidths, alignment } = options;

  try {
    let actualTableIndex: number | undefined;

    await Word.run(async (context) => {
      let table: Word.Table;

      // 如果没有提供表格索引，尝试获取选中的表格 / If no table index provided, try to get selected table
      if (tableIndex === undefined) {
        const selection = context.document.getSelection();
        // eslint-disable-next-line office-addins/no-navigational-load
        selection.load("parentTableOrNullObject");
        await context.sync();

        // 尝试获取父表格 / Try to get parent table
        const parentTable = selection.parentTableOrNullObject;
        parentTable.load("isNullObject");
        await context.sync();

        if (parentTable.isNullObject) {
          throw new Error("光标未在任何表格内 / Cursor is not inside any table");
        }

        // 直接使用 parentTable / Use parentTable directly
        table = parentTable;
        actualTableIndex = -1; // 无法确定索引 / Cannot determine index
      } else {
        // 使用提供的索引 / Use provided index
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`
          );
        }

        table = tables.items[tableIndex];
        actualTableIndex = tableIndex;
      }
      table.load("rowCount");
      await context.sync();

      // 更新表格数据 / Update table data
      if (data && data.length > 0) {
        table.columns.load("items");
        await context.sync();
        const columnCount = table.columns.items.length;

        for (let i = 0; i < data.length && i < table.rowCount; i++) {
          const rowData = data[i];
          for (let j = 0; j < rowData.length && j < columnCount; j++) {
            const cell = table.getCell(i, j);
            cell.value = rowData[j] || "";
          }
        }
      }

      // 更新列宽 / Update column widths
      if (columnWidths !== undefined) {
        table.columns.load("items");
        await context.sync();
        const columnCount = table.columns.items.length;

        if (typeof columnWidths === "number") {
          for (let j = 0; j < columnCount; j++) {
            table.columns.items[j].width = columnWidths;
          }
        } else if (Array.isArray(columnWidths)) {
          for (let j = 0; j < Math.min(columnWidths.length, columnCount); j++) {
            table.columns.items[j].width = columnWidths[j];
          }
        }
      }

      // 更新对齐方式 / Update alignment
      if (alignment) {
        table.alignment = alignment as Word.Alignment;
      }

      // 更新样式 / Update style
      if (styleOptions) {
        if (styleOptions.styleType) {
          table.styleBuiltIn = styleOptions.styleType;
        }
        if (styleOptions.firstRow !== undefined) {
          table.styleFirstColumn = styleOptions.firstRow;
        }
        if (styleOptions.lastRow !== undefined) {
          table.styleTotalRow = styleOptions.lastRow;
        }
        if (styleOptions.bandedRows !== undefined) {
          table.styleBandedRows = styleOptions.bandedRows;
        }
        if (styleOptions.bandedColumns !== undefined) {
          table.styleBandedColumns = styleOptions.bandedColumns;
        }
      }

      // 更新边框 / Update borders
      if (borderOptions) {
        const borders = table.getBorder(Word.BorderLocation.all);
        if (borderOptions.style) {
          borders.type = borderOptions.style as Word.BorderType;
        }
        if (borderOptions.width !== undefined) {
          borders.width = borderOptions.width;
        }
        if (borderOptions.color) {
          borders.color = borderOptions.color;
        }
      }

      await context.sync();
    });

    return {
      success: true,
      tableIndex: actualTableIndex,
    };
  } catch (error) {
    console.error("更新表格失败 / Update table failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 更新单元格
 * Update cell
 */
export async function updateCell(options: UpdateCellOptions): Promise<InsertTableResult> {
  const { tableIndex, rowIndex, columnIndex, content, format } = options;

  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        throw new Error(`表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`);
      }

      const table = tables.items[tableIndex];
      table.load("rowCount");
      table.columns.load("items");
      await context.sync();

      if (rowIndex < 0 || rowIndex >= table.rowCount) {
        throw new Error(`行索引 ${rowIndex} 超出范围 / Row index ${rowIndex} out of range`);
      }

      const columnCount = table.columns.items.length;
      if (columnIndex < 0 || columnIndex >= columnCount) {
        throw new Error(
          `列索引 ${columnIndex} 超出范围 / Column index ${columnIndex} out of range`
        );
      }

      const cell = table.getCell(rowIndex, columnIndex);

      // 更新内容 / Update content
      if (content !== undefined) {
        cell.value = content;
      }

      // 更新格式 / Update format
      if (format) {
        if (format.alignment) {
          cell.horizontalAlignment = format.alignment as Word.Alignment;
        }
        if (format.verticalAlignment) {
          cell.verticalAlignment = format.verticalAlignment as Word.VerticalAlignment;
        }
        if (format.backgroundColor) {
          cell.shadingColor = format.backgroundColor;
        }

        // 应用字体格式 / Apply font format
        // eslint-disable-next-line office-addins/no-navigational-load
        cell.load("body/font");
        await context.sync();
        const font = cell.body.font;
        if (format.fontName) {
          font.name = format.fontName;
        }
        if (format.fontSize !== undefined) {
          font.size = format.fontSize;
        }
        if (format.bold !== undefined) {
          font.bold = format.bold;
        }
        if (format.italic !== undefined) {
          font.italic = format.italic;
        }
        if (format.fontColor) {
          font.color = format.fontColor;
        }
      }

      await context.sync();
    });

    return {
      success: true,
      tableIndex,
    };
  } catch (error) {
    console.error("更新单元格失败 / Update cell failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 获取表格信息
 * Get table info
 * @param tableIndex 表格索引（可选）。如果为空，尝试获取当前选中的表格 / Table index (optional). If empty, try to get the currently selected table
 */
export async function getTableInfo(tableIndex?: number): Promise<TableInfo | null> {
  try {
    let tableInfo: TableInfo | null = null;

    await Word.run(async (context) => {
      let table: Word.Table;
      let actualTableIndex: number;

      // 如果没有提供表格索引，尝试获取选中的表格 / If no table index provided, try to get selected table
      if (tableIndex === undefined) {
        const selection = context.document.getSelection();
        // eslint-disable-next-line office-addins/no-navigational-load
        selection.load("parentTableOrNullObject");
        await context.sync();

        // 尝试获取父表格 / Try to get parent table
        const parentTable = selection.parentTableOrNullObject;
        parentTable.load("isNullObject");
        await context.sync();

        if (parentTable.isNullObject) {
          throw new Error("光标未在任何表格内 / Cursor is not inside any table");
        }

        // 直接使用 parentTable / Use parentTable directly
        table = parentTable;
        actualTableIndex = -1; // 无法确定索引 / Cannot determine index
      } else {
        // 使用提供的索引 / Use provided index
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`
          );
        }

        table = tables.items[tableIndex];
        actualTableIndex = tableIndex;
      }

      table.load("rowCount, style, alignment, width, values");
      table.columns.load("items");
      await context.sync();

      tableInfo = {
        index: actualTableIndex,
        rowCount: table.rowCount,
        columnCount: table.columns.items.length,
        data: table.values as string[][],
        style: table.style,
        alignment: table.alignment,
        width: table.width,
      };
    });

    return tableInfo;
  } catch (error) {
    console.error("获取表格信息失败 / Get table info failed:", error);
    return null;
  }
}

/**
 * 获取所有表格信息
 * Get all tables info
 */
export async function getAllTablesInfo(): Promise<TableInfo[]> {
  try {
    const tablesInfo: TableInfo[] = [];

    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      // 批量加载所有表格数据以避免循环中的sync / Batch load all table data to avoid sync in loop
      for (const table of tables.items) {
        table.load("rowCount, style, alignment, width, values");
        table.columns.load("items");
      }
      await context.sync();

      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        tablesInfo.push({
          index: i,
          rowCount: table.rowCount,
          columnCount: table.columns.items.length,
          data: table.values as string[][],
          style: table.style,
          alignment: table.alignment,
          width: table.width,
        });
      }
    });

    return tablesInfo;
  } catch (error) {
    console.error("获取所有表格信息失败 / Get all tables info failed:", error);
    return [];
  }
}

/**
 * 删除表格
 * Delete table
 * @param tableIndex 表格索引（可选）。如果为空，尝试删除当前选中的表格 / Table index (optional). If empty, try to delete the currently selected table
 */
export async function deleteTable(tableIndex?: number): Promise<InsertTableResult> {
  try {
    let actualTableIndex: number | undefined;

    await Word.run(async (context) => {
      let table: Word.Table | undefined;

      // 如果没有提供表格索引，尝试获取选中的表格 / If no table index provided, try to get selected table
      if (tableIndex === undefined) {
        const selection = context.document.getSelection();
        // eslint-disable-next-line office-addins/no-navigational-load
        selection.load("parentTableOrNullObject");
        await context.sync();

        // 尝试获取父表格 / Try to get parent table
        const parentTable = selection.parentTableOrNullObject;
        parentTable.load("isNullObject");
        await context.sync();

        if (parentTable.isNullObject) {
          throw new Error("光标未在任何表格内 / Cursor is not inside any table");
        }

        // 加载表格的基本属性以确保对象完整 / Load basic properties to ensure object is complete
        parentTable.load("rowCount");
        await context.sync();

        // 删除表格 / Delete table
        parentTable.delete();

        // 获取该表格在文档中的索引（用于返回值）/ Get the table index in document (for return value)
        const allTables = context.document.body.tables;
        allTables.load("items");
        await context.sync();

        // 由于表格已被删除，我们无法通过引用查找索引，返回 -1 表示未知索引
        // Since table is deleted, we cannot find index by reference, return -1 for unknown index
        actualTableIndex = -1;
      } else {
        // 使用提供的索引 / Use provided index
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`
          );
        }

        table = tables.items[tableIndex];
        actualTableIndex = tableIndex;

        table.delete();
      }

      await context.sync();
    });

    return {
      success: true,
      tableIndex: actualTableIndex,
    };
  } catch (error) {
    console.error("删除表格失败 / Delete table failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 在表格中添加行
 * Add rows to table
 * @param tableIndex 表格索引（可选）。如果为空，尝试使用当前选中的表格 / Table index (optional). If empty, try to use the currently selected table
 * @param rowCount
 * @param insertAt
 * @param values
 */
export async function addTableRows(
  tableIndex: number | undefined,
  rowCount: number,
  insertAt: "Start" | "End" = "End",
  values?: string[][]
): Promise<InsertTableResult> {
  try {
    let actualTableIndex: number | undefined;

    await Word.run(async (context) => {
      let table: Word.Table;

      // 如果没有提供表格索引，尝试获取选中的表格 / If no table index provided, try to get selected table
      if (tableIndex === undefined) {
        const selection = context.document.getSelection();
        // eslint-disable-next-line office-addins/no-navigational-load
        selection.load("parentTableOrNullObject");
        await context.sync();

        const parentTable = selection.parentTableOrNullObject;
        parentTable.load("isNullObject");
        await context.sync();

        if (parentTable.isNullObject) {
          throw new Error("光标未在任何表格内 / Cursor is not inside any table");
        }

        table = parentTable;
        actualTableIndex = undefined;
      } else {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`
          );
        }

        table = tables.items[tableIndex];
        actualTableIndex = tableIndex;
      }

      table.load(["rowCount"]);
      table.columns.load("items");
      await context.sync();
      const columnCount = table.columns.items.length;
      const originalRowCount = table.rowCount;

      // 添加行 / Add rows
      table.addRows(insertAt, rowCount);
      await context.sync();

      // 填充值 / Fill values
      if (values) {
        for (let i = 0; i < rowCount && i < values.length; i++) {
          const rowIndex = insertAt === "Start" ? i : originalRowCount + i;
          for (let j = 0; j < Math.min(values[i].length, columnCount); j++) {
            const cell = table.getCell(rowIndex, j);
            cell.value = values[i][j] || "";
          }
        }
      }

      await context.sync();
    });

    return {
      success: true,
      tableIndex: actualTableIndex,
    };
  } catch (error) {
    console.error("添加表格行失败 / Add table rows failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 在表格中添加列
 * Add columns to table
 * @param tableIndex 表格索引（可选）。如果为空，尝试使用当前选中的表格 / Table index (optional). If empty, try to use the currently selected table
 * @param columnCount
 * @param insertAt
 * @param values
 */
export async function addTableColumns(
  tableIndex: number | undefined,
  columnCount: number,
  insertAt: "Start" | "End" = "End",
  values?: string[][]
): Promise<InsertTableResult> {
  try {
    let actualTableIndex: number | undefined;

    await Word.run(async (context) => {
      let table: Word.Table;

      // 如果没有提供表格索引，尝试获取选中的表格 / If no table index provided, try to get selected table
      if (tableIndex === undefined) {
        const selection = context.document.getSelection();
        // eslint-disable-next-line office-addins/no-navigational-load
        selection.load("parentTableOrNullObject");
        await context.sync();

        const parentTable = selection.parentTableOrNullObject;
        parentTable.load("isNullObject");
        await context.sync();

        if (parentTable.isNullObject) {
          throw new Error("光标未在任何表格内 / Cursor is not inside any table");
        }

        table = parentTable;
        actualTableIndex = undefined;
      } else {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`
          );
        }

        table = tables.items[tableIndex];
        actualTableIndex = tableIndex;
      }

      table.load(["rowCount"]);
      table.columns.load("items");
      await context.sync();
      const originalColumnCount = table.columns.items.length;
      const rowCountInTable = table.rowCount;

      // 添加列 / Add columns
      table.addColumns(insertAt, columnCount);
      await context.sync();

      // 填充值 / Fill values
      if (values) {
        for (let i = 0; i < columnCount && i < values.length; i++) {
          const colIndex = insertAt === "Start" ? i : originalColumnCount + i;
          for (let j = 0; j < Math.min(values[i].length, rowCountInTable); j++) {
            const cell = table.getCell(j, colIndex);
            cell.value = values[i][j] || "";
          }
        }
      }

      await context.sync();
    });

    return {
      success: true,
      tableIndex: actualTableIndex,
    };
  } catch (error) {
    console.error("添加表格列失败 / Add table columns failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 删除表格行
 * Delete table rows
 * @param tableIndex 表格索引（可选）。如果为空，尝试使用当前选中的表格 / Table index (optional). If empty, try to use the currently selected table
 * @param startRowIndex
 * @param rowCount
 */
export async function deleteTableRows(
  tableIndex: number | undefined,
  startRowIndex: number,
  rowCount: number = 1
): Promise<InsertTableResult> {
  try {
    let actualTableIndex: number | undefined;

    await Word.run(async (context) => {
      let table: Word.Table;

      // 如果没有提供表格索引，尝试获取选中的表格 / If no table index provided, try to get selected table
      if (tableIndex === undefined) {
        const selection = context.document.getSelection();
        // eslint-disable-next-line office-addins/no-navigational-load
        selection.load("parentTableOrNullObject");
        await context.sync();

        const parentTable = selection.parentTableOrNullObject;
        parentTable.load("isNullObject");
        await context.sync();

        if (parentTable.isNullObject) {
          throw new Error("光标未在任何表格内 / Cursor is not inside any table");
        }

        table = parentTable;
        actualTableIndex = undefined;
      } else {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`
          );
        }

        table = tables.items[tableIndex];
        actualTableIndex = tableIndex;
      }

      table.load("rowCount");
      await context.sync();

      if (startRowIndex < 0 || startRowIndex >= table.rowCount) {
        throw new Error(
          `行索引 ${startRowIndex} 超出范围 / Row index ${startRowIndex} out of range`
        );
      }

      // 删除行 / Delete rows
      for (let i = 0; i < rowCount && startRowIndex < table.rowCount; i++) {
        const row = table.rows.items[startRowIndex];
        row.delete();
      }

      await context.sync();
    });

    return {
      success: true,
      tableIndex: actualTableIndex,
    };
  } catch (error) {
    console.error("删除表格行失败 / Delete table rows failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 删除表格列
 * Delete table columns
 * @param tableIndex 表格索引（可选）。如果为空，尝试使用当前选中的表格 / Table index (optional). If empty, try to use the currently selected table
 * @param startColumnIndex
 * @param columnCount
 */
export async function deleteTableColumns(
  tableIndex: number | undefined,
  startColumnIndex: number,
  columnCount: number = 1
): Promise<InsertTableResult> {
  try {
    let actualTableIndex: number | undefined;

    await Word.run(async (context) => {
      let table: Word.Table;

      // 如果没有提供表格索引，尝试获取选中的表格 / If no table index provided, try to get selected table
      if (tableIndex === undefined) {
        const selection = context.document.getSelection();
        // eslint-disable-next-line office-addins/no-navigational-load
        selection.load("parentTableOrNullObject");
        await context.sync();

        const parentTable = selection.parentTableOrNullObject;
        parentTable.load("isNullObject");
        await context.sync();

        if (parentTable.isNullObject) {
          throw new Error("光标未在任何表格内 / Cursor is not inside any table");
        }

        table = parentTable;
        actualTableIndex = undefined;
      } else {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`
          );
        }

        table = tables.items[tableIndex];
        actualTableIndex = tableIndex;
      }

      table.columns.load("items");
      await context.sync();
      const tableColumnCount = table.columns.items.length;

      if (startColumnIndex < 0 || startColumnIndex >= tableColumnCount) {
        throw new Error(
          `列索引 ${startColumnIndex} 超出范围 / Column index ${startColumnIndex} out of range`
        );
      }

      // 删除列 / Delete columns
      for (let i = 0; i < columnCount && startColumnIndex < tableColumnCount; i++) {
        const column = table.columns.items[startColumnIndex];
        column.delete();
      }

      await context.sync();
    });

    return {
      success: true,
      tableIndex: actualTableIndex,
    };
  } catch (error) {
    console.error("删除表格列失败 / Delete table columns failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * 合并单元格
 * Merge cells
 * @param tableIndex 表格索引（可选）。如果为空，尝试使用当前选中的表格 / Table index (optional). If empty, try to use the currently selected table
 * @param startRowIndex
 * @param startColumnIndex
 * @param endRowIndex
 * @param endColumnIndex
 */
export async function mergeCells(
  tableIndex: number | undefined,
  startRowIndex: number,
  startColumnIndex: number,
  endRowIndex: number,
  endColumnIndex: number
): Promise<InsertTableResult> {
  try {
    let actualTableIndex: number | undefined;

    await Word.run(async (context) => {
      let table: Word.Table;

      // 如果没有提供表格索引，尝试获取选中的表格 / If no table index provided, try to get selected table
      if (tableIndex === undefined) {
        const selection = context.document.getSelection();
        // eslint-disable-next-line office-addins/no-navigational-load
        selection.load("parentTableOrNullObject");
        await context.sync();

        const parentTable = selection.parentTableOrNullObject;
        parentTable.load("isNullObject");
        await context.sync();

        if (parentTable.isNullObject) {
          throw new Error("光标未在任何表格内 / Cursor is not inside any table");
        }

        table = parentTable;
        actualTableIndex = undefined;
      } else {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tableIndex < 0 || tableIndex >= tables.items.length) {
          throw new Error(
            `表格索引 ${tableIndex} 超出范围 / Table index ${tableIndex} out of range`
          );
        }

        table = tables.items[tableIndex];
        actualTableIndex = tableIndex;
      }

      table.load("rowCount");
      table.columns.load("items");
      await context.sync();
      const columnCount = table.columns.items.length;

      // 验证索引 / Validate indices
      if (
        startRowIndex < 0 ||
        startRowIndex >= table.rowCount ||
        endRowIndex < 0 ||
        endRowIndex >= table.rowCount ||
        startColumnIndex < 0 ||
        startColumnIndex >= columnCount ||
        endColumnIndex < 0 ||
        endColumnIndex >= columnCount
      ) {
        throw new Error("单元格索引超出范围 / Cell indices out of range");
      }

      // 获取起始和结束单元格 / Get start and end cells
      const startCell = table.getCell(startRowIndex, startColumnIndex);
      const endCell = table.getCell(endRowIndex, endColumnIndex);

      // 合并单元格 / Merge cells
      startCell.merge(endCell);

      await context.sync();
    });

    return {
      success: true,
      tableIndex: actualTableIndex,
    };
  } catch (error) {
    console.error("合并单元格失败 / Merge cells failed:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
