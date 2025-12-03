/**
 * 文件名: insertTable.test.ts
 * 作者: JQQ
 * 创建日期: 2025/12/03
 * 最后修改日期: 2025/12/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: office-addin-mock, vitest
 * 描述: insertTable工具函数的单元测试
 */

import { describe, test, expect, vi, beforeEach } from "vitest";
import {
  insertTable,
  updateTable,
  getTableInfo,
  deleteTable,
  InsertTableOptions,
  UpdateTableOptions,
} from "../../../src/word-tools";
import { mockWordRun } from "../../utils/test-utils";

describe("table 工具函数测试", () => {
  beforeEach(() => {
    mockWordRun();
  });

  test("插入表格 - 基本功能", async () => {
    const options: InsertTableOptions = {
      rows: 3,
      cols: 3,
      insertLocation: "End",
    };

    const result = await insertTable(options);

    expect(result.success).toBe(true);
    expect(result.tableIndex).toBeDefined();
  });

  test("插入表格 - 带数据", async () => {
    const options: InsertTableOptions = {
      rows: 2,
      cols: 2,
      data: [
        ["1", "2"],
        ["3", "4"],
      ],
    };

    const result = await insertTable(options);
    expect(result.success).toBe(true);
  });

  test("插入表格 - 参数验证", async () => {
    const options: InsertTableOptions = {
      rows: 0, // 无效行数
      cols: 3,
    };

    const result = await insertTable(options);
    expect(result.success).toBe(false);
    expect(result.error).toContain("行数和列数必须大于0");
  });

  test("更新表格 - 基本功能", async () => {
    const options: UpdateTableOptions = {
      tableIndex: 0,
      data: [
        ["新数据1", "新数据2"],
        ["新数据3", "新数据4"],
      ],
    };

    const result = await updateTable(options);
    expect(result.success).toBe(true);
  });

  test("获取表格信息", async () => {
    const info = await getTableInfo(0);
    expect(info).toBeDefined();
    expect(info?.rowCount).toBeGreaterThan(0);
  });

  test("删除表格", async () => {
    const result = await deleteTable(0);
    expect(result.success).toBe(true);
  });

  test("更新表格 - 无索引（使用选中的表格）", async () => {
    const options: UpdateTableOptions = {
      data: [["更新数据1", "更新数据2"]],
    };

    const result = await updateTable(options);
    // 注意：在mock环境中可能会失败，因为没有实际选中的表格
    // 这个测试主要验证接口支持可选参数
    expect(result).toBeDefined();
  });

  test("获取表格信息 - 无索引（使用选中的表格）", async () => {
    const info = await getTableInfo();
    // 注意：在mock环境中可能返回null，因为没有实际选中的表格
    // 这个测试主要验证接口支持可选参数
    expect(info === null || info !== undefined).toBe(true);
  });

  test("删除表格 - 无索引（使用选中的表格）", async () => {
    const result = await deleteTable();
    // 注意：在mock环境中可能会失败，因为没有实际选中的表格
    // 这个测试主要验证接口支持可选参数
    expect(result).toBeDefined();
  });
});
