/**
 * 测试文件：验证如何正确获取选中的表格
 * Test file: Verify how to correctly get selected table
 */

/* global Word */

async function testGetSelectedTable() {
  await Word.run(async (context) => {
    // 方法1：通过 getSelection() 获取 Range，然后获取 parentTable
    const range = context.document.getSelection();
    const parentTable = range.parentTableOrNullObject;
    parentTable.load("isNullObject");
    await context.sync();

    console.log("parentTable.isNullObject:", parentTable.isNullObject);

    if (!parentTable.isNullObject) {
      console.log("成功获取到表格对象");
      
      // 测试是否可以直接操作
      parentTable.load("rowCount");
      await context.sync();
      console.log("表格行数:", parentTable.rowCount);

      // 测试删除
      // parentTable.delete();
      // await context.sync();
    } else {
      console.log("光标不在表格内");
    }
  });
}

async function testGetSelectedTableAlternative() {
  await Word.run(async (context) => {
    // 方法2：通过 selection.tables 获取
    const selection = context.document.getSelection();
    const tables = selection.tables;
    tables.load("items");
    await context.sync();

    console.log("selection.tables.items.length:", tables.items.length);

    if (tables.items.length > 0) {
      const table = tables.items[0];
      table.load("rowCount");
      await context.sync();
      console.log("表格行数:", table.rowCount);
    } else {
      console.log("selection.tables 为空");
    }
  });
}

// 导出测试函数
export { testGetSelectedTable, testGetSelectedTableAlternative };
