/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// 当 Office 加载项准备就绪时执行
Office.onReady(() => {
  // 如果需要，可以在此处调用 Office.js API
});

/**
 * 执行加载项命令时的处理函数
 * @param event Office 命令事件对象
 */
function action(event: Office.AddinCommands.Event) {
  // 测试通知功能 - 在 PowerPoint 中显示对话框
  Office.context.ui.displayDialogAsync(
    // TODO 未来这里需要替换成线上聊天界面
    'https://localhost:3003/taskpane.html',
    { height: 50, width: 30 },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('打开对话框失败:', result.error.message);
      } else {
        console.log('对话框打开成功');
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
          dialog.close();
        });
      }
      // 必须调用 completed 来通知 Office 命令已完成
      event.completed();
    }
  );
}

// 将 action 函数注册为 Office 命令 "action" 的处理程序
Office.actions.associate("action", action);
