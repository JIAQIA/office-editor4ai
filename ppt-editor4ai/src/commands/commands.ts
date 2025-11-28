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
 * 执行加载项命令时显示通知
 * @param event Office 命令事件对象
 */
function action(event: Office.AddinCommands.Event) {
  // 创建通知消息配置
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, // 信息类型通知
    message: "命令执行成功", // 通知内容
    icon: "Icon.80x80", // 图标文件
    persistent: true, // 持久化显示
  };

  // 替换现有通知（ID为"ActionPerformanceNotification"）
  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // 必须调用 event.completed() 表示命令执行完成
  event.completed();
}

// 将 action 函数注册为 Office 命令 "action" 的处理程序
Office.actions.associate("action", action);
