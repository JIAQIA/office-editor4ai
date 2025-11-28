/**
 * 文件名: global.d.ts
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: None
 * 描述: 全局类型声明文件 | Global type declarations
 */

// Office.js 类型接口 | Office.js type interfaces
interface OfficeGlobal {
  onReady: (callback?: (info: { host: string; platform: string }) => void) => Promise<{ host: string; platform: string }>;
  context: {
    document: any;
    mailbox: any;
  };
  actions: {
    associate: (actionId: string, handler: () => void) => void;
  };
}

// PowerPoint 类型接口 | PowerPoint type interface
interface PowerPointGlobal {
  run: <T>(callback: (context: any) => Promise<T>) => Promise<T>;
}

declare global {
  // 全局对象类型声明 | Global object type declarations
  var Office: OfficeGlobal;
  var PowerPoint: PowerPointGlobal;
  
  // globalThis 类型扩展 | globalThis type extension
  interface GlobalThis {
    Office: OfficeGlobal;
    PowerPoint: PowerPointGlobal;
  }
}

export {};
