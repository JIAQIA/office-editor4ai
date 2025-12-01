/**
 * 文件名: vitest.workspace.ts
 * 作者: JQQ
 * 创建日期: 2025/12/1
 * 最后修改日期: 2025/12/1
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: vitest
 * 描述: Vitest Workspace 配置文件，用于 monorepo 中的测试识别 | Vitest Workspace configuration for monorepo test recognition
 */

import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    // Vitest 4.x 使用 projects 配置来支持 monorepo | Vitest 4.x uses projects config for monorepo support
    projects: [
      // 使用 glob 模式指向配置文件 | Use glob pattern to point to config files
      './ppt-editor4ai/vitest.config.ts',
      './word-editor4ai/vitest.config.ts',
    ],
  },
})
