/**
 * 文件名: eslint.config.js
 * 作者: JQQ
 * 创建日期: 2025/11/03
 * 最后修改日期: 2025/11/03
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: eslint, eslint-plugin-office-addins, @typescript-eslint/parser, typescript-eslint
 * 描述: ESLint Flat Config 配置，启用 Office Add-ins 与 TypeScript 推荐规则。
 *
 * Filename: eslint.config.js
 * Author: JQQ
 * Created: 2025/11/03
 * Last Modified: 2025/11/03
 * Copyright: 2023 JQQ. All rights reserved.
 * Dependencies: eslint, eslint-plugin-office-addins, @typescript-eslint/parser, typescript-eslint
 * Description: ESLint flat config enabling Office Add-ins and TypeScript recommended rules.
 */

// 说明(中文): 使用 Flat Config（ESLint 9+）。确保安装 eslint@^9、typescript-eslint@^7、@typescript-eslint/parser@^7。
// Note (EN): Using Flat Config (ESLint 9+). Ensure eslint@^9, typescript-eslint@^7, and @typescript-eslint/parser@^7 are installed.

// 必须的依赖（用户指定） | Required imports (as requested)
const officeAddins = require("eslint-plugin-office-addins");
const tsParser = require("@typescript-eslint/parser");
const tsEsLint = require("typescript-eslint");

// 其他最佳实践设置（可选） | Additional best practices (optional)
const commonIgnores = ["dist/", "node_modules/", "**/*.d.ts"];

export default [
  // 忽略目录 | Ignore patterns
  {
    ignores: commonIgnores,
  },

  // TypeScript 与 Office 插件推荐配置（保留顺序）
  // TS and Office plugin recommended configs (keep order)
  ...tsEsLint.configs.recommended,
  ...officeAddins.configs.recommended,

  // 用户要求的基础块 | The required base block from user
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
      ecmaVersion: 2022,
      sourceType: "module",
    },
  },

  // 针对 src 的文件匹配与示例规则 | Files matcher and sample rules for src
  {
    files: ["src/**/*.{ts,tsx,js}", "*.ts", "*.js"],
    rules: {
      // 可在此加入项目级规则 | Put project-level rules here
      // "no-console": "warn",
    },
  },
];
