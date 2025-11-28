/**
* 文件名: eslint.config.mjs
* 作者: JQQ
* 创建日期: 2025/11/28
* 最后修改日期: 2025/11/28
* 版权: 2023 JQQ. All rights reserved.
* 依赖: eslint-plugin-office-addins, typescript-eslint
* 描述: Word Editor ESLint 配置文件 / Word Editor ESLint configuration file
*/

import tseslint from "typescript-eslint";
import officeAddins from "eslint-plugin-office-addins";

export default tseslint.config(
  {
    files: ["src/**/*.{ts,tsx}"],
  },
  ...tseslint.configs.recommended,
  ...(Array.isArray(officeAddins.configs.react) 
    ? officeAddins.configs.react 
    : [officeAddins.configs.react]
  ),
  {
    rules: {
      // 在 TypeScript 文件中禁用基础 ESLint 的 no-unused-vars 规则
      // Disable base ESLint no-unused-vars rule in TypeScript files
      // 因为 @typescript-eslint/no-unused-vars 提供了更好的 TypeScript 支持
      // Because @typescript-eslint/no-unused-vars provides better TypeScript support
      "no-unused-vars": "off",
      // 允许下划线开头的未使用变量（包括接口中的参数占位符）
      // Allow unused variables starting with underscore (including parameter placeholders in interfaces)
      "@typescript-eslint/no-unused-vars": [
        "error",
        {
          "argsIgnorePattern": "^_",
          "varsIgnorePattern": "^_"
        }
      ]
    }
  }
);
