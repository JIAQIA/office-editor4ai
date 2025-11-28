/**
* 文件名: eslint.config.mjs
* 作者: JQQ
* 创建日期: 2025/11/28
* 最后修改日期: 2025/11/28
* 版权: 2023 JQQ. All rights reserved.
* 依赖: eslint-plugin-office-addins, typescript-eslint
* 描述: Excel Editor ESLint 配置文件 / Excel Editor ESLint configuration file
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
);
