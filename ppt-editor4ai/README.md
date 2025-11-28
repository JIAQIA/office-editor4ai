<!--
文件名: README.md
作者: JQQ
创建日期: 2025/11/28
最后修改日期: 2025/11/28
版权: 2025 JQQ. All rights reserved.
依赖: None
描述: ppt-editor4ai项目说明文档
-->

# ppt-editor4ai 项目说明

## 技术原理

本项⽬是⼀个 PowerPoint Office Add-in，通过 Office JS API 与 PowerPoint 交互。主要功能包括：

1. 在 PowerPoint 任务窗格中提供 UI 界⾯
2. 通过 `PowerPoint.run()` API 访问当前演⽰⽂档
3. 在选中的幻灯⽚上添加⽂本框并设置样式（填充颜⾊、边框等）
4. 使⽤ React 构建任务窗格 UI 组件

## 重要依赖库

- **核⼼依赖**:
  - `@types/office-js`: Office JS API 类型定义
  - `react`/`react-dom`: UI 组件框架
  - `core-js`: JavaScript 标准库 polyfill

- **开发⼯具**:
  - `webpack`: 模块打包⼯具
  - `typescript`: 类型安全的 JavaScript 超集
  - `office-addin-debugging`: Office Add-in 调试⼯具
  - `office-addin-dev-certs`: 开发证书管理

## 开发⽅式

注意无论在哪个目录下，均需要启动两个服务：

1. 开发服务器 对应 pnpm **dev**
2. 启动加载项调试服务 对应 pnpm **start**

如果要调试页面元素样式，建议在浏览器访问：开发服务器，一般为： http://localhost:3003/taskpane.html

### 如果在 ppt-editor4ai 目录下

1. **启动开发服务器**:
   ```bash
   pnpm dev-server
   ```
2. **调试 Add-in**:
   ```bash
   pnpm start
   ```
3. **⽣产环境构建**:
   ```bash
   pnpm build
   ```
4. **代码质量检查**:
   ```bash
   pnpm lint
   ```

### 如果是父目录（项目根目录）下

1. **启动开发服务器**：
   ```bash
   pnpm dev:ppt
   ```
   - 自动进入 ppt-editor4ai 目录
   - 启动 Webpack 开发服务器 (端口 3003)

2. **调试 Add-in**：
   ```bash
   pnpm start:ppt
   ```
   - 注册本地加载项到 PowerPoint
   - 自动打开 PowerPoint 并加载插件

3. **停止调试**：
   ```bash
   pnpm stop:ppt
   ```
   - 从 PowerPoint 卸载开发加载项
   - 停止后台调试进程

4. **验证清单文件**：
   ```bash
   pnpm validate:ppt
   ```
   - 检查 manifest.xml 是否符合 Office 规范
   - 确保所有资源路径正确

5. **同时管理多个加载项**：
   ```bash
   # 启动所有 Office 加载项开发服务器
   pnpm -r --parallel dev-server
   
   # 同时调试 Excel/Word/PPT 加载项
   pnpm start:excel & pnpm start:word & pnpm start:ppt
   ```

## 与⽗项⽬的关系

本项⽬是 `office-editor4ai` monorepo 的⼦项⽬：

1. **独⽴性**:
   - 拥有独⽴的 `manifest.xml` 配置⽂件
   - 独⽴的 Webpack 构建配置 (`webpack.config.js`)
   - 专⻔针对 PowerPoint 的 Office Add-in 实现

2. **共享机制**:
   - 通过⽗级 `package.json` 统⼀管理 monorepo 脚本
   - 共享开发依赖（ESLint/Prettier 配置等）
   - 统⼀的项⽬结构规范

3. **运⾏隔离**:
   - 独⽴的开发服务器端⼝ (3003)
   - 独⽴的 PowerPoint 加载项注册
