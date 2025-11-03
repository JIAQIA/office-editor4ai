# Office Editor4AI

Office AddIn for AI - 支持 Excel、Word 和 PowerPoint 的多平台 AI 编辑器插件  
Office AddIn for AI - Multi-platform AI editor add-in supporting Excel, Word, and PowerPoint

## 📋 项目概述 | Project Overview

本项目是一个基于 **pnpm workspace** 的 monorepo 结构，包含三个独立的 Office AddIn 应用：

- **excel-editor4ai**: Excel 插件
- **word-editor4ai**: Word 插件  
- **ppt-editor4ai**: PowerPoint 插件

每个插件都是独立的应用，但共享相同的依赖管理和构建流程。

This project is a **pnpm workspace** based monorepo containing three independent Office AddIn applications:

- **excel-editor4ai**: Excel add-in
- **word-editor4ai**: Word add-in
- **ppt-editor4ai**: PowerPoint add-in

Each add-in is an independent application but shares the same dependency management and build process.

## 🏗️ 项目结构 | Project Structure

```
office-editor4ai/
├── excel-editor4ai/          # Excel 插件 | Excel add-in
│   ├── src/                  # 源代码 | Source code
│   ├── assets/               # 静态资源 | Static assets
│   ├── manifest.xml          # Office 插件清单 | Office add-in manifest
│   ├── package.json          # 依赖配置（端口: 3001）| Dependencies (port: 3001)
│   └── webpack.config.js     # Webpack 配置 | Webpack configuration
│
├── word-editor4ai/           # Word 插件 | Word add-in
│   ├── src/                  # 源代码 | Source code
│   ├── assets/               # 静态资源 | Static assets
│   ├── manifest.xml          # Office 插件清单 | Office add-in manifest
│   ├── package.json          # 依赖配置（端口: 3002）| Dependencies (port: 3002)
│   └── webpack.config.js     # Webpack 配置 | Webpack configuration
│
├── ppt-editor4ai/            # PowerPoint 插件 | PowerPoint add-in
│   ├── src/                  # 源代码 | Source code
│   ├── assets/               # 静态资源 | Static assets
│   ├── manifest.xml          # Office 插件清单 | Office add-in manifest
│   ├── package.json          # 依赖配置（端口: 3003）| Dependencies (port: 3003)
│   └── webpack.config.js     # Webpack 配置 | Webpack configuration
│
├── pnpm-workspace.yaml       # pnpm workspace 配置 | pnpm workspace config
├── package.json              # 根项目配置 | Root project config
└── README.md                 # 项目文档 | Project documentation
```

## 🚀 快速开始 | Quick Start

### 前置要求 | Prerequisites

- **Node.js**: >= 18.0.0
- **pnpm**: >= 8.0.0
- **Office 应用**: Excel、Word 或 PowerPoint（桌面版或 Office 365）

### 安装依赖 | Install Dependencies

```bash
# 安装 pnpm（如果尚未安装）| Install pnpm (if not already installed)
npm install -g pnpm

# 安装所有依赖 | Install all dependencies
pnpm install
```

pnpm workspace 会自动处理所有子项目的依赖安装，并通过符号链接共享公共依赖，大大减少磁盘空间占用。

pnpm workspace automatically handles dependency installation for all sub-projects and shares common dependencies through symbolic links, significantly reducing disk space usage.

## 📦 开发命令 | Development Commands

### 构建项目 | Build Projects

```bash
# 构建所有插件（生产模式）| Build all add-ins (production mode)
pnpm build

# 构建所有插件（开发模式）| Build all add-ins (development mode)
pnpm build:dev
```

### 启动开发服务器 | Start Development Server

每个插件运行在不同的端口上以避免冲突：  
Each add-in runs on a different port to avoid conflicts:

```bash
# 启动 Excel 开发服务器（端口 3001）| Start Excel dev server (port 3001)
pnpm dev:excel

# 启动 Word 开发服务器（端口 3002）| Start Word dev server (port 3002)
pnpm dev:word

# 启动 PowerPoint 开发服务器（端口 3003）| Start PowerPoint dev server (port 3003)
pnpm dev:ppt
```

### 调试插件 | Debug Add-ins

```bash
# 在 Excel 中启动插件 | Start add-in in Excel
pnpm start:excel

# 在 Word 中启动插件 | Start add-in in Word
pnpm start:word

# 在 PowerPoint 中启动插件 | Start add-in in PowerPoint
pnpm start:ppt
```

### 停止调试 | Stop Debugging

```bash
# 停止 Excel 插件 | Stop Excel add-in
pnpm stop:excel

# 停止 Word 插件 | Stop Word add-in
pnpm stop:word

# 停止 PowerPoint 插件 | Stop PowerPoint add-in
pnpm stop:ppt
```

### 验证清单文件 | Validate Manifest

```bash
# 验证单个插件的清单 | Validate individual add-in manifest
pnpm validate:excel
pnpm validate:word
pnpm validate:ppt

# 验证所有插件的清单 | Validate all add-in manifests
pnpm validate:all
```

### 代码检查 | Linting

```bash
# 检查所有插件的代码 | Lint all add-ins
pnpm lint

# 自动修复代码问题 | Auto-fix code issues
pnpm lint:fix
```

### 清理项目 | Clean Project

```bash
# 删除所有 node_modules 和构建产物 | Remove all node_modules and build artifacts
pnpm clean

# 清理 Office AddIn 缓存（解决加载问题）| Clear Office AddIn cache (fixes loading issues)
pnpm clear-cache
```

## 🔧 在子项目中工作 | Working in Sub-projects

如果你需要在特定的插件中执行命令，可以使用 pnpm filter：  
If you need to execute commands in a specific add-in, use pnpm filter:

```bash
# 在 Excel 插件中执行命令 | Execute command in Excel add-in
pnpm --filter excel-editor4ai <command>

# 示例：在 Excel 插件中安装新依赖 | Example: Install new dependency in Excel add-in
pnpm --filter excel-editor4ai add <package-name>

# 示例：在所有插件中安装相同的依赖 | Example: Install same dependency in all add-ins
pnpm -r add <package-name>
```

## 🎯 为什么使用 pnpm workspace？ | Why pnpm Workspace?

### 优势 | Advantages

1. **节省磁盘空间** | **Save Disk Space**  
   通过符号链接共享依赖，避免重复安装相同的包。  
   Share dependencies through symbolic links, avoiding duplicate installations of the same packages.

2. **统一依赖管理** | **Unified Dependency Management**  
   在根目录统一管理所有子项目的依赖版本。  
   Manage dependency versions for all sub-projects from the root directory.

3. **快速安装** | **Fast Installation**  
   pnpm 的安装速度比 npm 和 yarn 更快。  
   pnpm installation is faster than npm and yarn.

4. **严格的依赖隔离** | **Strict Dependency Isolation**  
   避免幽灵依赖问题，确保每个包只能访问声明的依赖。  
   Avoid phantom dependency issues, ensuring each package can only access declared dependencies.

5. **便捷的脚本管理** | **Convenient Script Management**  
   从根目录统一执行所有子项目的命令。  
   Execute commands for all sub-projects from the root directory.

### 与传统方式的对比 | Comparison with Traditional Approach

**传统方式（三个独立项目）**：  
**Traditional Approach (Three Independent Projects)**:
- ❌ 每个项目都有独立的 `node_modules`，占用大量磁盘空间
- ❌ 需要在每个项目目录中分别执行命令
- ❌ 依赖版本可能不一致，导致潜在问题
- ❌ 更新依赖需要在三个项目中分别操作

**pnpm workspace 方式**：  
**pnpm Workspace Approach**:
- ✅ 共享依赖，节省 60-70% 的磁盘空间
- ✅ 从根目录统一管理所有项目
- ✅ 确保依赖版本一致
- ✅ 一次命令更新所有项目

## 🔍 技术栈 | Tech Stack

- **框架** | **Framework**: React 18
- **UI 库** | **UI Library**: Fluent UI React Components
- **构建工具** | **Build Tool**: Webpack 5
- **语言** | **Language**: TypeScript
- **包管理器** | **Package Manager**: pnpm
- **Office API**: Office.js

## 📝 开发注意事项 | Development Notes

1. **端口配置** | **Port Configuration**  
   - Excel: 3001
   - Word: 3002
   - PowerPoint: 3003
   
   请勿修改这些端口，以避免冲突。端口配置需要在三个地方保持一致：  
   Do not modify these ports to avoid conflicts. Port configuration must be consistent in three places:
   - `package.json` 中的 `config.dev_server_port`
   - `manifest.xml` 中的所有 URL
   - `webpack.config.js` 中的 `urlDev`

2. **工作目录问题** | **Working Directory Issue**  
   ⚠️ **重要**: 从根目录运行的命令会自动切换到正确的子目录。如果遇到加载问题，可以直接在子目录中运行命令：  
   ⚠️ **Important**: Commands run from root will automatically switch to the correct subdirectory. If you encounter loading issues, you can run commands directly in subdirectories:
   ```bash
   cd ppt-editor4ai && pnpm start
   ```

3. **清单文件** | **Manifest Files**  
   每个插件都有自己的 `manifest.xml` 文件，用于定义插件的元数据和权限。  
   Each add-in has its own `manifest.xml` file defining metadata and permissions.

4. **Office 缓存** | **Office Cache**  
   如果修改了 `manifest.xml` 或端口配置后插件无法加载，运行 `pnpm clear-cache` 清理 Office 缓存。  
   If the add-in fails to load after modifying `manifest.xml` or port configuration, run `pnpm clear-cache` to clear Office cache.

5. **共享代码** | **Shared Code**  
   如果需要在多个插件之间共享代码，建议创建一个 `packages/shared` 目录，并在 `pnpm-workspace.yaml` 中添加配置。  
   If you need to share code between add-ins, consider creating a `packages/shared` directory and adding it to `pnpm-workspace.yaml`.

6. **调试证书** | **Debug Certificates**  
   首次运行时，Office AddIn 工具会自动生成自签名证书用于 HTTPS 调试。  
   On first run, Office AddIn tools will automatically generate self-signed certificates for HTTPS debugging.

## 🤝 贡献指南 | Contributing

1. 克隆仓库 | Clone the repository
2. 创建功能分支 | Create a feature branch
3. 提交更改 | Commit your changes
4. 推送到分支 | Push to the branch
5. 创建 Pull Request | Create a Pull Request

## 📄 许可证 | License

MIT

## 📧 联系方式 | Contact

如有问题或建议，请提交 Issue。  
For questions or suggestions, please submit an Issue.

---

**最后更新** | **Last Updated**: 2025-11-03  
**维护者** | **Maintainer**: JQQ
