# 开发者指南 | Developer Guide

## 🎯 快速参考 | Quick Reference

### 常用命令速查表 | Common Commands Cheat Sheet

| 操作 Operation | 命令 Command | 说明 Description |
|---|---|---|
| 安装依赖 | `pnpm install` | 安装所有子项目依赖 |
| 开发 Excel | `pnpm dev:excel` | 启动 Excel 开发服务器 (3001) |
| 开发 Word | `pnpm dev:word` | 启动 Word 开发服务器 (3002) |
| 开发 PPT | `pnpm dev:ppt` | 启动 PPT 开发服务器 (3003) |
| 构建所有 | `pnpm build` | 生产环境构建所有插件 |
| 代码检查 | `pnpm lint` | 检查所有代码 |
| 清理项目 | `pnpm clean` | 删除所有 node_modules |

### 端口分配 | Port Allocation

- **Excel**: `http://localhost:3001`
- **Word**: `http://localhost:3002`
- **PowerPoint**: `http://localhost:3003`

## 🔧 项目架构说明 | Architecture Explanation

### 为什么是 Monorepo？ | Why Monorepo?

由于 `yo office` 脚手架工具不支持一次性创建多平台 AddIn，我们采用了以下方案：

1. 使用 `yo office` 分别创建三个独立的 AddIn 项目
2. 使用 pnpm workspace 将它们组织成 monorepo 结构
3. 共享依赖，统一管理

Since the `yo office` scaffolding tool doesn't support creating multi-platform AddIns at once, we adopted this approach:

1. Use `yo office` to create three independent AddIn projects separately
2. Use pnpm workspace to organize them into a monorepo structure
3. Share dependencies and manage them uniformly

### 依赖管理策略 | Dependency Management Strategy

```
office-editor4ai/
├── node_modules/              # 根级别共享依赖 | Root-level shared dependencies
│   ├── react/                 # 所有子项目共享 | Shared by all sub-projects
│   ├── typescript/
│   └── ...
├── excel-editor4ai/
│   └── node_modules/          # Excel 特有依赖的符号链接 | Symlinks to Excel-specific deps
├── word-editor4ai/
│   └── node_modules/          # Word 特有依赖的符号链接 | Symlinks to Word-specific deps
└── ppt-editor4ai/
    └── node_modules/          # PPT 特有依赖的符号链接 | Symlinks to PPT-specific deps
```

pnpm 会自动：
- 将公共依赖提升到根目录
- 为每个子项目创建符号链接
- 确保依赖隔离和版本一致性

pnpm automatically:
- Hoists common dependencies to the root
- Creates symbolic links for each sub-project
- Ensures dependency isolation and version consistency

## 🚀 开发工作流 | Development Workflow

### 1. 初次设置 | Initial Setup

```bash
# 克隆项目 | Clone project
git clone <repository-url>
cd office-editor4ai

# 安装依赖 | Install dependencies
pnpm install

# 验证安装 | Verify installation
pnpm validate:all
```

### 2. 日常开发 | Daily Development

```bash
# 启动你要开发的插件 | Start the add-in you want to develop
pnpm dev:excel   # 或 word/ppt | or word/ppt

# 在另一个终端中启动调试 | Start debugging in another terminal
pnpm start:excel # 或 word/ppt | or word/ppt
```

### 3. 添加新依赖 | Adding New Dependencies

```bash
# 为特定插件添加依赖 | Add dependency to specific add-in
pnpm --filter excel-editor4ai add <package-name>

# 为所有插件添加相同依赖 | Add same dependency to all add-ins
pnpm -r add <package-name>

# 添加开发依赖 | Add dev dependency
pnpm --filter excel-editor4ai add -D <package-name>
```

### 4. 代码提交前 | Before Committing

```bash
# 运行代码检查 | Run linting
pnpm lint

# 自动修复问题 | Auto-fix issues
pnpm lint:fix

# 构建测试 | Build test
pnpm build
```

## 🐛 常见问题 | Troubleshooting

### 问题 1: 端口被占用 | Port Already in Use

**症状** | **Symptom**: `Error: listen EADDRINUSE: address already in use :::3001`

**解决方案** | **Solution**:
```bash
# macOS/Linux
lsof -ti:3001 | xargs kill -9

# Windows
netstat -ano | findstr :3001
taskkill /PID <PID> /F
```

### 问题 2: 依赖安装失败 | Dependency Installation Failed

**解决方案** | **Solution**:
```bash
# 清理缓存 | Clear cache
pnpm store prune

# 删除所有 node_modules | Remove all node_modules
pnpm clean

# 重新安装 | Reinstall
pnpm install
```

### 问题 3: Office 无法加载插件 | Office Can't Load Add-in

**症状** | **Symptom**: Office 显示"加载项错误"或插件无法加载

**检查清单** | **Checklist**:
1. ✅ 开发服务器是否正在运行？ | Is the dev server running?
2. ✅ 证书是否已信任？ | Is the certificate trusted?
3. ✅ manifest.xml 中的 URL 是否正确？ | Is the URL in manifest.xml correct?
4. ✅ 端口号是否匹配？ | Does the port number match?
5. ✅ 端口配置在三个地方是否一致？ | Is port configuration consistent in three places?
   - `package.json` → `config.dev_server_port`
   - `manifest.xml` → 所有 `localhost` URL
   - `webpack.config.js` → `urlDev` 变量

**解决方案** | **Solution**:
```bash
# 1. 清理 Office 缓存 | Clear Office cache
pnpm clear-cache

# 2. 关闭 Office 应用 | Close Office application
# 手动关闭或使用命令 | Manually or use command:
killall "Microsoft PowerPoint"  # 或 Excel/Word

# 3. 重新验证清单 | Re-validate manifest
pnpm validate:ppt

# 4. 重启开发服务器 | Restart dev server
pnpm dev:ppt

# 5. 在新终端中启动插件 | Start add-in in new terminal
pnpm start:ppt

# 如果仍然失败，尝试直接在子目录运行 | If still failing, try running directly in subdirectory
cd ppt-editor4ai && pnpm start
```

### 问题 4: TypeScript 编译错误 | TypeScript Compilation Error

**解决方案** | **Solution**:
```bash
# 清理 TypeScript 缓存 | Clear TypeScript cache
rm -rf */node_modules/.cache
rm -rf */*.tsbuildinfo

# 重新构建 | Rebuild
pnpm build:dev
```

## 📚 项目约定 | Project Conventions

### 代码风格 | Code Style

- 使用 TypeScript 严格模式 | Use TypeScript strict mode
- 遵循 ESLint 规则 | Follow ESLint rules
- 使用 Prettier 格式化代码 | Use Prettier for code formatting

### 提交信息 | Commit Messages

```
<type>(<scope>): <subject>

type: feat, fix, docs, style, refactor, test, chore
scope: excel, word, ppt, shared, root
```

示例 | Examples:
- `feat(excel): add new chart feature`
- `fix(word): resolve text formatting issue`
- `docs(root): update README`

### 分支策略 | Branch Strategy

- `main`: 生产分支 | Production branch
- `develop`: 开发分支 | Development branch
- `feature/*`: 功能分支 | Feature branches
- `fix/*`: 修复分支 | Fix branches

## 🔍 调试技巧 | Debugging Tips

### 1. 浏览器开发者工具 | Browser DevTools

Office AddIn 运行在嵌入式浏览器中，可以使用开发者工具调试：

- **Windows**: F12 或右键 → 检查
- **macOS**: 需要使用 Safari 开发者工具连接

Office AddIn runs in an embedded browser, use DevTools for debugging:

- **Windows**: F12 or right-click → Inspect
- **macOS**: Need to use Safari Developer Tools to connect

### 2. 日志调试 | Console Logging

```typescript
// 在代码中添加日志 | Add logging in code
console.log('Debug info:', data);

// 使用 Office.context.ui.displayDialogAsync 显示错误
Office.context.ui.displayDialogAsync(
  'https://localhost:3001/error.html',
  { height: 30, width: 20 }
);
```

### 3. 网络请求调试 | Network Debugging

在 manifest.xml 中确保允许外部请求：

```xml
<AppDomains>
  <AppDomain>https://your-api-domain.com</AppDomain>
</AppDomains>
```

## 📦 构建和部署 | Build and Deployment

### 生产构建 | Production Build

```bash
# 构建所有插件 | Build all add-ins
pnpm build

# 构建产物位置 | Build output location
# excel-editor4ai/dist/
# word-editor4ai/dist/
# ppt-editor4ai/dist/
```

### 部署清单 | Deployment Checklist

1. ✅ 更新 manifest.xml 中的生产 URL
2. ✅ 运行生产构建
3. ✅ 验证所有清单文件
4. ✅ 测试所有功能
5. ✅ 上传到 Office 应用商店或企业目录

## 🎓 学习资源 | Learning Resources

- [Office Add-ins 官方文档](https://docs.microsoft.com/office/dev/add-ins/)
- [pnpm 官方文档](https://pnpm.io/)
- [React 官方文档](https://react.dev/)
- [Fluent UI 文档](https://react.fluentui.dev/)

## 💡 最佳实践 | Best Practices

1. **定期更新依赖** | **Regular Dependency Updates**
   ```bash
   pnpm update -r --latest
   ```

2. **使用 TypeScript 类型** | **Use TypeScript Types**
   ```typescript
   // 使用 Office.js 类型定义 | Use Office.js type definitions
   async function insertText(text: string): Promise<void> {
     await Word.run(async (context) => {
       // ...
     });
   }
   ```

3. **错误处理** | **Error Handling**
   ```typescript
   try {
     await Office.onReady();
   } catch (error) {
     console.error('Office initialization failed:', error);
   }
   ```

4. **性能优化** | **Performance Optimization**
   - 使用 React.memo 避免不必要的重渲染
   - 使用 Office.js 批处理 API
   - 懒加载大型组件

---

**提示** | **Tip**: 将此文档添加到书签，开发时随时查阅！  
Bookmark this document for quick reference during development!
