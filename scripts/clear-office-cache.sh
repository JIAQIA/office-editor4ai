#!/bin/bash
# 清理 Office AddIn 缓存 | Clear Office AddIn Cache
# 用于解决 Office 加载旧的 manifest 文件的问题 | Fixes issues with Office loading old manifest files

echo "🧹 清理 Office AddIn 缓存 | Clearing Office AddIn cache..."

# PowerPoint 缓存 | PowerPoint cache
PPT_CACHE="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
if [ -d "$PPT_CACHE" ]; then
    echo "  清理 PowerPoint 缓存 | Clearing PowerPoint cache..."
    rm -rf "$PPT_CACHE"/*
    echo "  ✅ PowerPoint 缓存已清理 | PowerPoint cache cleared"
fi

# Word 缓存 | Word cache
WORD_CACHE="$HOME/Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
if [ -d "$WORD_CACHE" ]; then
    echo "  清理 Word 缓存 | Clearing Word cache..."
    rm -rf "$WORD_CACHE"/*
    echo "  ✅ Word 缓存已清理 | Word cache cleared"
fi

# Excel 缓存 | Excel cache
EXCEL_CACHE="$HOME/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
if [ -d "$EXCEL_CACHE" ]; then
    echo "  清理 Excel 缓存 | Clearing Excel cache..."
    rm -rf "$EXCEL_CACHE"/*
    echo "  ✅ Excel 缓存已清理 | Excel cache cleared"
fi

echo ""
echo "✨ 缓存清理完成！| Cache clearing complete!"
echo ""
echo "📝 下一步 | Next steps:"
echo "  1. 关闭所有 Office 应用 | Close all Office applications"
echo "  2. 重新启动开发服务器 | Restart dev server: pnpm dev:ppt"
echo "  3. 重新加载插件 | Reload add-in: pnpm start:ppt"
