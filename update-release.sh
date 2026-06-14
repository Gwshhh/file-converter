#!/bin/bash
# 一键更新发布脚本
# 用途: 修改代码后，自动打包、提交、发布新版本

echo "=========================================="
echo "  文件转换器 - 一键更新发布"
echo "=========================================="
echo ""

# 检查是否有未提交的修改
if [[ -n $(git status -s) ]]; then
    echo "⚠️  检测到未提交的修改"
    git status -s
    echo ""
    read -p "请输入提交信息: " commit_msg

    if [ -z "$commit_msg" ]; then
        echo "❌ 提交信息不能为空"
        exit 1
    fi

    echo ""
    echo "📝 提交代码..."
    git add .
    git commit -m "$commit_msg"

    if [ $? -ne 0 ]; then
        echo "❌ 提交失败"
        exit 1
    fi
    echo "✅ 代码已提交"
else
    echo "✅ 没有未提交的修改"
fi

echo ""
echo "📤 推送代码到 GitHub..."
git push

if [ $? -ne 0 ]; then
    echo "❌ 推送失败"
    exit 1
fi
echo "✅ 代码已推送"

echo ""
echo "🔨 开始打包程序..."
python -m PyInstaller 文件转换器.spec --clean

if [ $? -ne 0 ]; then
    echo "❌ 打包失败"
    exit 1
fi

if [ ! -f "dist/文件转换器.exe" ]; then
    echo "❌ 未找到打包文件"
    exit 1
fi

echo "✅ 打包完成"
EXE_SIZE=$(ls -lh "dist/文件转换器.exe" | awk '{print $5}')
echo "   文件大小: $EXE_SIZE"

echo ""
read -p "请输入新版本号 (例如: 1.0.2): " new_version

if [ -z "$new_version" ]; then
    echo "❌ 版本号不能为空"
    exit 1
fi

echo ""
echo "📝 更新 updater.py 版本号..."
sed -i "s/CURRENT_VERSION = \".*\"/CURRENT_VERSION = \"$new_version\"/" updater.py
echo "✅ 版本号已更新"

echo ""
echo "📝 提交版本号更新..."
git add updater.py
git commit -m "chore: update version to $new_version"
git push
echo "✅ 版本号已推送"

echo ""
read -p "请输入版本说明 (按回车使用默认): " release_notes

if [ -z "$release_notes" ]; then
    release_notes="版本 $new_version 更新"
fi

echo ""
echo "🚀 发布 v$new_version 到 GitHub..."

gh release create "v$new_version" \
  "dist/文件转换器.exe" \
  --title "v$new_version" \
  --notes "$release_notes"

if [ $? -eq 0 ]; then
    echo ""
    echo "=========================================="
    echo "  ✅ 发布成功！"
    echo "=========================================="
    echo ""
    echo "📌 版本: v$new_version"
    echo "📌 下载地址:"
    echo "   https://github.com/Gwshhh/file-converter/releases/tag/v$new_version"
    echo ""
else
    echo "❌ 发布失败"
    exit 1
fi
