# 🚀 GitHub 发布指南

本文档将指导您如何将此项目发布到 GitHub。

## 📋 发布前准备

### 1. 确认文件完整性

已创建的文件：
- ✅ `md_converter.py` - 主程序
- ✅ `updater.py` - 自动更新模块
- ✅ `requirements.txt` - Python 依赖
- ✅ `文件转换器.spec` - PyInstaller 配置
- ✅ `README.md` - 项目说明文档
- ✅ `LICENSE` - MIT 开源协议
- ✅ `.gitignore` - Git 忽略配置
- ✅ Git 仓库已初始化

### 2. 修改配置

**重要**：发布前必须修改 `updater.py` 中的仓库信息：

```python
# 将 "your-username" 改为你的 GitHub 用户名
GITHUB_REPO = "your-username/file-converter"
```

## 🌐 在 GitHub 上创建仓库

### 方法一：通过 GitHub 网页创建

1. **登录 GitHub**
   - 访问 https://github.com
   - 登录你的账号

2. **创建新仓库**
   - 点击右上角 `+` → `New repository`
   - 仓库名称：`file-converter`（或其他名称）
   - 描述：`一款功能强大的文档格式转换工具，支持 Markdown、HTML、Word、PDF、TXT 等多种格式互转`
   - 选择 `Public`（公开）
   - **不要**勾选 "Add a README file"
   - **不要**添加 .gitignore 或 license（我们已经创建了）
   - 点击 `Create repository`

3. **记录仓库地址**
   - 创建后会显示类似这样的地址：
   ```
   https://github.com/your-username/file-converter.git
   ```

### 方法二：使用 GitHub CLI（推荐）

如果已安装 `gh` 命令行工具：

```bash
# 在当前目录创建 GitHub 仓库
gh repo create file-converter --public --source=. --description="文档格式转换工具" --push
```

## 📤 推送代码到 GitHub

### 使用 HTTPS（推荐新手）

```bash
# 添加远程仓库（替换 your-username 为你的 GitHub 用户名）
git remote add origin https://github.com/your-username/file-converter.git

# 推送代码到 GitHub
git branch -M main
git push -u origin main
```

### 使用 SSH（推荐有经验用户）

```bash
# 添加远程仓库
git remote add origin git@github.com:your-username/file-converter.git

# 推送代码
git branch -M main
git push -u origin main
```

## 🏷️ 创建首个 Release

### 1. 打包可执行文件

```bash
# 确保已安装依赖
pip install -r requirements.txt
pip install pyinstaller

# 打包为单文件 exe
pyinstaller 文件转换器.spec
```

打包完成后，`dist/文件转换器.exe` 即为可分发的程序。

### 2. 通过 GitHub 网页创建 Release

1. 进入你的仓库页面
2. 点击右侧 `Releases` → `Create a new release`
3. 填写信息：
   - **Tag version**: `v1.0.0`
   - **Release title**: `v1.0.0 - 首个正式版本`
   - **Description**: 复制下面的模板

```markdown
## 🎉 首个正式版本发布

### ✨ 功能特性

- ✅ 支持 Markdown、HTML、Word、PDF、TXT 多格式互转
- ✅ 现代化图形界面，支持拖拽文件
- ✅ 多线程并行转换，高效快速
- ✅ 智能文件冲突处理（重命名/覆盖/跳过）
- ✅ 自动检查软件更新
- ✅ 详细的转换进度和结果报告

### 📥 下载安装

**Windows 用户**：下载 `文件转换器.exe`，双击即可运行，无需安装。

**系统要求**：
- Windows 10/11 64位
- Microsoft Office（用于 PDF ↔ Word 转换）

### ⚠️ 免责声明

本软件仅供学习交流使用，使用产生的任何后果由使用者自行承担。

### 🐛 已知问题

- PDF 转换依赖本机安装的 Microsoft Word
- 不支持 WPS Office

---

**完整使用说明请查看 [README.md](https://github.com/your-username/file-converter#readme)**
```

4. **上传文件**
   - 点击 "Attach binaries" 或拖拽文件
   - 上传 `dist/文件转换器.exe`

5. **发布**
   - 勾选 `Set as the latest release`
   - 点击 `Publish release`

### 3. 使用 GitHub CLI 创建 Release

```bash
# 创建 release 并上传文件
gh release create v1.0.0 \
  dist/文件转换器.exe \
  --title "v1.0.0 - 首个正式版本" \
  --notes "首个正式版本，支持多格式文档转换"
```

## 🔄 更新 updater.py 配置

**重要步骤**：发布后必须更新仓库信息：

1. 编辑 `updater.py`：

```python
# 修改为你的实际仓库地址
GITHUB_REPO = "your-username/file-converter"  # 替换 your-username
```

2. 提交更改并推送：

```bash
git add updater.py
git commit -m "chore: 更新仓库地址配置"
git push
```

## ✅ 验证发布

### 1. 检查仓库页面
- 访问 `https://github.com/your-username/file-converter`
- 确认 README 显示正常
- 检查文件列表完整

### 2. 测试 Release 下载
- 点击 `Releases` 标签
- 确认可以下载 `文件转换器.exe`
- 下载后测试程序是否正常运行

### 3. 测试自动更新功能
- 运行程序
- 点击 `帮助` → `检查更新`
- 应该显示"已是最新版本"

## 📝 后续更新流程

### 发布新版本步骤：

1. **修改代码**
   ```bash
   # 编辑文件...
   git add .
   git commit -m "feat: 新功能描述"
   ```

2. **更新版本号**
   - 修改 `updater.py` 中的 `CURRENT_VERSION`
   ```python
   CURRENT_VERSION = "1.1.0"  # 更新版本号
   ```

3. **推送代码**
   ```bash
   git push
   ```

4. **重新打包**
   ```bash
   pyinstaller 文件转换器.spec
   ```

5. **创建新 Release**
   ```bash
   gh release create v1.1.0 \
     dist/文件转换器.exe \
     --title "v1.1.0 - 更新说明" \
     --notes "更新内容详情"
   ```

## 🎨 美化 GitHub 仓库（可选）

### 添加徽章（Badges）

README.md 顶部已包含徽章：
- Version（版本号）
- Platform（平台）
- License（许可证）

### 添加 Topics

在仓库页面点击 `⚙️ Settings` → 设置 Topics：
- `file-converter`
- `document-converter`
- `markdown`
- `pdf`
- `python`
- `pyside6`
- `pandoc`

### 设置仓库描述

在仓库首页，点击 `About` 旁边的 `⚙️` 图标，添加：
- **Description**: `文档格式转换工具 - 支持 Markdown、HTML、Word、PDF、TXT 互转`
- **Website**: （可选）如有项目主页
- **Topics**: 如上所述

## 🔗 分享链接

发布完成后，你可以分享以下链接：

- **仓库主页**: `https://github.com/your-username/file-converter`
- **下载地址**: `https://github.com/your-username/file-converter/releases/latest`
- **问题反馈**: `https://github.com/your-username/file-converter/issues`

## ⚠️ 注意事项

1. **不要提交敏感信息**
   - 不要上传包含个人信息的文件
   - 不要提交密码或密钥

2. **不要提交大文件**
   - 不要提交 `build/` 和 `dist/` 目录（已在 .gitignore）
   - exe 文件只通过 Releases 发布，不直接提交

3. **遵守开源协议**
   - 本项目使用 MIT 协议
   - 允许他人自由使用和修改

4. **定期维护**
   - 及时回复 Issues
   - 审查 Pull Requests
   - 保持 README 更新

## 🆘 常见问题

### Q: 推送时要求输入密码？
**A**: GitHub 已停止密码认证，需要使用 Personal Access Token 或 SSH 密钥。

### Q: 文件太大无法上传？
**A**: GitHub 单文件限制 100MB，Release 附件限制 2GB。如果 exe 超过限制，考虑压缩或使用外部下载链接。

### Q: 如何删除 Release？
**A**: 在 Releases 页面，点击对应版本的 `Delete` 按钮。

---

## ✅ 检查清单

发布前确认：

- [ ] 修改了 `updater.py` 中的 `GITHUB_REPO` 配置
- [ ] 所有代码已提交到本地仓库
- [ ] 在 GitHub 创建了仓库
- [ ] 推送代码成功
- [ ] 打包生成了可执行文件
- [ ] 创建了 v1.0.0 Release
- [ ] 上传了 exe 文件
- [ ] 测试下载链接可用
- [ ] 测试程序运行正常

**完成以上步骤后，你的项目就成功发布到 GitHub 了！** 🎉
