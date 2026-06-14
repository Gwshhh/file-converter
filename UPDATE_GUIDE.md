# 🚀 一键更新发布指南

修改代码后，使用此脚本快速发布新版本到 GitHub。

---

## 📝 使用方法

### 方法一：双击运行（推荐）

直接双击 **`update-release.bat`** 文件即可！

### 方法二：命令行

```bash
# Windows
update-release.bat

# Linux/Mac
bash update-release.sh
```

---

## 🔄 脚本会自动完成

1. ✅ **检查并提交代码**
   - 检测未提交的修改
   - 提示输入提交信息
   - 自动 git add + commit

2. ✅ **推送到 GitHub**
   - 自动 git push

3. ✅ **打包程序**
   - 清理旧文件
   - 重新打包 exe
   - 显示文件大小

4. ✅ **发布新版本**
   - 提示输入版本号（如 1.0.2）
   - 提示输入更新说明
   - 自动创建 GitHub Release
   - 自动上传 exe 文件

---

## 📋 使用示例

### 场景：修复了一个 Bug

1. 修改代码
2. 双击 `update-release.bat`
3. 按提示操作：

```
Enter commit message: fix: 修复文件冲突处理的Bug
[自动推送到GitHub]
[自动打包...]
Enter new version: 1.0.2
Enter release notes: 修复文件冲突处理Bug
[自动发布到GitHub]
✅ 完成！
```

---

## ⚙️ 高级用法

### 自定义发布说明

脚本会提示输入版本说明，你可以输入详细的更新内容：

```
Enter release notes: - 修复Bug A
- 优化性能B
- 新增功能C
```

### 查看发布历史

```bash
gh release list
```

### 删除发布（如果发错了）

```bash
gh release delete v1.0.2 --yes
```

---

## ⚠️ 注意事项

### 版本号规则

使用语义化版本号：

- **主版本号**：重大变更（1.0.0 → 2.0.0）
- **次版本号**：新功能（1.0.0 → 1.1.0）
- **修订号**：Bug修复（1.0.0 → 1.0.1）

示例：
- `1.0.2` - Bug修复
- `1.1.0` - 添加新功能
- `2.0.0` - 重大更新

### 更新 updater.py 中的版本号

如果发布新版本，记得修改 `updater.py`：

```python
CURRENT_VERSION = "1.0.2"  # 改为新版本号
```

然后再次运行脚本发布。

---

## 🎯 完整工作流程

### 日常更新流程

```
修改代码
    ↓
运行 update-release.bat
    ↓
输入提交信息
    ↓
输入版本号
    ↓
输入更新说明
    ↓
✅ 自动完成发布
    ↓
用户可以下载新版本
```

### 发布大版本流程

1. **修改版本号**
   - 编辑 `updater.py` 的 `CURRENT_VERSION`

2. **运行脚本**
   - 双击 `update-release.bat`

3. **宣传推广**
   - 分享下载链接
   - 通知用户更新

---

## 🔧 故障排除

### 问题1：提示找不到 Python

**解决**：确保 Python 已安装并在 PATH 中

```bash
python --version
```

### 问题2：提示找不到 gh

**解决**：确保 GitHub CLI 已安装并登录

```bash
gh --version
gh auth status
```

### 问题3：打包失败

**解决**：检查依赖是否完整

```bash
pip install -r requirements.txt
pip install pyinstaller
```

### 问题4：Release 创建失败

**原因**：版本号已存在

**解决**：先删除旧版本或使用新版本号

```bash
gh release delete v1.0.2 --yes
```

---

## 📚 相关文件

- `update-release.bat` - Windows 一键脚本
- `update-release.sh` - Linux/Mac 脚本
- `文件转换器.spec` - PyInstaller 配置
- `updater.py` - 自动更新模块（记得改版本号）

---

## ✅ 检查清单

发布前确认：

- [ ] 代码已测试无误
- [ ] 更新了 `updater.py` 的版本号
- [ ] 版本号符合规范（x.y.z）
- [ ] 更新说明清晰准确

---

**就是这么简单！修改代码 → 双击脚本 → 自动发布！** 🚀
