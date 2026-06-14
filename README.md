# 📄 文件格式转换器

<div align="center">

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

一款功能强大的文档格式转换工具，支持 Markdown、HTML、Word、PDF、TXT 等多种格式互转。

[功能特点](#-功能特点) • [下载安装](#-下载安装) • [使用说明](#-使用说明) • [构建指南](#-构建指南) • [常见问题](#-常见问题)

</div>

---

## ⚠️ 免责声明

**本软件仅供学习交流使用，请勿用于商业用途。**

- 使用本软件进行文件转换产生的任何后果由使用者自行承担
- 本软件依赖 Microsoft Word 进行 PDF ↔ Word 转换，需要本机安装 Microsoft Office
- 请勿使用本软件处理机密或敏感文档，开发者不对数据安全性负责
- 本软件完全免费开源，如有收费或捆绑行为均非本项目所为

**使用本软件即表示您已阅读并同意上述条款。**

---

## ✨ 功能特点

### 🔄 多格式支持
- **Markdown** (.md, .markdown) - 轻量级标记语言
- **HTML** (.html, .htm) - 网页文档
- **Word** (.docx) - Microsoft Word 文档
- **PDF** (.pdf) - 便携式文档格式
- **TXT** (.txt) - 纯文本文件

### 🎨 现代化界面
- 简洁美观的图形界面
- 支持拖拽文件快速导入
- 实时显示转换进度
- 智能文件冲突处理

### ⚡ 高效转换
- 多线程并行转换
- 智能格式识别
- 保留原始文档样式
- 自动嵌入图片和资源

### 🛡️ 智能功能
- 自动检查软件更新
- 同名文件智能重命名
- 详细的转换结果报告
- 错误提示和异常处理

---

## 📥 下载安装

### 方式一：直接下载（推荐）

1. 前往 [Releases 页面](https://github.com/your-username/file-converter/releases/latest)
2. 下载最新版本的 `文件转换器.exe`
3. 双击运行即可，无需安装

### 方式二：从源码构建

#### 前置要求

- Python 3.8 或更高版本
- pip 包管理器
- Microsoft Office（用于 PDF ↔ Word 转换）

#### 安装依赖

```bash
pip install -r requirements.txt
```

#### 运行程序

```bash
python md_converter.py
```

---

## 📖 使用说明

### 基础使用

1. **选择源文件**
   - 点击「选择文件」按钮
   - 或直接拖拽文件到窗口中

2. **选择输出格式**
   - 勾选需要转换的目标格式（可多选）
   - 灰色按钮表示与源文件格式相同

3. **设置保存位置**
   - 默认保存在源文件所在目录
   - 可点击「更改」按钮自定义保存位置

4. **开始转换**
   - 点击「开始转换」按钮
   - 等待进度条完成
   - 转换完成后可直接打开文件夹

### 文件冲突处理

当目标目录已存在同名文件时，提供四种处理方式：

- **智能重命名**（推荐）：自动添加序号，如 `文档_1.pdf`
- **覆盖替换**：直接覆盖原文件
- **跳过**：不转换该格式
- **取消**：取消整个转换任务

### 支持的转换路径

| 源格式 | 可转换为 |
|--------|----------|
| Markdown | HTML, Word, PDF, TXT |
| HTML | Markdown, Word, PDF, TXT |
| Word | Markdown, HTML, PDF, TXT |
| PDF | Markdown, HTML, Word, TXT |
| TXT | Markdown, HTML, Word, PDF |

### 快捷操作

- **拖拽文件**：直接将文件拖入窗口
- **检查更新**：菜单栏 → 帮助 → 检查更新
- **关于信息**：菜单栏 → 帮助 → 关于

---

## 🔧 构建指南

### 生成独立可执行文件

使用 PyInstaller 打包为单文件 exe：

```bash
# 安装 PyInstaller
pip install pyinstaller

# 打包程序（使用提供的 spec 文件）
pyinstaller 文件转换器.spec
```

生成的可执行文件位于 `dist/` 目录下。

### 自定义打包

如需修改打包配置，编辑 `文件转换器.spec` 文件：

```python
exe = EXE(
    ...
    name='文件转换器',      # 输出文件名
    console=False,           # 不显示控制台
    icon='icon.ico',         # 自定义图标（可选）
)
```

---

## 📦 依赖项

核心依赖库：

```
PySide6 >= 6.5.0        # Qt6 图形界面框架
pypandoc >= 1.11        # Pandoc 文档转换包装器
pywin32 >= 305          # Windows API 调用（Word 操作）
packaging >= 23.0       # 版本号比较
```

系统依赖：

- **Microsoft Office**（Word）：用于高质量 PDF ↔ Word 转换
- **Pandoc**（内置）：核心转换引擎

---

## 🎯 技术架构

### 核心技术

- **界面框架**：PySide6 (Qt6)
- **转换引擎**：Pandoc + Microsoft Word COM API
- **多线程处理**：ThreadPoolExecutor 并发转换
- **自动更新**：基于 GitHub Releases API

### 转换策略

| 转换类型 | 使用工具 | 特点 |
|----------|----------|------|
| Markdown → HTML/Word | Pandoc | 快速，格式还原度高 |
| Word → PDF | Word COM API | 完美保留格式 |
| PDF → Word | Word COM API | 依赖 Office 版本 |
| 其他组合 | Pandoc + 临时文件 | 灵活转换 |

### 项目结构

```
file-converter/
├── md_converter.py          # 主程序
├── updater.py               # 自动更新模块
├── requirements.txt         # Python 依赖
├── 文件转换器.spec          # PyInstaller 配置
├── README.md               # 说明文档
├── LICENSE                 # MIT 许可证
└── .gitignore              # Git 忽略配置
```

---

## ❓ 常见问题

### Q1: PDF 转换失败怎么办？

**A:** PDF 转换依赖 Microsoft Word，请确保：
- 已安装完整版 Microsoft Office（不支持 WPS）
- Office 版本支持 PDF 功能（推荐 Office 2016 及以上）
- 尝试先转为 Word 格式，再转为其他格式

### Q2: 转换后格式丢失或错乱？

**A:** 不同格式之间存在固有差异：
- Markdown → PDF/Word 效果最佳
- PDF → 其他格式 可能丢失复杂排版
- 建议尽量使用源格式为 Markdown 或 Word 的文档

### Q3: 程序启动报错缺少 DLL？

**A:** 需要安装 Visual C++ 运行库：
- 下载 [VC++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)
- 或从源码运行（需要 Python 环境）

### Q4: 能否批量转换多个文件？

**A:** 当前版本仅支持单文件多格式转换，批量功能计划在后续版本添加。

### Q5: 转换速度慢怎么办？

**A:** 影响因素：
- 文件大小和复杂度
- 包含的图片数量
- PDF 转换需要启动 Word，较耗时
- 建议避免转换过大的 PDF 文件（>50MB）

---

## 🗺️ 开发计划

- [ ] 批量文件转换
- [ ] 拖拽多文件支持
- [ ] 转换参数自定义（DPI、页边距等）
- [ ] 支持更多格式（EPUB、RTF、ODT）
- [ ] Linux/macOS 版本支持
- [ ] 命令行接口（CLI）

---

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request！

### 贡献流程

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

### 开发建议

- 遵循 PEP 8 代码规范
- 添加必要的注释和文档
- 测试新功能在 Windows 10/11 上的兼容性

---

## 📄 开源协议

本项目采用 [MIT License](LICENSE) 开源协议。

```
MIT License

Copyright (c) 2024 File Converter Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction...
```

---

## 🙏 致谢

本项目使用了以下优秀的开源项目：

- [Pandoc](https://pandoc.org/) - 通用文档转换器
- [PySide6](https://wiki.qt.io/Qt_for_Python) - Python Qt 绑定
- [PyInstaller](https://pyinstaller.org/) - Python 打包工具

---

## 📧 联系方式

- **问题反馈**：[GitHub Issues](https://github.com/your-username/file-converter/issues)
- **功能建议**：[GitHub Discussions](https://github.com/your-username/file-converter/discussions)

---

<div align="center">

**如果这个项目对你有帮助，请给个 ⭐ Star 支持一下！**

Made with ❤️ by File Converter Contributors

</div>
