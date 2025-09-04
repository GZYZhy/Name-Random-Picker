# Name Random Picker 🎲

一个基于Python的智能随机抽签工具，支持姓名/分组抽取、特效展示和请假管理

使用说明见[USAGE.md](https://github.com/GZYZhy/Name-Random-Picker/blob/main/USAGE.md)

## ✨ 主要功能

### 🎯 核心功能
- **双模式抽取**：支持个人姓名和小组分组两种抽取模式
- **智能记忆**：自动记录已抽取项，避免重复直到重置
- **请假管理**：提供可视化请假名单编辑界面，支持实时更新
- **配置可视化编辑**：提供配置文件编辑器，支持通过Excel导入/导出

### 🎨 特色特效
- **多媒体彩蛋**：
  - 📸 支持展示自定义图片（PNG/JPG）
  - 🎧 播放专属音效/背景音乐（MP3/WAV）
  - 🎨 自定义文字颜色（支持7种主题色）
- **语音播报**：
  - 🔊 内置TTS语音合成（跨平台支持）
  - ⚡ 离线语音播报功能

### 🖥 窗口控制
- 永久置顶显示
- 透明化设计
- 支持拖拽移动
- 自动窗口大小适配

## 🚀 快速开始

1. 前往[Release](https://github.com/GZYZhy/Name-Random-Picker/releases)页面下载安装包（.exe）

2. 直接安装程序，并按实际情况编辑配置文件。

## ⚙️ 配置文件

V4.0.0新特性：现在可通过右键程序主窗口-编辑配置实现可视化编辑配置文件啦！

对于想亲自编辑源配置文件的极客用户：

通过编辑`config.json`可配置：
- 姓名/分组列表
- 彩蛋特效设置
- 语音参数调整
- 界面显示设置

示例配置文件详见[config.json](https://github.com/GZYZhy/Name-Random-Picker/blob/main/config.json)。您可以结合试运行的表现来很容易地理解各个字段的含义。

在程序开始运行时，会自动检测配置文件的正确性。不合法的字段会以弹窗报错来指导您修改。

尽管我们兼容UTF-8、ANSI等多种编码格式，为了兼容性考虑，我们仍然建议您使用UTF-8编码来保存配置文件。

（示例配置文件可通过关于窗口生成）

## 🤝 参与贡献
欢迎提交Issue和PR！请遵循：
1. Fork本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 发起Pull Request

## 📄 许可证
- 项目免费开源，不存在任何盈利现象。
- 项目中可能引用了部分网络公开图片作为示例图，如有侵权请提交Issue来要求删除。
- 本项目采用 [Apache License 2.0](LICENSE) 授权
- 由 @GZYzhy 发表，WuSiyu（not on GitHub）参与开发
- Copyright (c)2025 GZYzhy.
