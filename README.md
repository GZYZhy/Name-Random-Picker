# Name Random Picker 🎲

一个基于Python的智能随机抽签工具，支持姓名/分组抽取、特效展示和请假管理

## ✨ 主要功能

### 🎯 核心功能
- **双模式抽取**：支持个人姓名和小组分组两种抽取模式
- **智能记忆**：自动记录已抽取项，避免重复直到重置
- **请假管理**：提供可视化请假名单编辑界面，支持实时更新

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
- 对于Linux/MacOS

```bash
git clone https://github.com/gzyzhy/Name-Random-Picker.git
cd Name-Random-Picker
pip install -r requirements.txt
python main.py
```
- 对于Windows

1. 前往Release页面下载预构建二进制包（.exe）

2. 将.exe文件放在电脑的软件目录，建议采用`C:\ProgramFiles\RandomPicker`

3. 右键.exe文件，选择`发送到->桌面快捷方式`来将软件添加到桌面快捷方式

4. 运行程序，选择生成示例配置文件，在选择目录时建议选择.exe文件所在的目录

5. 按需修改配置文件并保存

6. 再次运行程序，您会看到屏幕左下角的蓝色和红色方块。单击蓝色方块来抽取姓名，单击红色方块来抽取小组，抽取后再次单击来隐藏弹出的名字窗口，右键方块（在教学一体机上可能是长按）来进行其他操作。


## ⚙️ 配置文件
通过编辑`config.json`可配置：
- 姓名/分组列表
- 彩蛋特效设置
- 语音参数调整
- 界面显示设置

（示例配置文件可通过关于窗口生成）

## 🤝 参与贡献
欢迎提交Issue和PR！请遵循：
1. Fork本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 发起Pull Request

## 📄 许可证
- 本项目采用 [Apache License 2.0](LICENSE) 授权
- 由 @GZYzhy 发表，WuSiyu（not on GitHub）参与开发
- Copyright (c)2025 GZYzhy.
