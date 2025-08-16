# Excel 批量批注插件

[![Build Status](https://github.com/yourusername/BatchCommentAddin/workflows/Build/badge.svg)](https://github.com/yourusername/BatchCommentAddin/actions)
[![Release](https://github.com/yourusername/BatchCommentAddin/workflows/Release/badge.svg)](https://github.com/yourusername/BatchCommentAddin/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

一个专业的Excel批量批注插件，支持多种批注来源、格式设置和模板管理。

## 功能特性

- 🚀 **批量处理**: 支持大规模数据的批量批注操作
- 📝 **多种来源**: 支持从单元格、固定文本、外部文件获取批注内容
- 🎨 **格式设置**: 丰富的字体、颜色、大小设置选项
- 📋 **模板管理**: 预设和自定义批注模板
- 📊 **进度显示**: 实时处理进度和取消支持
- 🔧 **错误处理**: 完善的错误处理和恢复机制
- 💾 **历史记录**: 操作历史和撤销功能

## 系统要求

- Microsoft Excel 2016 或更高版本
- Windows 10 或更高版本
- .NET Framework 4.7.2 或更高版本

## 快速开始

### 下载安装

1. 从 [Releases](https://github.com/yourusername/BatchCommentAddin/releases) 页面下载最新版本
2. 运行 `BatchCommentAddin-Setup.exe` 进行安装
3. 启动Excel，在"加载项"选项卡中找到"批量批注"功能

### 手动安装

1. 下载 `BatchCommentAddin.xlam` 文件
2. 将文件复制到Excel加载项目录
3. 在Excel中启用加载项

## 使用说明

详细使用说明请参考 [用户指南](docs/USER_GUIDE.md)

## 开发

### 构建要求

- Visual Studio 2019 或更高版本
- Office Developer Tools
- PowerShell 5.0 或更高版本

### 本地构建

```bash
git clone https://github.com/yourusername/BatchCommentAddin.git
cd BatchCommentAddin
.\src\Scripts\build.ps1
```

## 贡献

欢迎提交Issue和Pull Request！请参考 [贡献指南](CONTRIBUTING.md)

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 更新日志

详见 [CHANGELOG.md](CHANGELOG.md)