# 安装指南

本文档详细介绍了如何安装和配置Excel批量批注插件。

## 系统要求

### 最低要求
- **操作系统**: Windows 10 或更高版本
- **Microsoft Excel**: 2016 或更高版本
- **.NET Framework**: 4.7.2 或更高版本
- **内存**: 至少 4GB RAM
- **磁盘空间**: 至少 50MB 可用空间

### 推荐配置
- **操作系统**: Windows 11
- **Microsoft Excel**: 2021 或 Microsoft 365
- **.NET Framework**: 4.8 或更高版本
- **内存**: 8GB RAM 或更多
- **磁盘空间**: 100MB 可用空间

## 安装方法

### 方法一：使用安装程序（推荐）

1. **下载安装程序**
   - 访问 [GitHub Releases](https://github.com/yourusername/BatchCommentAddin/releases) 页面
   - 下载最新版本的 `BatchCommentAddin-Setup-x.x.x.exe`

2. **运行安装程序**
   - 双击下载的安装程序
   - 按照安装向导的提示完成安装
   - 安装程序会自动将插件复制到正确的位置

3. **启用插件**
   - 启动 Microsoft Excel
   - 点击 **文件** > **选项** > **加载项**
   - 在底部的"管理"下拉框中选择 **Excel 加载项**，然后点击 **转到**
   - 在加载项列表中找到 **批量批注助手**，勾选启用
   - 点击 **确定**

### 方法二：手动安装

1. **下载插件文件**
   - 从 [GitHub Releases](https://github.com/yourusername/BatchCommentAddin/releases) 下载 `BatchCommentAddin.xlam` 文件

2. **复制到加载项目录**
   ```
   %APPDATA%\Microsoft\AddIns\
   ```
   - 按 `Win + R` 打开运行对话框
   - 输入 `%APPDATA%\Microsoft\AddIns\` 并按回车
   - 将 `BatchCommentAddin.xlam` 文件复制到此目录

3. **启用插件**
   - 按照方法一的第3步启用插件

### 方法三：使用PowerShell脚本安装

1. **下载安装脚本**
   - 下载项目中的 `src/Scripts/install.ps1` 文件

2. **运行安装脚本**
   ```powershell
   # 以管理员身份运行PowerShell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   .\install.ps1 -AddinPath "C:\path\to\BatchCommentAddin.xlam"
   ```

## 验证安装

安装完成后，您可以通过以下方式验证插件是否正确安装：

1. **检查快速访问工具栏**
   - 在Excel中查看快速访问工具栏
   - 应该能看到"批量批注"按钮

2. **检查开发工具菜单**
   - 如果启用了开发工具选项卡
   - 在开发工具菜单中应该能找到"批量批注工具"选项

3. **测试功能**
   - 点击"批量批注"按钮
   - 应该能打开批量批注工具对话框

## 故障排除

### 常见问题

#### 1. 插件未出现在加载项列表中

**可能原因**：
- 文件未正确复制到加载项目录
- Excel版本不兼容
- 文件被防病毒软件阻止

**解决方案**：
- 确认文件位置：`%APPDATA%\Microsoft\AddIns\BatchCommentAddin.xlam`
- 检查文件属性，确保没有被阻止
- 右键点击文件 > 属性 > 如果有"解除阻止"按钮，点击它

#### 2. 插件加载失败

**可能原因**：
- 宏安全设置过高
- .NET Framework版本不兼容
- Excel权限不足

**解决方案**：
- 调整宏安全设置：文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置
- 选择"禁用所有宏，并发出通知"或"启用所有宏"
- 确保安装了.NET Framework 4.7.2或更高版本

#### 3. 功能按钮不显示

**可能原因**：
- 插件已加载但界面未更新
- Excel缓存问题

**解决方案**：
- 重启Excel
- 清除Excel缓存：关闭Excel，删除临时文件
- 重新启用插件

#### 4. 权限错误

**可能原因**：
- 用户权限不足
- 企业环境限制

**解决方案**：
- 以管理员身份运行Excel
- 联系IT管理员获取权限
- 使用便携版本（如果可用）

### 日志文件

如果遇到问题，可以查看日志文件获取更多信息：

**日志位置**：
```
%APPDATA%\BatchCommentAddin\logs\
```

**日志文件**：
- `operations_YYYYMMDD.log` - 操作日志
- `errors_YYYYMMDD.log` - 错误日志

### 获取帮助

如果以上方法都无法解决问题，请：

1. **查看文档**
   - 阅读 [用户指南](USER_GUIDE.md)
   - 查看 [常见问题](FAQ.md)

2. **提交问题**
   - 访问 [GitHub Issues](https://github.com/yourusername/BatchCommentAddin/issues)
   - 提供详细的错误信息和系统环境

3. **联系支持**
   - 发送邮件至：support@example.com
   - 包含日志文件和错误截图

## 卸载

### 使用安装程序卸载

1. 打开 **控制面板** > **程序和功能**
2. 找到 **BatchCommentAddin**
3. 点击 **卸载**

### 手动卸载

1. **禁用插件**
   - 在Excel中：文件 > 选项 > 加载项
   - 取消勾选"批量批注助手"

2. **删除文件**
   ```
   %APPDATA%\Microsoft\AddIns\BatchCommentAddin.xlam
   %APPDATA%\BatchCommentAddin\
   ```

3. **清理注册表**（可选）
   - 运行 `regedit`
   - 删除相关注册表项（如果存在）

### 使用PowerShell脚本卸载

```powershell
.\install.ps1 -Uninstall
```

## 更新

### 自动更新

插件支持检查更新功能：
- 在插件中点击"检查更新"
- 如果有新版本，会提示下载

### 手动更新

1. 下载新版本的安装程序
2. 运行安装程序（会自动覆盖旧版本）
3. 重启Excel

## 企业部署

### 批量部署

对于企业环境，可以使用以下方法批量部署：

1. **使用组策略**
   - 将插件文件复制到网络共享位置
   - 通过组策略推送到用户的加载项目录

2. **使用SCCM**
   - 创建SCCM包
   - 部署到目标计算机

3. **使用PowerShell脚本**
   ```powershell
   # 批量安装脚本示例
   $computers = Get-Content "computers.txt"
   foreach ($computer in $computers) {
       Copy-Item "BatchCommentAddin.xlam" "\\$computer\c$\Users\*\AppData\Roaming\Microsoft\AddIns\"
   }
   ```

### 配置管理

企业可以通过以下方式管理插件配置：

1. **集中配置文件**
   - 在网络位置放置配置文件
   - 插件启动时读取集中配置

2. **注册表设置**
   - 通过组策略设置注册表值
   - 控制插件行为和权限

3. **模板管理**
   - 预配置批注模板
   - 限制用户自定义模板

## 许可证

本插件采用MIT许可证，详见 [LICENSE](../LICENSE) 文件。

## 版本历史

查看 [CHANGELOG.md](../CHANGELOG.md) 了解版本更新历史。