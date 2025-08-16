param(
    [string]$Version = "1.0.0",
    [string]$Configuration = "Debug"
)

Write-Host "Building BatchCommentAddin v$Version ($Configuration) for CI" -ForegroundColor Green

# 设置路径
$rootPath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$srcPath = Join-Path $rootPath "src"
$distPath = Join-Path $rootPath "dist"
$vbaPath = Join-Path $srcPath "VBA"

# 创建输出目录
if (!(Test-Path $distPath)) {
    New-Item -ItemType Directory -Path $distPath -Force
}

# 清理旧文件
Get-ChildItem $distPath -Filter "*.xlam" -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem $distPath -Filter "*.zip" -ErrorAction SilentlyContinue | Remove-Item -Force

Write-Host "Creating VBA project package for CI..." -ForegroundColor Yellow

try {
    # 检查是否有预构建的 .xlam 文件
    $prebuiltAddin = Join-Path $srcPath "BatchCommentAddin.xlam"

    if (Test-Path $prebuiltAddin) {
        Write-Host "Using pre-built addin file..." -ForegroundColor Cyan
        $targetPath = Join-Path $distPath "BatchCommentAddin-$Version.xlam"
        Copy-Item $prebuiltAddin $targetPath -Force

        # 也创建不带版本号的副本
        $defaultPath = Join-Path $distPath "BatchCommentAddin.xlam"
        Copy-Item $prebuiltAddin $defaultPath -Force

        Write-Host "Build completed successfully!" -ForegroundColor Green
        Write-Host "Output: $targetPath" -ForegroundColor Green
    }
    else {
        Write-Host "No pre-built addin found. Creating comprehensive VBA source package..." -ForegroundColor Cyan

        # 验证VBA项目结构
        $modulesPath = Join-Path $vbaPath "Modules"
        $formsPath = Join-Path $vbaPath "Forms"
        $thisWorkbookPath = Join-Path $vbaPath "ThisWorkbook.cls"

        $hasModules = Test-Path $modulesPath
        $hasForms = Test-Path $formsPath
        $hasThisWorkbook = Test-Path $thisWorkbookPath

        Write-Host "VBA Project Structure:" -ForegroundColor Cyan
        Write-Host "  Modules: $(if($hasModules) { 'Found' } else { 'Missing' })" -ForegroundColor $(if($hasModules) { 'Green' } else { 'Yellow' })
        Write-Host "  Forms: $(if($hasForms) { 'Found' } else { 'Missing' })" -ForegroundColor $(if($hasForms) { 'Green' } else { 'Yellow' })
        Write-Host "  ThisWorkbook: $(if($hasThisWorkbook) { 'Found' } else { 'Missing' })" -ForegroundColor $(if($hasThisWorkbook) { 'Green' } else { 'Yellow' })

        # 创建详细的源码包
        $packagePath = Join-Path $distPath "BatchCommentAddin-Source-$Version.zip"

        # 创建临时打包目录
        $tempPackagePath = Join-Path $distPath "temp_package"
        if (Test-Path $tempPackagePath) {
            Remove-Item $tempPackagePath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $tempPackagePath -Force

        # 复制VBA源码
        if (Test-Path $vbaPath) {
            Copy-Item $vbaPath (Join-Path $tempPackagePath "VBA") -Recurse -Force
            Write-Host "VBA source files copied" -ForegroundColor Green
        }

        # 复制脚本文件
        $scriptsPath = Join-Path $srcPath "Scripts"
        if (Test-Path $scriptsPath) {
            Copy-Item $scriptsPath (Join-Path $tempPackagePath "Scripts") -Recurse -Force
            Write-Host "Build scripts copied" -ForegroundColor Green
        }

        # 复制文档
        $docsToInclude = @("README.md", "LICENSE", "CHANGELOG.md")
        foreach ($doc in $docsToInclude) {
            $docPath = Join-Path $rootPath $doc
            if (Test-Path $docPath) {
                Copy-Item $docPath $tempPackagePath -Force
                Write-Host "Copied $doc" -ForegroundColor Green
            }
        }

        # 复制docs目录
        $docsPath = Join-Path $rootPath "docs"
        if (Test-Path $docsPath) {
            Copy-Item $docsPath $tempPackagePath -Recurse -Force
            Write-Host "Documentation copied" -ForegroundColor Green
        }

        # 压缩源码包
        if (Get-Command Compress-Archive -ErrorAction SilentlyContinue) {
            Compress-Archive -Path "$tempPackagePath\*" -DestinationPath $packagePath -Force
            Write-Host "VBA source package created: $packagePath" -ForegroundColor Green
        } else {
            Write-Warning "Compress-Archive not available. Creating manual archive..."
            # 使用.NET方法创建ZIP
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::CreateFromDirectory($tempPackagePath, $packagePath)
            Write-Host "VBA source package created: $packagePath" -ForegroundColor Green
        }

        # 清理临时目录
        Remove-Item $tempPackagePath -Recurse -Force

        # 创建详细的安装说明
        $readmePath = Join-Path $distPath "INSTALL_INSTRUCTIONS.txt"
        $installInstructions = @"
Excel批量批注插件安装说明
==============================

版本：$Version
构建时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
构建配置：$Configuration

由于CI环境限制，无法自动编译Excel插件。请按以下步骤手动安装：

方法1：使用预构建文件（推荐）
========================================
1. 下载 BatchCommentAddin.xlam 文件
2. 将文件复制到 Excel 加载项目录：
   Windows: %APPDATA%\Microsoft\AddIns\
   Mac: ~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins/
3. 在Excel中启用加载项：
   Windows: 文件 > 选项 > 加载项 > Excel加载项 > 转到 > 勾选"批量批注助手"
   Mac: Excel > 首选项 > 加载项 > 勾选"批量批注助手"

方法2：从源码构建
========================================
1. 解压 BatchCommentAddin-Source-$Version.zip
2. 在安装了Excel的Windows环境中运行：
   powershell -ExecutionPolicy Bypass -File "Scripts/build.ps1" -Version "$Version"
3. 按方法1的步骤2-3安装生成的.xlam文件

系统要求
========================================
- Microsoft Excel 2016 或更高版本
- Windows 10 或更高版本（用于构建）
- .NET Framework 4.7.2 或更高版本

VBA项目结构
========================================
- VBA/Modules/: VBA模块文件 (.bas)
- VBA/Forms/: 用户窗体文件 (.frm)
- VBA/ThisWorkbook.cls: 工作簿类模块
- Scripts/: 构建和安装脚本

故障排除
========================================
如果遇到问题，请：
1. 确保Excel版本兼容
2. 检查宏安全设置
3. 查看项目文档：docs/INSTALLATION.md
4. 提交Issue：https://github.com/yourusername/BatchCommentAddin/issues

技术支持
========================================
- 文档：https://github.com/yourusername/BatchCommentAddin/docs
- Issues：https://github.com/yourusername/BatchCommentAddin/issues
- 邮件：support@example.com
"@

        $installInstructions | Out-File $readmePath -Encoding UTF8
        Write-Host "Installation instructions created: $readmePath" -ForegroundColor Green

        # 创建版本信息文件
        $versionInfo = @{
            Version = $Version
            Configuration = $Configuration
            BuildTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            BuildEnvironment = "CI"
            VBAStructure = @{
                HasModules = $hasModules
                HasForms = $hasForms
                HasThisWorkbook = $hasThisWorkbook
            }
        }

        $versionInfoPath = Join-Path $distPath "version-info.json"
        $versionInfo | ConvertTo-Json -Depth 10 | Out-File $versionInfoPath -Encoding UTF8
        Write-Host "Version info created: $versionInfoPath" -ForegroundColor Green

        Write-Host "CI build completed with comprehensive source package!" -ForegroundColor Green
        Write-Host "Files created:" -ForegroundColor Cyan
        Write-Host "  - $packagePath" -ForegroundColor White
        Write-Host "  - $readmePath" -ForegroundColor White
        Write-Host "  - $versionInfoPath" -ForegroundColor White
    }

} catch {
    Write-Error "Build failed: $($_.Exception.Message)"
    Write-Host "Stack trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}

Write-Host "CI build process completed successfully." -ForegroundColor Green