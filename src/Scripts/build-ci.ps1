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
        Write-Host "Creating VBA source package..." -ForegroundColor Cyan
        
        # 创建源码包
        $packagePath = Join-Path $distPath "BatchCommentAddin-Source-$Version.zip"
        
        # 压缩 VBA 源码
        if (Get-Command Compress-Archive -ErrorAction SilentlyContinue) {
            Compress-Archive -Path $vbaPath -DestinationPath $packagePath -Force
            Write-Host "VBA source package created: $packagePath" -ForegroundColor Green
        }
        
        # 创建安装说明
        $readmePath = Join-Path $distPath "INSTALL_INSTRUCTIONS.txt"
        $installInstructions = @"
Excel批量批注插件安装说明
==============================

由于CI环境限制，无法自动编译Excel插件。请按以下步骤手动安装：

方法1：使用预构建文件（推荐）
1. 下载 BatchCommentAddin.xlam 文件
2. 将文件复制到 Excel 加载项目录：
   %APPDATA%\Microsoft\AddIns\
3. 在Excel中启用加载项：
   文件 > 选项 > 加载项 > Excel加载项 > 转到 > 勾选"批量批注助手"

方法2：从源码构建
1. 在安装了Excel的Windows环境中运行：
   powershell -ExecutionPolicy Bypass -File "src/Scripts/build.ps1"
2. 按方法1的步骤2-3安装生成的.xlam文件

版本：$Version
构建时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
        
        $installInstructions | Out-File $readmePath -Encoding UTF8
        Write-Host "Installation instructions created: $readmePath" -ForegroundColor Green
        
        Write-Host "CI build completed with source package!" -ForegroundColor Green
    }
    
} catch {
    Write-Error "Build failed: $($_.Exception.Message)"
    exit 1
}

Write-Host "CI build process completed." -ForegroundColor Green