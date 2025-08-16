param(
    [string]$AddinPath = "",
    [switch]$Uninstall = $false
)

Write-Host "Excel批量批注插件安装程序" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green

# 获取Excel加载项目录
$excelAddinsPath = Join-Path $env:APPDATA "Microsoft\AddIns"
$addinFileName = "BatchCommentAddin.xlam"
$targetPath = Join-Path $excelAddinsPath $addinFileName

# 创建加载项目录（如果不存在）
if (!(Test-Path $excelAddinsPath)) {
    New-Item -ItemType Directory -Path $excelAddinsPath -Force | Out-Null
    Write-Host "创建加载项目录: $excelAddinsPath" -ForegroundColor Yellow
}

if ($Uninstall) {
    # 卸载插件
    Write-Host "正在卸载插件..." -ForegroundColor Yellow
    
    if (Test-Path $targetPath) {
        try {
            Remove-Item $targetPath -Force
            Write-Host "插件卸载成功！" -ForegroundColor Green
            Write-Host "文件已从以下位置删除: $targetPath" -ForegroundColor Cyan
        } catch {
            Write-Error "卸载失败: $($_.Exception.Message)"
            exit 1
        }
    } else {
        Write-Host "插件未安装或已被删除。" -ForegroundColor Yellow
    }
    
    Write-Host "`n请重启Excel以完成卸载。" -ForegroundColor Cyan
    exit 0
}

# 安装插件
Write-Host "正在安装插件..." -ForegroundColor Yellow

# 确定源文件路径
if ($AddinPath -eq "") {
    # 在当前目录查找
    $currentDir = Get-Location
    $possiblePaths = @(
        (Join-Path $currentDir $addinFileName),
        (Join-Path $currentDir "dist\$addinFileName"),
        (Join-Path $currentDir "..\dist\$addinFileName")
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $AddinPath = $path
            break
        }
    }
    
    if ($AddinPath -eq "") {
        Write-Error "找不到插件文件 $addinFileName"
        Write-Host "请指定插件文件路径，例如: .\install.ps1 -AddinPath 'C:\path\to\BatchCommentAddin.xlam'" -ForegroundColor Yellow
        exit 1
    }
}

# 验证源文件存在
if (!(Test-Path $AddinPath)) {
    Write-Error "指定的插件文件不存在: $AddinPath"
    exit 1
}

# 检查Excel是否正在运行
$excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Warning "检测到Excel正在运行，建议先关闭Excel再进行安装。"
    $response = Read-Host "是否继续安装？(y/N)"
    if ($response -ne "y" -and $response -ne "Y") {
        Write-Host "安装已取消。" -ForegroundColor Yellow
        exit 0
    }
}

# 备份现有文件（如果存在）
if (Test-Path $targetPath) {
    $backupPath = $targetPath + ".backup." + (Get-Date -Format "yyyyMMdd_HHmmss")
    try {
        Copy-Item $targetPath $backupPath
        Write-Host "已备份现有插件到: $backupPath" -ForegroundColor Cyan
    } catch {
        Write-Warning "无法创建备份: $($_.Exception.Message)"
    }
}

# 复制插件文件
try {
    Copy-Item $AddinPath $targetPath -Force
    Write-Host "插件安装成功！" -ForegroundColor Green
    Write-Host "安装位置: $targetPath" -ForegroundColor Cyan
} catch {
    Write-Error "安装失败: $($_.Exception.Message)"
    exit 1
}

# 提供启用插件的说明
Write-Host "`n安装完成！请按以下步骤启用插件:" -ForegroundColor Green
Write-Host "1. 启动Microsoft Excel" -ForegroundColor White
Write-Host "2. 点击 文件 > 选项 > 加载项" -ForegroundColor White
Write-Host "3. 在底部的管理下拉框中选择 'Excel加载项'，然后点击 '转到'" -ForegroundColor White
Write-Host "4. 在加载项列表中找到 '批量批注助手'，勾选启用" -ForegroundColor White
Write-Host "5. 点击确定" -ForegroundColor White
Write-Host "`n启用后，您可以在快速访问工具栏或开发工具菜单中找到批量批注功能。" -ForegroundColor Cyan

# 询问是否立即启动Excel
$response = Read-Host "`n是否立即启动Excel？(y/N)"
if ($response -eq "y" -or $response -eq "Y") {
    try {
        Start-Process "excel.exe"
        Write-Host "正在启动Excel..." -ForegroundColor Green
    } catch {
        Write-Warning "无法启动Excel: $($_.Exception.Message)"
    }
}

Write-Host "`n感谢使用批量批注插件！" -ForegroundColor Green