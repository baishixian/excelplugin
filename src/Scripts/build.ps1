param(
    [string]$Version = "1.0.0",
    [string]$Configuration = "Debug"
)

Write-Host "Building BatchCommentAddin v$Version ($Configuration)" -ForegroundColor Green

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
Get-ChildItem $distPath -Filter "*.xlam" | Remove-Item -Force

Write-Host "Compiling VBA project..." -ForegroundColor Yellow

try {
    # 启动Excel应用程序
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # 创建新工作簿
    $workbook = $excel.Workbooks.Add()
    
    # 导入VBA模块
    $vbProject = $workbook.VBProject
    
    # 导入模块文件
    $moduleFiles = Get-ChildItem (Join-Path $vbaPath "Modules") -Filter "*.bas"
    foreach ($moduleFile in $moduleFiles) {
        Write-Host "Importing module: $($moduleFile.Name)" -ForegroundColor Cyan
        $vbProject.VBComponents.Import($moduleFile.FullName)
    }
    
    # 导入窗体文件
    $formFiles = Get-ChildItem (Join-Path $vbaPath "Forms") -Filter "*.frm"
    foreach ($formFile in $formFiles) {
        Write-Host "Importing form: $($formFile.Name)" -ForegroundColor Cyan
        $vbProject.VBComponents.Import($formFile.FullName)
    }
    
    # 导入ThisWorkbook类
    $thisWorkbookFile = Join-Path $vbaPath "ThisWorkbook.cls"
    if (Test-Path $thisWorkbookFile) {
        Write-Host "Importing ThisWorkbook class..." -ForegroundColor Cyan
        $vbProject.VBComponents.Import($thisWorkbookFile)
    }
    
    # 更新版本信息
    $versionModule = $vbProject.VBComponents("BatchCommentAddin").CodeModule
    $lineCount = $versionModule.CountOfLines
    for ($i = 1; $i -le $lineCount; $i++) {
        $line = $versionModule.Lines($i, 1)
        if ($line -match 'ADDIN_VERSION.*=.*".*"') {
            $newLine = $line -replace '".*"', "`"$Version`""
            $versionModule.ReplaceLine($i, $newLine)
            break
        }
    }
    
    # 保存为加载项
    $addinPath = Join-Path $distPath "BatchCommentAddin.xlam"
    $workbook.SaveAs($addinPath, 55) # xlAddIn format
    
    Write-Host "Build completed successfully!" -ForegroundColor Green
    Write-Host "Output: $addinPath" -ForegroundColor Green
    
} catch {
    Write-Error "Build failed: $($_.Exception.Message)"
    exit 1
} finally {
    # 清理
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { 
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

Write-Host "Build process completed." -ForegroundColor Green