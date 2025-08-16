param(
    [string]$Version = "1.0.0"
)

Write-Host "Packaging BatchCommentAddin v$Version" -ForegroundColor Green

$rootPath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$distPath = Join-Path $rootPath "dist"
$resourcesPath = Join-Path $rootPath "src\Resources"

# 创建包目录
$packagePath = Join-Path $distPath "package"
if (Test-Path $packagePath) {
    Remove-Item $packagePath -Recurse -Force
}
New-Item -ItemType Directory -Path $packagePath -Force

# 复制主文件
Copy-Item (Join-Path $distPath "BatchCommentAddin.xlam") $packagePath

# 复制资源文件
if (Test-Path $resourcesPath) {
    Copy-Item $resourcesPath $packagePath -Recurse
}

# 复制文档
Copy-Item (Join-Path $rootPath "README.md") $packagePath
Copy-Item (Join-Path $rootPath "LICENSE") $packagePath
Copy-Item (Join-Path $rootPath "CHANGELOG.md") $packagePath

# 创建安装脚本
$installScript = @"
@echo off
echo Installing BatchCommentAddin v$Version
echo.

set "ADDIN_PATH=%APPDATA%\Microsoft\AddIns"
if not exist "%ADDIN_PATH%" mkdir "%ADDIN_PATH%"

copy "BatchCommentAddin.xlam" "%ADDIN_PATH%\" /Y
if %errorlevel% equ 0 (
    echo Installation completed successfully!
    echo Please restart Excel and enable the add-in.
) else (
    echo Installation failed!
)

pause
"@

$installScript | Out-File (Join-Path $packagePath "install.bat") -Encoding ASCII

# 创建ZIP包
$zipPath = Join-Path $distPath "BatchCommentAddin-$Version.zip"
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($packagePath, $zipPath)

Write-Host "Package created: $zipPath" -ForegroundColor Green

# 清理临时目录
Remove-Item $packagePath -Recurse -Force

Write-Host "Packaging completed." -ForegroundColor Green