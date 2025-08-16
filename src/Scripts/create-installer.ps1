param(
    [string]$Version = "1.0.0"
)

Write-Host "Creating installer for BatchCommentAddin v$Version" -ForegroundColor Green

$rootPath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$distPath = Join-Path $rootPath "dist"

# 检查必要文件是否存在
$addinFile = Join-Path $distPath "BatchCommentAddin.xlam"
if (!(Test-Path $addinFile)) {
    Write-Warning "BatchCommentAddin.xlam not found in $distPath"
    Write-Host "Available files in dist:" -ForegroundColor Yellow
    Get-ChildItem $distPath -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  - $($_.Name)" }

    # 尝试查找带版本号的文件
    $versionedAddin = Join-Path $distPath "BatchCommentAddin-$Version.xlam"
    if (Test-Path $versionedAddin) {
        Write-Host "Using versioned addin file: $versionedAddin" -ForegroundColor Cyan
        Copy-Item $versionedAddin $addinFile -Force
    } else {
        Write-Error "No addin file found. Cannot create installer."
        exit 1
    }
}

try {
    # 检查是否安装了WiX Toolset
    $wixPaths = @(
        "${env:ProgramFiles(x86)}\WiX Toolset v3.11\bin",
        "${env:ProgramFiles}\WiX Toolset v3.11\bin",
        "${env:ProgramFiles(x86)}\WiX Toolset v4.0\bin",
        "${env:ProgramFiles}\WiX Toolset v4.0\bin"
    )

    $wixPath = $null
    foreach ($path in $wixPaths) {
        if (Test-Path $path) {
            $wixPath = $path
            Write-Host "Found WiX Toolset at: $wixPath" -ForegroundColor Green
            break
        }
    }

    if (!$wixPath) {
        Write-Warning "WiX Toolset not found in standard locations."

        # 在CI环境中尝试安装WiX
        if ($env:CI -eq "true" -or $env:GITHUB_ACTIONS -eq "true") {
            Write-Host "CI environment detected. Attempting to install WiX Toolset..." -ForegroundColor Yellow
            try {
                # 尝试使用Chocolatey安装
                if (Get-Command choco -ErrorAction SilentlyContinue) {
                    choco install wixtoolset -y --no-progress
                    Start-Sleep -Seconds 5

                    # 重新检查WiX路径
                    foreach ($path in $wixPaths) {
                        if (Test-Path $path) {
                            $wixPath = $path
                            Write-Host "WiX Toolset installed successfully at: $wixPath" -ForegroundColor Green
                            break
                        }
                    }
                } else {
                    Write-Warning "Chocolatey not available for WiX installation"
                }
            } catch {
                Write-Warning "Failed to install WiX Toolset: $($_.Exception.Message)"
            }
        }

        if (!$wixPath) {
            Write-Warning "WiX Toolset not available. Creating alternative installer package..."

            # 创建简单的安装包作为备用方案
            $packagePath = Join-Path $distPath "package"
            if (Test-Path $packagePath) {
                Remove-Item $packagePath -Recurse -Force
            }
            New-Item -ItemType Directory -Path $packagePath -Force

            # 复制主文件
            Copy-Item $addinFile $packagePath -Force

            # 复制文档
            $docsToInclude = @("README.md", "LICENSE", "CHANGELOG.md")
            foreach ($doc in $docsToInclude) {
                $docPath = Join-Path $rootPath $doc
                if (Test-Path $docPath) {
                    Copy-Item $docPath $packagePath -Force
                }
            }

            # 复制安装说明
            $installInstructionsPath = Join-Path $distPath "INSTALL_INSTRUCTIONS.txt"
            if (Test-Path $installInstructionsPath) {
                Copy-Item $installInstructionsPath $packagePath -Force
            }

            # 创建增强的安装脚本
            $installScript = @"
@echo off
setlocal enabledelayedexpansion

echo ========================================
echo BatchCommentAddin v$Version Installer
echo ========================================
echo.

REM 检查Excel是否安装
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" >nul 2>&1
if errorlevel 1 (
    reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office" >nul 2>&1
    if errorlevel 1 (
        echo WARNING: Microsoft Office may not be installed.
        echo Please ensure Excel is installed before proceeding.
        echo.
        pause
    )
)

set "ADDIN_PATH=%APPDATA%\Microsoft\AddIns"
echo Installing to: !ADDIN_PATH!

REM 创建加载项目录
if not exist "!ADDIN_PATH!" (
    echo Creating AddIns directory...
    mkdir "!ADDIN_PATH!"
)

REM 备份现有文件
if exist "!ADDIN_PATH!\BatchCommentAddin.xlam" (
    echo Backing up existing installation...
    copy "!ADDIN_PATH!\BatchCommentAddin.xlam" "!ADDIN_PATH!\BatchCommentAddin.xlam.backup" >nul
)

REM 复制新文件
echo Copying BatchCommentAddin.xlam...
copy "BatchCommentAddin.xlam" "!ADDIN_PATH!\" /Y >nul
if !errorlevel! equ 0 (
    echo.
    echo ========================================
    echo Installation completed successfully!
    echo ========================================
    echo.
    echo Next steps:
    echo 1. Start Microsoft Excel
    echo 2. Go to File ^> Options ^> Add-ins
    echo 3. Select "Excel Add-ins" and click "Go..."
    echo 4. Check "BatchCommentAddin" and click OK
    echo.
    echo For detailed instructions, see INSTALL_INSTRUCTIONS.txt
    echo.
) else (
    echo.
    echo ========================================
    echo Installation failed!
    echo ========================================
    echo Please check permissions and try running as administrator.
    echo.
)

pause
"@

            $installScript | Out-File (Join-Path $packagePath "install.bat") -Encoding ASCII

            # 创建卸载脚本
            $uninstallScript = @"
@echo off
setlocal enabledelayedexpansion

echo ========================================
echo BatchCommentAddin v$Version Uninstaller
echo ========================================
echo.

set "ADDIN_PATH=%APPDATA%\Microsoft\AddIns"
set "ADDIN_FILE=!ADDIN_PATH!\BatchCommentAddin.xlam"

if exist "!ADDIN_FILE!" (
    echo Removing BatchCommentAddin.xlam...
    del "!ADDIN_FILE!" >nul 2>&1
    if !errorlevel! equ 0 (
        echo Uninstallation completed successfully!
        echo Please restart Excel to complete the removal.
    ) else (
        echo Failed to remove addin file. Please remove manually:
        echo !ADDIN_FILE!
    )
) else (
    echo BatchCommentAddin.xlam not found in AddIns directory.
    echo It may have been already removed.
)

echo.
pause
"@

            $uninstallScript | Out-File (Join-Path $packagePath "uninstall.bat") -Encoding ASCII

            # 创建ZIP包
            $zipPath = Join-Path $distPath "BatchCommentAddin-Setup-$Version.zip"
            if (Test-Path $zipPath) {
                Remove-Item $zipPath -Force
            }

            if (Get-Command Compress-Archive -ErrorAction SilentlyContinue) {
                Compress-Archive -Path "$packagePath\*" -DestinationPath $zipPath -Force
                Write-Host "Alternative installer package created: $zipPath" -ForegroundColor Green
            } else {
                # 使用.NET方法创建ZIP
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                [System.IO.Compression.ZipFile]::CreateFromDirectory($packagePath, $zipPath)
                Write-Host "Alternative installer package created: $zipPath" -ForegroundColor Green
            }

            # 清理临时目录
            Remove-Item $packagePath -Recurse -Force

            Write-Host "Alternative installer creation completed." -ForegroundColor Green
            return
        }
    }

    # 使用WiX创建MSI安装程序
    Write-Host "Creating MSI installer using WiX Toolset..." -ForegroundColor Yellow

    # 生成GUID
    $productGuid = [System.Guid]::NewGuid().ToString().ToUpper()
    $componentGuid = [System.Guid]::NewGuid().ToString().ToUpper()

    # 创建WiX源文件
    $wixSource = @"
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="$productGuid"
           Name="BatchCommentAddin"
           Language="1033"
           Version="$Version.0"
           Manufacturer="BatchCommentAddin Team"
           UpgradeCode="12345678-1234-1234-1234-123456789012">

    <Package InstallerVersion="200"
             Compressed="yes"
             InstallScope="perUser"
             Description="Excel Batch Comment Add-in v$Version"
             Comments="Professional Excel batch comment tool" />

    <MajorUpgrade DowngradeErrorMessage="A newer version is already installed." />
    <MediaTemplate EmbedCab="yes" />

    <Feature Id="ProductFeature" Title="BatchCommentAddin" Level="1">
      <ComponentGroupRef Id="ProductComponents" />
    </Feature>

    <Property Id="WIXUI_INSTALLDIR" Value="INSTALLFOLDER" />
    <UIRef Id="WixUI_InstallDir" />

  </Product>

  <Fragment>
    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="AppDataFolder">
        <Directory Id="MicrosoftFolder" Name="Microsoft">
          <Directory Id="INSTALLFOLDER" Name="AddIns" />
        </Directory>
      </Directory>
    </Directory>
  </Fragment>

  <Fragment>
    <ComponentGroup Id="ProductComponents" Directory="INSTALLFOLDER">
      <Component Id="MainAddin" Guid="$componentGuid">
        <File Id="BatchCommentAddin.xlam"
              Source="$addinFile"
              KeyPath="yes" />
      </Component>
    </ComponentGroup>
  </Fragment>
</Wix>
"@

    $wixFile = Join-Path $distPath "installer.wxs"
    $wixSource | Out-File $wixFile -Encoding UTF8

    # 编译安装程序
    $candleExe = Join-Path $wixPath "candle.exe"
    $lightExe = Join-Path $wixPath "light.exe"

    if (!(Test-Path $candleExe) -or !(Test-Path $lightExe)) {
        throw "WiX executables not found at $wixPath"
    }

    $wixObjFile = Join-Path $distPath "installer.wixobj"
    $msiFile = Join-Path $distPath "BatchCommentAddin-Setup-$Version.msi"

    # 运行candle
    Write-Host "Running candle.exe..." -ForegroundColor Cyan
    & $candleExe $wixFile -out $wixObjFile
    if ($LASTEXITCODE -ne 0) {
        throw "Candle compilation failed with exit code $LASTEXITCODE"
    }

    # 运行light
    Write-Host "Running light.exe..." -ForegroundColor Cyan
    & $lightExe $wixObjFile -out $msiFile -ext WixUIExtension
    if ($LASTEXITCODE -ne 0) {
        throw "Light linking failed with exit code $LASTEXITCODE"
    }

    # 清理临时文件
    Remove-Item $wixFile -Force -ErrorAction SilentlyContinue
    Remove-Item $wixObjFile -Force -ErrorAction SilentlyContinue

    Write-Host "MSI installer created successfully: $msiFile" -ForegroundColor Green

} catch {
    Write-Error "Installer creation failed: $($_.Exception.Message)"
    Write-Host "Stack trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red

    Write-Host "Attempting to create fallback installer..." -ForegroundColor Yellow

    # 创建简单的EXE包装器作为最后的备用方案
    $exePath = Join-Path $distPath "BatchCommentAddin-Setup-$Version.exe"
    $batContent = @"
@echo off
echo Installing BatchCommentAddin v$Version...
copy "%~dp0BatchCommentAddin.xlam" "%APPDATA%\Microsoft\AddIns\" /Y
if %errorlevel% equ 0 (
    echo Installation completed successfully!
    echo Please restart Excel and enable the add-in.
) else (
    echo Installation failed!
)
pause
"@

    # 这里可以使用更复杂的EXE创建逻辑，但为了简化，我们创建一个批处理文件
    $batContent | Out-File (Join-Path $distPath "BatchCommentAddin-Setup-$Version.bat") -Encoding ASCII
    Write-Host "Fallback installer created as .bat file" -ForegroundColor Yellow
}

Write-Host "Installer creation process completed." -ForegroundColor Green