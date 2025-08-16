param(
    [string]$Version = "1.0.0"
)

Write-Host "Creating installer for BatchCommentAddin v$Version" -ForegroundColor Green

$rootPath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$distPath = Join-Path $rootPath "dist"

# 检查是否安装了WiX Toolset
$wixPath = "${env:ProgramFiles(x86)}\WiX Toolset v3.11\bin"
if (!(Test-Path $wixPath)) {
    Write-Warning "WiX Toolset not found. Installing via Chocolatey..."
    choco install wixtoolset -y
}

# 创建WiX源文件
$wixSource = @"
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*" 
           Name="BatchCommentAddin" 
           Language="2052" 
           Version="$Version" 
           Manufacturer="OneDay Team" 
           UpgradeCode="12345678-1234-1234-1234-123456789012">
    
    <Package InstallerVersion="200" 
             Compressed="yes" 
             InstallScope="perUser" 
             Description="Excel批量批注插件" 
             Comments="专业的Excel批量批注工具" />

    <MajorUpgrade DowngradeErrorMessage="A newer version is already installed." />
    <MediaTemplate EmbedCab="yes" />

    <Feature Id="ProductFeature" Title="BatchCommentAddin" Level="1">
      <ComponentGroupRef Id="ProductComponents" />
    </Feature>
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
      <Component Id="MainAddin" Guid="*">
        <File Id="BatchCommentAddin.xlam" 
              Source="$distPath\BatchCommentAddin.xlam" 
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

& $candleExe $wixFile -out (Join-Path $distPath "installer.wixobj")
& $lightExe (Join-Path $distPath "installer.wixobj") -out (Join-Path $distPath "BatchCommentAddin-Setup-$Version.msi")

Write-Host "Installer created successfully!" -ForegroundColor Green