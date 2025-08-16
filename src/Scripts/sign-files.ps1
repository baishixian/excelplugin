param(
    [string]$CertificatePath = $env:SIGN_CERTIFICATE,
    [string]$CertificatePassword = $env:SIGN_PASSWORD,
    [string]$TimestampUrl = "http://timestamp.digicert.com"
)

Write-Host "Signing BatchCommentAddin files..." -ForegroundColor Green

$rootPath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$distPath = Join-Path $rootPath "dist"

# 检查证书文件是否存在
if (!(Test-Path $CertificatePath)) {
    Write-Error "Certificate file not found: $CertificatePath"
    exit 1
}

# 查找signtool.exe
$signToolPaths = @(
    "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe",
    "${env:ProgramFiles}\Windows Kits\10\bin\*\x64\signtool.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\*\bin\signtool.exe"
)

$signTool = $null
foreach ($path in $signToolPaths) {
    $found = Get-ChildItem $path -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) {
        $signTool = $found.FullName
        break
    }
}

if (!$signTool) {
    Write-Error "SignTool.exe not found. Please install Windows SDK."
    exit 1
}

Write-Host "Using SignTool: $signTool" -ForegroundColor Cyan

# 要签名的文件列表
$filesToSign = @()

# 查找需要签名的文件
$xlam = Get-ChildItem $distPath -Filter "*.xlam" -ErrorAction SilentlyContinue
if ($xlam) { $filesToSign += $xlam.FullName }

$exe = Get-ChildItem $distPath -Filter "*.exe" -ErrorAction SilentlyContinue
if ($exe) { $filesToSign += $exe.FullName }

$msi = Get-ChildItem $distPath -Filter "*.msi" -ErrorAction SilentlyContinue
if ($msi) { $filesToSign += $msi.FullName }

if ($filesToSign.Count -eq 0) {
    Write-Warning "No files found to sign in $distPath"
    exit 0
}

# 签名每个文件
foreach ($file in $filesToSign) {
    Write-Host "Signing: $file" -ForegroundColor Yellow
    
    $arguments = @(
        "sign",
        "/f", "`"$CertificatePath`"",
        "/p", "`"$CertificatePassword`"",
        "/t", $TimestampUrl,
        "/v",
        "`"$file`""
    )
    
    try {
        $process = Start-Process -FilePath $signTool -ArgumentList $arguments -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Host "✓ Successfully signed: $(Split-Path $file -Leaf)" -ForegroundColor Green
        } else {
            Write-Error "✗ Failed to sign: $(Split-Path $file -Leaf) (Exit code: $($process.ExitCode))"
        }
    } catch {
        Write-Error "✗ Error signing $(Split-Path $file -Leaf): $($_.Exception.Message)"
    }
}

Write-Host "File signing completed." -ForegroundColor Green