param(
    [string]$Version = "1.0.0"
)

Write-Host "Updating release statistics for BatchCommentAddin v$Version" -ForegroundColor Green

$rootPath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$statsPath = Join-Path $rootPath "stats"

# 创建统计目录
if (!(Test-Path $statsPath)) {
    New-Item -ItemType Directory -Path $statsPath -Force | Out-Null
}

# 统计文件路径
$statsFile = Join-Path $statsPath "release-stats.json"
$downloadStatsFile = Join-Path $statsPath "download-stats.json"

# 读取现有统计数据
$stats = @{
    totalReleases = 0
    releases = @()
    lastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

if (Test-Path $statsFile) {
    try {
        $existingStats = Get-Content $statsFile -Raw | ConvertFrom-Json
        $stats.totalReleases = $existingStats.totalReleases
        $stats.releases = $existingStats.releases
    } catch {
        Write-Warning "Could not read existing stats file, starting fresh."
    }
}

# 添加新版本统计
$newRelease = @{
    version = $Version
    releaseDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    buildNumber = $env:GITHUB_RUN_NUMBER
    commitSha = $env:GITHUB_SHA
    downloads = 0
}

# 检查是否已存在该版本
$existingIndex = -1
for ($i = 0; $i -lt $stats.releases.Count; $i++) {
    if ($stats.releases[$i].version -eq $Version) {
        $existingIndex = $i
        break
    }
}

if ($existingIndex -ge 0) {
    # 更新现有版本
    $stats.releases[$existingIndex] = $newRelease
    Write-Host "Updated existing release stats for version $Version" -ForegroundColor Yellow
} else {
    # 添加新版本
    $stats.releases += $newRelease
    $stats.totalReleases++
    Write-Host "Added new release stats for version $Version" -ForegroundColor Green
}

# 更新时间戳
$stats.lastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# 保存统计数据
try {
    $stats | ConvertTo-Json -Depth 10 | Out-File $statsFile -Encoding UTF8
    Write-Host "Release statistics updated successfully" -ForegroundColor Green
} catch {
    Write-Error "Failed to save release statistics: $($_.Exception.Message)"
}

# 创建下载统计模板
$downloadStats = @{
    totalDownloads = 0
    downloadsByVersion = @{}
    downloadsByDate = @{}
    lastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

if (!(Test-Path $downloadStatsFile)) {
    try {
        $downloadStats | ConvertTo-Json -Depth 10 | Out-File $downloadStatsFile -Encoding UTF8
        Write-Host "Download statistics template created" -ForegroundColor Green
    } catch {
        Write-Error "Failed to create download statistics template: $($_.Exception.Message)"
    }
}

# 生成发布报告
$reportPath = Join-Path $statsPath "release-report-$Version.md"
$report = @"
# Release Report - BatchCommentAddin v$Version

## Release Information
- **Version**: $Version
- **Release Date**: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
- **Build Number**: $($env:GITHUB_RUN_NUMBER)
- **Commit SHA**: $($env:GITHUB_SHA)

## Statistics
- **Total Releases**: $($stats.totalReleases)
- **Previous Versions**: $($stats.releases.Count - 1)

## Build Artifacts
- BatchCommentAddin.xlam
- BatchCommentAddin-Setup-$Version.exe
- BatchCommentAddin-$Version.msi
- BatchCommentAddin-$Version.zip

## Release Notes
Please refer to CHANGELOG.md for detailed release notes.

## Download Links
- [GitHub Releases](https://github.com/yourusername/BatchCommentAddin/releases/tag/v$Version)

---
Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@

try {
    $report | Out-File $reportPath -Encoding UTF8
    Write-Host "Release report generated: $reportPath" -ForegroundColor Green
} catch {
    Write-Error "Failed to generate release report: $($_.Exception.Message)"
}

Write-Host "Statistics update completed for version $Version" -ForegroundColor Green