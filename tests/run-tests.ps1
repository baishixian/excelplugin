Write-Host "Running BatchCommentAddin Tests" -ForegroundColor Green

$rootPath = Split-Path -Parent $PSScriptRoot
$testPath = Join-Path $rootPath "tests"
$resultsPath = Join-Path $testPath "results"

# 创建结果目录
if (!(Test-Path $resultsPath)) {
    New-Item -ItemType Directory -Path $resultsPath -Force
}

# 运行单元测试
Write-Host "Running unit tests..." -ForegroundColor Yellow
$unitTestPath = Join-Path $testPath "unit"
$unitTests = Get-ChildItem $unitTestPath -Filter "*.ps1" -ErrorAction SilentlyContinue

$totalTests = 0
$passedTests = 0
$failedTests = 0
$testResults = @()

if ($unitTests) {
    foreach ($test in $unitTests) {
        Write-Host "Running $($test.Name)..." -ForegroundColor Cyan
        
        $testResult = @{
            Name = $test.Name
            Status = "Unknown"
            Duration = 0
            Error = $null
            StartTime = Get-Date
        }
        
        try {
            $startTime = Get-Date
            & $test.FullName
            $endTime = Get-Date
            $testResult.Duration = ($endTime - $startTime).TotalSeconds
            $testResult.Status = "Passed"
            $passedTests++
            Write-Host "✓ PASSED ($($testResult.Duration.ToString("F2"))s)" -ForegroundColor Green
        } catch {
            $endTime = Get-Date
            $testResult.Duration = ($endTime - $startTime).TotalSeconds
            $testResult.Status = "Failed"
            $testResult.Error = $_.Exception.Message
            $failedTests++
            Write-Host "✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        $testResults += $testResult
        $totalTests++
    }
} else {
    Write-Host "No unit tests found in $unitTestPath" -ForegroundColor Yellow
}

# 运行集成测试
Write-Host "`nRunning integration tests..." -ForegroundColor Yellow
$integrationTestPath = Join-Path $testPath "integration"
$integrationTests = Get-ChildItem $integrationTestPath -Filter "*.ps1" -ErrorAction SilentlyContinue

if ($integrationTests) {
    foreach ($test in $integrationTests) {
        Write-Host "Running $($test.Name)..." -ForegroundColor Cyan
        
        $testResult = @{
            Name = $test.Name
            Status = "Unknown"
            Duration = 0
            Error = $null
            StartTime = Get-Date
        }
        
        try {
            $startTime = Get-Date
            & $test.FullName
            $endTime = Get-Date
            $testResult.Duration = ($endTime - $startTime).TotalSeconds
            $testResult.Status = "Passed"
            $passedTests++
            Write-Host "✓ PASSED ($($testResult.Duration.ToString("F2"))s)" -ForegroundColor Green
        } catch {
            $endTime = Get-Date
            $testResult.Duration = ($endTime - $startTime).TotalSeconds
            $testResult.Status = "Failed"
            $testResult.Error = $_.Exception.Message
            $failedTests++
            Write-Host "✗ FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        $testResults += $testResult
        $totalTests++
    }
} else {
    Write-Host "No integration tests found in $integrationTestPath" -ForegroundColor Yellow
}

# 生成测试报告
$report = @{
    TestRun = @{
        Total = $totalTests
        Passed = $passedTests
        Failed = $failedTests
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Duration = ($testResults | Measure-Object -Property Duration -Sum).Sum
        Results = $testResults
    }
}

$report | ConvertTo-Json -Depth 10 | Out-File (Join-Path $resultsPath "test-results.json") -Encoding UTF8

# 生成HTML报告
$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <title>BatchCommentAddin Test Results</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #f0f0f0; padding: 20px; border-radius: 5px; }
        .summary { margin: 20px 0; }
        .passed { color: green; }
        .failed { color: red; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .test-passed { background-color: #d4edda; }
        .test-failed { background-color: #f8d7da; }
    </style>
</head>
<body>
    <div class="header">
        <h1>BatchCommentAddin Test Results</h1>
        <p>Generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
    </div>
    
    <div class="summary">
        <h2>Summary</h2>
        <p><strong>Total Tests:</strong> $totalTests</p>
        <p><strong class="passed">Passed:</strong> $passedTests</p>
        <p><strong class="failed">Failed:</strong> $failedTests</p>
        <p><strong>Success Rate:</strong> $(if($totalTests -gt 0) { [math]::Round(($passedTests / $totalTests) * 100, 2) } else { 0 })%</p>
        <p><strong>Total Duration:</strong> $([math]::Round($report.TestRun.Duration, 2)) seconds</p>
    </div>
    
    <h2>Test Details</h2>
    <table>
        <tr>
            <th>Test Name</th>
            <th>Status</th>
            <th>Duration (s)</th>
            <th>Error</th>
        </tr>
"@

foreach ($result in $testResults) {
    $cssClass = if ($result.Status -eq "Passed") { "test-passed" } else { "test-failed" }
    $error = if ($result.Error) { $result.Error } else { "" }
    $htmlReport += @"
        <tr class="$cssClass">
            <td>$($result.Name)</td>
            <td>$($result.Status)</td>
            <td>$([math]::Round($result.Duration, 2))</td>
            <td>$error</td>
        </tr>
"@
}

$htmlReport += @"
    </table>
</body>
</html>
"@

$htmlReport | Out-File (Join-Path $resultsPath "test-results.html") -Encoding UTF8

Write-Host "`nTest Summary:" -ForegroundColor Yellow
Write-Host "Total: $totalTests" -ForegroundColor White
Write-Host "Passed: $passedTests" -ForegroundColor Green
Write-Host "Failed: $failedTests" -ForegroundColor Red

if ($totalTests -gt 0) {
    $successRate = [math]::Round(($passedTests / $totalTests) * 100, 2)
    Write-Host "Success Rate: $successRate%" -ForegroundColor Cyan
}

Write-Host "Total Duration: $([math]::Round($report.TestRun.Duration, 2)) seconds" -ForegroundColor Cyan
Write-Host "Results saved to: $resultsPath" -ForegroundColor Cyan

if ($failedTests -gt 0) {
    Write-Host "`nSome tests failed. Check the detailed report for more information." -ForegroundColor Red
    exit 1
}

Write-Host "`nAll tests passed!" -ForegroundColor Green
exit 0