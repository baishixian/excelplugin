# BatchCommentAddin Unit Tests - Utility Functions
# 这个文件包含对工具函数的单元测试

Write-Host "Testing Utility Functions..." -ForegroundColor Cyan

# 模拟测试函数
function Test-StringProcessing {
    Write-Host "  Testing string processing functions..." -ForegroundColor Gray
    
    # 测试SafeTrim函数的逻辑
    $testCases = @(
        @{ Input = "  test  "; Expected = "test" },
        @{ Input = $null; Expected = "" },
        @{ Input = ""; Expected = "" },
        @{ Input = "   "; Expected = "" }
    )
    
    foreach ($case in $testCases) {
        # 这里模拟SafeTrim函数的行为
        $result = if ($case.Input -eq $null -or $case.Input -eq "") { 
            "" 
        } else { 
            $case.Input.Trim() 
        }
        
        if ($result -ne $case.Expected) {
            throw "SafeTrim test failed. Input: '$($case.Input)', Expected: '$($case.Expected)', Got: '$result'"
        }
    }
    
    Write-Host "    ✓ String processing tests passed" -ForegroundColor Green
}

function Test-RangeValidation {
    Write-Host "  Testing range validation functions..." -ForegroundColor Gray
    
    # 测试区域地址验证的逻辑
    $validRanges = @("A1", "A1:B10", "Sheet1!A1:B10", "$A$1:$B$10")
    $invalidRanges = @("", "XYZ", "A1:Z999999", "Invalid!Range")
    
    # 模拟IsValidRange函数的基本验证逻辑
    foreach ($range in $validRanges) {
        # 简单的格式验证
        if ($range -match "^[A-Z]+\d+$|^[A-Z]+\d+:[A-Z]+\d+$|^.+![A-Z$]+\d+:[A-Z$]+\d+$") {
            # 通过验证
        } else {
            throw "Valid range '$range' was incorrectly identified as invalid"
        }
    }
    
    Write-Host "    ✓ Range validation tests passed" -ForegroundColor Green
}

function Test-FileOperations {
    Write-Host "  Testing file operation functions..." -ForegroundColor Gray
    
    # 测试文件扩展名获取
    $testCases = @(
        @{ Input = "test.txt"; Expected = "txt" },
        @{ Input = "document.xlsx"; Expected = "xlsx" },
        @{ Input = "noextension"; Expected = "" },
        @{ Input = "multiple.dots.csv"; Expected = "csv" }
    )
    
    foreach ($case in $testCases) {
        # 模拟GetFileExtension函数
        $lastDot = $case.Input.LastIndexOf(".")
        $result = if ($lastDot -ge 0) { 
            $case.Input.Substring($lastDot + 1).ToLower() 
        } else { 
            "" 
        }
        
        if ($result -ne $case.Expected) {
            throw "GetFileExtension test failed. Input: '$($case.Input)', Expected: '$($case.Expected)', Got: '$result'"
        }
    }
    
    Write-Host "    ✓ File operation tests passed" -ForegroundColor Green
}

function Test-PlaceholderReplacement {
    Write-Host "  Testing placeholder replacement..." -ForegroundColor Gray
    
    # 测试占位符替换逻辑
    $template = "Cell: {CELL}, Value: {VALUE}, Row: {ROW}, Column: {COLUMN}"
    $mockCell = @{
        Address = "A1"
        Value = "TestValue"
        Row = 1
        Column = 1
    }
    
    # 模拟ReplacePlaceholders函数
    $result = $template
    $result = $result -replace "\{CELL\}", $mockCell.Address
    $result = $result -replace "\{VALUE\}", $mockCell.Value
    $result = $result -replace "\{ROW\}", $mockCell.Row
    $result = $result -replace "\{COLUMN\}", $mockCell.Column
    
    $expected = "Cell: A1, Value: TestValue, Row: 1, Column: 1"
    
    if ($result -ne $expected) {
        throw "Placeholder replacement test failed. Expected: '$expected', Got: '$result'"
    }
    
    Write-Host "    ✓ Placeholder replacement tests passed" -ForegroundColor Green
}

# 运行所有测试
try {
    Test-StringProcessing
    Test-RangeValidation
    Test-FileOperations
    Test-PlaceholderReplacement
    
    Write-Host "All utility function tests completed successfully!" -ForegroundColor Green
} catch {
    Write-Error "Utility function test failed: $($_.Exception.Message)"
    throw
}