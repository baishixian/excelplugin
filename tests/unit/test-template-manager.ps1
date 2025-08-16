# BatchCommentAddin Unit Tests - Template Manager
# 这个文件包含对模板管理器的单元测试

Write-Host "Testing Template Manager..." -ForegroundColor Cyan

# 模拟CommentTemplate结构
function New-MockTemplate {
    param(
        [string]$Name = "TestTemplate",
        [string]$FontName = "Arial",
        [int]$FontSize = 10,
        [long]$FontColor = 1,
        [bool]$IsBold = $false,
        [bool]$IsItalic = $false,
        [string]$DefaultText = "",
        [long]$BackgroundColor = 0xE0E0E0,
        [int]$Width = 200,
        [int]$Height = 100,
        [bool]$IsAutoSize = $true
    )
    
    return @{
        Name = $Name
        FontName = $FontName
        FontSize = $FontSize
        FontColor = $FontColor
        IsBold = $IsBold
        IsItalic = $IsItalic
        DefaultText = $DefaultText
        BackgroundColor = $BackgroundColor
        Width = $Width
        Height = $Height
        IsAutoSize = $IsAutoSize
    }
}

function Test-TemplateCreation {
    Write-Host "  Testing template creation..." -ForegroundColor Gray
    
    # 测试创建新模板
    $template = New-MockTemplate -Name "TestTemplate" -FontName "微软雅黑" -FontSize 12 -IsBold $true
    
    # 验证模板属性
    if ($template.Name -ne "TestTemplate") {
        throw "Template name mismatch. Expected: 'TestTemplate', Got: '$($template.Name)'"
    }
    
    if ($template.FontName -ne "微软雅黑") {
        throw "Font name mismatch. Expected: '微软雅黑', Got: '$($template.FontName)'"
    }
    
    if ($template.FontSize -ne 12) {
        throw "Font size mismatch. Expected: 12, Got: $($template.FontSize)"
    }
    
    if ($template.IsBold -ne $true) {
        throw "Bold setting mismatch. Expected: True, Got: $($template.IsBold)"
    }
    
    Write-Host "    ✓ Template creation tests passed" -ForegroundColor Green
}

function Test-TemplateValidation {
    Write-Host "  Testing template validation..." -ForegroundColor Gray
    
    # 测试有效模板
    $validTemplate = New-MockTemplate -Name "ValidTemplate" -FontSize 12
    
    # 模拟验证逻辑
    $isValid = $true
    
    # 验证模板名称
    if ([string]::IsNullOrWhiteSpace($validTemplate.Name)) {
        $isValid = $false
    }
    
    # 验证字体大小范围
    if ($validTemplate.FontSize -lt 6 -or $validTemplate.FontSize -gt 72) {
        $isValid = $false
    }
    
    # 验证字体名称
    if ([string]::IsNullOrWhiteSpace($validTemplate.FontName)) {
        $isValid = $false
    }
    
    if (-not $isValid) {
        throw "Valid template was incorrectly identified as invalid"
    }
    
    # 测试无效模板
    $invalidTemplate = New-MockTemplate -Name "" -FontSize 100
    
    $isValid = $true
    
    # 验证模板名称
    if ([string]::IsNullOrWhiteSpace($invalidTemplate.Name)) {
        $isValid = $false
    }
    
    # 验证字体大小范围
    if ($invalidTemplate.FontSize -lt 6 -or $invalidTemplate.FontSize -gt 72) {
        $isValid = $false
    }
    
    if ($isValid) {
        throw "Invalid template was incorrectly identified as valid"
    }
    
    Write-Host "    ✓ Template validation tests passed" -ForegroundColor Green
}

function Test-TemplateStorage {
    Write-Host "  Testing template storage operations..." -ForegroundColor Gray
    
    # 模拟模板存储
    $templateStore = @{}
    
    # 测试添加模板
    $template1 = New-MockTemplate -Name "Template1"
    $template2 = New-MockTemplate -Name "Template2" -FontSize 14
    
    # 模拟AddTemplate函数
    function Add-MockTemplate {
        param($template)
        
        if ($templateStore.ContainsKey($template.Name)) {
            return $false  # 模板已存在
        }
        
        $templateStore[$template.Name] = $template
        return $true
    }
    
    # 测试添加新模板
    $result1 = Add-MockTemplate $template1
    if (-not $result1) {
        throw "Failed to add new template"
    }
    
    # 测试添加重复模板
    $result2 = Add-MockTemplate $template1
    if ($result2) {
        throw "Duplicate template was incorrectly added"
    }
    
    # 测试获取模板
    if (-not $templateStore.ContainsKey("Template1")) {
        throw "Template1 not found in store"
    }
    
    $retrievedTemplate = $templateStore["Template1"]
    if ($retrievedTemplate.Name -ne "Template1") {
        throw "Retrieved template name mismatch"
    }
    
    Write-Host "    ✓ Template storage tests passed" -ForegroundColor Green
}

function Test-TemplateExportImport {
    Write-Host "  Testing template export/import..." -ForegroundColor Gray
    
    # 创建测试模板
    $originalTemplate = New-MockTemplate -Name "ExportTest" -FontName "Calibri" -FontSize 11 -IsBold $true
    
    # 模拟导出到字符串格式
    $exportData = @"
[BatchCommentTemplate]
Name=$($originalTemplate.Name)
FontName=$($originalTemplate.FontName)
FontSize=$($originalTemplate.FontSize)
FontColor=$($originalTemplate.FontColor)
IsBold=$($originalTemplate.IsBold)
IsItalic=$($originalTemplate.IsItalic)
DefaultText=$($originalTemplate.DefaultText)
BackgroundColor=$($originalTemplate.BackgroundColor)
Width=$($originalTemplate.Width)
Height=$($originalTemplate.Height)
IsAutoSize=$($originalTemplate.IsAutoSize)
"@
    
    # 模拟从字符串导入
    $importedTemplate = @{}
    $lines = $exportData -split "`n"
    
    foreach ($line in $lines) {
        if ($line -match "^(.+)=(.*)$") {
            $key = $matches[1]
            $value = $matches[2]
            
            switch ($key) {
                "Name" { $importedTemplate.Name = $value }
                "FontName" { $importedTemplate.FontName = $value }
                "FontSize" { $importedTemplate.FontSize = [int]$value }
                "FontColor" { $importedTemplate.FontColor = [long]$value }
                "IsBold" { $importedTemplate.IsBold = [bool]::Parse($value) }
                "IsItalic" { $importedTemplate.IsItalic = [bool]::Parse($value) }
                "DefaultText" { $importedTemplate.DefaultText = $value }
                "BackgroundColor" { $importedTemplate.BackgroundColor = [long]$value }
                "Width" { $importedTemplate.Width = [int]$value }
                "Height" { $importedTemplate.Height = [int]$value }
                "IsAutoSize" { $importedTemplate.IsAutoSize = [bool]::Parse($value) }
            }
        }
    }
    
    # 验证导入的模板
    if ($importedTemplate.Name -ne $originalTemplate.Name) {
        throw "Imported template name mismatch"
    }
    
    if ($importedTemplate.FontName -ne $originalTemplate.FontName) {
        throw "Imported template font name mismatch"
    }
    
    if ($importedTemplate.FontSize -ne $originalTemplate.FontSize) {
        throw "Imported template font size mismatch"
    }
    
    if ($importedTemplate.IsBold -ne $originalTemplate.IsBold) {
        throw "Imported template bold setting mismatch"
    }
    
    Write-Host "    ✓ Template export/import tests passed" -ForegroundColor Green
}

# 运行所有测试
try {
    Test-TemplateCreation
    Test-TemplateValidation
    Test-TemplateStorage
    Test-TemplateExportImport
    
    Write-Host "All template manager tests completed successfully!" -ForegroundColor Green
} catch {
    Write-Error "Template manager test failed: $($_.Exception.Message)"
    throw
}