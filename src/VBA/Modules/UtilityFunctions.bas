Attribute VB_Name = "UtilityFunctions"
' ===============================================================================
' 模块: UtilityFunctions
' 描述: 工具函数集合
' 版本: 1.0.0
' 作者: OneDay Team
' 创建: 2024-01-01
' ===============================================================================

Option Explicit

' ===============================================================================
' 字符串处理函数
' ===============================================================================

' 去除字符串两端空格并处理空值
Public Function SafeTrim(inputStr As Variant) As String
    On Error Resume Next
    
    If IsNull(inputStr) Or IsEmpty(inputStr) Then
        SafeTrim = ""
    Else
        SafeTrim = Trim(CStr(inputStr))
    End If
    
    On Error GoTo 0
End Function

' 字符串是否为空或空白
Public Function IsStringEmpty(inputStr As Variant) As Boolean
    IsStringEmpty = (SafeTrim(inputStr) = "")
End Function

' 替换字符串中的占位符
Public Function ReplacePlaceholders(text As String, cell As Range) As String
    On Error Resume Next
    
    Dim result As String
    result = text
    
    ' 替换基本占位符
    result = Replace(result, "{CELL}", cell.Address)
    result = Replace(result, "{VALUE}", CStr(cell.Value))
    result = Replace(result, "{ROW}", CStr(cell.Row))
    result = Replace(result, "{COLUMN}", CStr(cell.Column))
    result = Replace(result, "{SHEET}", cell.Worksheet.Name)
    result = Replace(result, "{WORKBOOK}", cell.Worksheet.Parent.Name)
    
    ' 替换日期时间占位符
    result = Replace(result, "{DATE}", Format(Date, "yyyy-mm-dd"))
    result = Replace(result, "{TIME}", Format(Time, "hh:mm:ss"))
    result = Replace(result, "{DATETIME}", Format(Now, "yyyy-mm-dd hh:mm:ss"))
    
    ' 替换用户信息占位符
    result = Replace(result, "{USER}", Environ("USERNAME"))
    result = Replace(result, "{COMPUTER}", Environ("COMPUTERNAME"))
    
    ReplacePlaceholders = result
    
    On Error GoTo 0
End Function

' 分割字符串为数组
Public Function SplitString(inputStr As String, delimiter As String) As String()
    On Error GoTo ErrorHandler
    
    If inputStr = "" Then
        Dim emptyArray(0) As String
        SplitString = emptyArray
        Exit Function
    End If
    
    SplitString = Split(inputStr, delimiter)
    Exit Function
    
ErrorHandler:
    Dim errorArray(0) As String
    SplitString = errorArray
End Function

' ===============================================================================
' 区域处理函数
' ===============================================================================

' 验证区域地址是否有效
Public Function IsValidRange(rangeAddress As String) As Boolean
    On Error GoTo ErrorHandler
    
    IsValidRange = False
    
    If SafeTrim(rangeAddress) = "" Then Exit Function
    
    ' 尝试创建区域对象
    Dim testRange As Range
    Set testRange = Range(rangeAddress)
    
    IsValidRange = True
    Exit Function
    
ErrorHandler:
    IsValidRange = False
End Function

' 获取区域的行数
Public Function GetRangeRows(rangeAddress As String) As Long
    On Error GoTo ErrorHandler
    
    GetRangeRows = 0
    
    If Not IsValidRange(rangeAddress) Then Exit Function
    
    Dim targetRange As Range
    Set targetRange = Range(rangeAddress)
    
    GetRangeRows = targetRange.Rows.Count
    Exit Function
    
ErrorHandler:
    GetRangeRows = 0
End Function

' 获取区域的列数
Public Function GetRangeColumns(rangeAddress As String) As Long
    On Error GoTo ErrorHandler
    
    GetRangeColumns = 0
    
    If Not IsValidRange(rangeAddress) Then Exit Function
    
    Dim targetRange As Range
    Set targetRange = Range(rangeAddress)
    
    GetRangeColumns = targetRange.Columns.Count
    Exit Function
    
ErrorHandler:
    GetRangeColumns = 0
End Function

' 获取区域的单元格数量
Public Function GetRangeCellCount(rangeAddress As String) As Long
    On Error GoTo ErrorHandler
    
    GetRangeCellCount = 0
    
    If Not IsValidRange(rangeAddress) Then Exit Function
    
    Dim targetRange As Range
    Set targetRange = Range(rangeAddress)
    
    GetRangeCellCount = targetRange.Cells.Count
    Exit Function
    
ErrorHandler:
    GetRangeCellCount = 0
End Function

' 检查区域是否包含数据
Public Function RangeHasData(rangeAddress As String) As Boolean
    On Error GoTo ErrorHandler
    
    RangeHasData = False
    
    If Not IsValidRange(rangeAddress) Then Exit Function
    
    Dim targetRange As Range
    Set targetRange = Range(rangeAddress)
    
    Dim cell As Range
    For Each cell In targetRange
        If Not IsEmpty(cell.Value) And SafeTrim(cell.Value) <> "" Then
            RangeHasData = True
            Exit Function
        End If
    Next cell
    
    Exit Function
    
ErrorHandler:
    RangeHasData = False
End Function

' ===============================================================================
' 文件处理函数
' ===============================================================================

' 检查文件是否存在
Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

' 获取文件扩展名
Public Function GetFileExtension(filePath As String) As String
    On Error Resume Next
    
    Dim pos As Long
    pos = InStrRev(filePath, ".")
    
    If pos > 0 Then
        GetFileExtension = LCase(Mid(filePath, pos + 1))
    Else
        GetFileExtension = ""
    End If
    
    On Error GoTo 0
End Function

' 从文件加载文本内容
Public Function LoadTextFromFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    LoadTextFromFile = ""
    
    If Not FileExists(filePath) Then Exit Function
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    
    Dim content As String
    Dim line As String
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        content = content & line & vbCrLf
    Loop
    
    Close #fileNum
    
    ' 移除最后的换行符
    If Len(content) > 2 Then
        content = Left(content, Len(content) - 2)
    End If
    
    LoadTextFromFile = content
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    LoadTextFromFile = ""
End Function

' 从CSV文件加载数据
Public Function LoadDataFromCSV(filePath As String) As Variant
    On Error GoTo ErrorHandler
    
    LoadDataFromCSV = Empty
    
    If Not FileExists(filePath) Then Exit Function
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    
    ' 计算行数
    Dim lineCount As Long
    lineCount = 0
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, ""
        lineCount = lineCount + 1
    Loop
    
    Close #fileNum
    
    If lineCount = 0 Then Exit Function
    
    ' 重新打开文件读取数据
    Open filePath For Input As #fileNum
    
    Dim data() As String
    ReDim data(1 To lineCount)
    
    Dim i As Long
    i = 1
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, data(i)
        i = i + 1
    Loop
    
    Close #fileNum
    
    LoadDataFromCSV = data
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    LoadDataFromCSV = Empty
End Function

' 从Excel文件加载批注数据
Public Function LoadCommentFromFile(filePath As String) As Variant
    On Error GoTo ErrorHandler
    
    LoadCommentFromFile = Empty
    
    If Not FileExists(filePath) Then Exit Function
    
    Dim fileExt As String
    fileExt = GetFileExtension(filePath)
    
    Select Case fileExt
        Case "txt"
            LoadCommentFromFile = LoadTextFromFile(filePath)
            
        Case "csv"
            LoadCommentFromFile = LoadDataFromCSV(filePath)
            
        Case "xlsx", "xls"
            LoadCommentFromFile = LoadDataFromExcel(filePath)
            
        Case Else
            ' 默认按文本文件处理
            LoadCommentFromFile = LoadTextFromFile(filePath)
    End Select
    
    Exit Function
    
ErrorHandler:
    LoadCommentFromFile = Empty
End Function

' 从Excel文件加载数据
Private Function LoadDataFromExcel(filePath As String) As Variant
    On Error GoTo ErrorHandler
    
    LoadDataFromExcel = Empty
    
    Dim excelApp As Object
    Dim workbook As Object
    Dim worksheet As Object
    
    ' 创建Excel应用程序实例
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    
    ' 打开工作簿
    Set workbook = excelApp.Workbooks.Open(filePath)
    Set worksheet = workbook.Worksheets(1)
    
    ' 获取使用区域
    Dim usedRange As Object
    Set usedRange = worksheet.UsedRange
    
    If usedRange.Cells.Count > 0 Then
        LoadDataFromExcel = usedRange.Value
    End If
    
    ' 清理资源
    workbook.Close False
    excelApp.Quit
    
    Set worksheet = Nothing
    Set workbook = Nothing
    Set excelApp = Nothing
    
    Exit Function
    
ErrorHandler:
    ' 清理资源
    If Not workbook Is Nothing Then workbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
    
    Set worksheet = Nothing
    Set workbook = Nothing
    Set excelApp = Nothing
    
    LoadDataFromExcel = Empty
End Function

' ===============================================================================
' 系统工具函数
' ===============================================================================

' 清理临时文件
Public Sub CleanupTempFiles()
    On Error Resume Next
    
    Dim tempPath As String
    tempPath = g_ConfigPath & "temp\"
    
    If Dir(tempPath, vbDirectory) <> "" Then
        ' 删除临时目录中的所有文件
        Dim fileName As String
        fileName = Dir(tempPath & "*.*")
        
        Do While fileName <> ""
            Kill tempPath & fileName
            fileName = Dir
        Loop
    End If
    
    On Error GoTo 0
End Sub

' 记录操作日志
Public Sub LogOperation(operationType As String, description As String)
    On Error Resume Next
    
    Dim logFile As String
    logFile = g_LogPath & "operations_" & Format(Date, "yyyymmdd") & ".log"
    
    ' 确保日志目录存在
    If Dir(g_LogPath, vbDirectory) = "" Then
        MkDir g_LogPath
    End If
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As #fileNum
    
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:mm:ss") & " [" & operationType & "] " & description
    
    Close #fileNum
    
    On Error GoTo 0
End Sub

' 创建工作表备份
Public Sub CreateBackup(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim backupName As String
    backupName = ws.Name & "_Backup_" & Format(Now, "yyyymmdd_hhmmss")
    
    ' 复制工作表
    ws.Copy After:=ws
    
    ' 重命名备份工作表
    ActiveSheet.Name = backupName
    
    ' 记录备份操作
    Call LogOperation("BACKUP", "创建工作表备份: " & backupName)
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "UtilityFunctions", "CreateBackup")
End Sub

' ===============================================================================
' 用户界面工具函数
' ===============================================================================

' 显示信息消息
Public Sub ShowInfo(message As String)
    MsgBox message, vbInformation, ADDIN_NAME
End Sub

' 显示警告消息
Public Sub ShowWarning(message As String)
    MsgBox message, vbExclamation, ADDIN_NAME
End Sub

' 显示错误消息
Public Sub ShowError(message As String)
    MsgBox message, vbCritical, ADDIN_NAME
End Sub

' 显示确认对话框
Public Function ShowConfirm(message As String) As Boolean
    ShowConfirm = (MsgBox(message, vbYesNo + vbQuestion, ADDIN_NAME) = vbYes)
End Function

' ===============================================================================
' 数据验证函数
' ===============================================================================

' 验证数字范围
Public Function IsNumberInRange(value As Variant, minValue As Double, maxValue As Double) As Boolean
    On Error Resume Next
    
    IsNumberInRange = False
    
    If IsNumeric(value) Then
        Dim numValue As Double
        numValue = CDbl(value)
        IsNumberInRange = (numValue >= minValue And numValue <= maxValue)
    End If
    
    On Error GoTo 0
End Function

' 验证颜色值
Public Function IsValidColor(colorValue As Variant) As Boolean
    On Error Resume Next
    
    IsValidColor = False
    
    If IsNumeric(colorValue) Then
        Dim colorNum As Long
        colorNum = CLng(colorValue)
        IsValidColor = (colorNum >= 0 And colorNum <= 16777215) ' RGB最大值
    End If
    
    On Error GoTo 0
End Function

' 验证字体名称
Public Function IsValidFontName(fontName As String) As Boolean
    On Error GoTo ErrorHandler
    
    IsValidFontName = False
    
    If SafeTrim(fontName) = "" Then Exit Function
    
    ' 尝试设置字体名称来验证
    Dim testRange As Range
    Set testRange = ActiveSheet.Range("A1")
    
    Dim originalFont As String
    originalFont = testRange.Font.Name
    
    testRange.Font.Name = fontName
    IsValidFontName = (testRange.Font.Name = fontName)
    
    ' 恢复原字体
    testRange.Font.Name = originalFont
    
    Exit Function
    
ErrorHandler:
    IsValidFontName = False
End Function

' ===============================================================================
' 性能优化函数
' ===============================================================================

' 禁用屏幕更新和事件
Public Sub DisableUpdates()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

' 启用屏幕更新和事件
Public Sub EnableUpdates()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' 获取系统性能信息
Public Function GetSystemInfo() As String
    On Error Resume Next
    
    Dim info As String
    info = "系统信息:" & vbCrLf
    info = info & "Excel版本: " & Application.Version & vbCrLf
    info = info & "操作系统: " & Application.OperatingSystem & vbCrLf
    info = info & "用户名: " & Environ("USERNAME") & vbCrLf
    info = info & "计算机名: " & Environ("COMPUTERNAME") & vbCrLf
    info = info & "内存使用: " & Application.MemoryUsed & " KB" & vbCrLf
    info = info & "当前时间: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    GetSystemInfo = info
    
    On Error GoTo 0
End Function