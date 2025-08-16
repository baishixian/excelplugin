Attribute VB_Name = "ErrorHandler"
' ===============================================================================
' 模块: ErrorHandler
' 描述: 统一错误处理模块
' 版本: 1.0.0
' 作者: OneDay Team
' 创建: 2024-01-01
' ===============================================================================

Option Explicit

' 错误类型枚举
Public Enum ErrorType
    etGeneral = 0           ' 一般错误
    etFileAccess = 1        ' 文件访问错误
    etRangeError = 2        ' 区域错误
    etDataValidation = 3    ' 数据验证错误
    etTemplateError = 4     ' 模板错误
    etUserInterface = 5     ' 用户界面错误
    etSystemError = 6       ' 系统错误
End Enum

' 错误信息结构
Public Type ErrorInfo
    ErrorType As ErrorType
    ErrorNumber As Long
    ErrorDescription As String
    ModuleName As String
    ProcedureName As String
    Timestamp As Date
    UserAction As String
    SystemInfo As String
End Type

' 全局变量
Private m_ErrorLog() As ErrorInfo
Private m_ErrorCount As Long
Private m_MaxLogEntries As Long
Private m_IsLoggingEnabled As Boolean

' ===============================================================================
' 错误处理器初始化
' ===============================================================================

' 初始化错误处理器
Public Sub InitializeErrorHandler()
    On Error Resume Next
    
    m_MaxLogEntries = 1000
    m_ErrorCount = 0
    m_IsLoggingEnabled = True
    
    ReDim m_ErrorLog(0 To m_MaxLogEntries - 1)
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 主要错误处理函数
' ===============================================================================

' 处理错误的主函数
Public Sub HandleError(errorType As ErrorType, errorNumber As Long, errorDescription As String, _
                      moduleName As String, procedureName As String, Optional userAction As String = "")
    On Error Resume Next
    
    ' 创建错误信息
    Dim errorInfo As ErrorInfo
    With errorInfo
        .ErrorType = errorType
        .ErrorNumber = errorNumber
        .ErrorDescription = errorDescription
        .ModuleName = moduleName
        .ProcedureName = procedureName
        .Timestamp = Now
        .UserAction = userAction
        .SystemInfo = GetBasicSystemInfo()
    End With
    
    ' 记录错误
    Call LogError(errorInfo)
    
    ' 显示错误信息（根据错误类型决定）
    Call DisplayError(errorInfo)
    
    ' 执行错误恢复操作
    Call RecoverFromError(errorInfo)
    
    On Error GoTo 0
End Sub

' 记录错误到内存日志
Private Sub LogError(errorInfo As ErrorInfo)
    On Error Resume Next
    
    If Not m_IsLoggingEnabled Then Exit Sub
    
    ' 如果日志已满，移动数组
    If m_ErrorCount >= m_MaxLogEntries Then
        Dim i As Long
        For i = 0 To m_MaxLogEntries - 2
            m_ErrorLog(i) = m_ErrorLog(i + 1)
        Next i
        m_ErrorCount = m_MaxLogEntries - 1
    End If
    
    ' 添加新错误
    m_ErrorLog(m_ErrorCount) = errorInfo
    m_ErrorCount = m_ErrorCount + 1
    
    ' 写入文件日志
    Call WriteErrorToFile(errorInfo)
    
    On Error GoTo 0
End Sub

' 将错误写入文件
Private Sub WriteErrorToFile(errorInfo As ErrorInfo)
    On Error Resume Next
    
    Dim logFile As String
    logFile = g_LogPath & "errors_" & Format(Date, "yyyymmdd") & ".log"
    
    ' 确保日志目录存在
    If Dir(g_LogPath, vbDirectory) = "" Then
        MkDir g_LogPath
    End If
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As #fileNum
    
    With errorInfo
        Print #fileNum, "==================== ERROR LOG ===================="
        Print #fileNum, "时间: " & Format(.Timestamp, "yyyy-mm-dd hh:mm:ss")
        Print #fileNum, "类型: " & GetErrorTypeName(.ErrorType)
        Print #fileNum, "编号: " & .ErrorNumber
        Print #fileNum, "描述: " & .ErrorDescription
        Print #fileNum, "模块: " & .ModuleName
        Print #fileNum, "过程: " & .ProcedureName
        If .UserAction <> "" Then Print #fileNum, "用户操作: " & .UserAction
        Print #fileNum, "系统信息: " & .SystemInfo
        Print #fileNum, "=================================================="
        Print #fileNum, ""
    End With
    
    Close #fileNum
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 错误显示函数
' ===============================================================================

' 显示错误信息
Private Sub DisplayError(errorInfo As ErrorInfo)
    On Error Resume Next
    
    Dim message As String
    Dim title As String
    Dim icon As VbMsgBoxStyle
    
    ' 根据错误类型设置显示方式
    Select Case errorInfo.ErrorType
        Case etGeneral
            title = "一般错误"
            icon = vbExclamation
            message = "发生了一个错误：" & vbCrLf & errorInfo.ErrorDescription
            
        Case etFileAccess
            title = "文件访问错误"
            icon = vbCritical
            message = "文件操作失败：" & vbCrLf & errorInfo.ErrorDescription & vbCrLf & vbCrLf & _
                     "请检查文件是否存在、是否有访问权限或文件是否被其他程序占用。"
            
        Case etRangeError
            title = "区域错误"
            icon = vbExclamation
            message = "区域操作失败：" & vbCrLf & errorInfo.ErrorDescription & vbCrLf & vbCrLf & _
                     "请检查区域地址是否正确，或者选择的区域是否有效。"
            
        Case etDataValidation
            title = "数据验证错误"
            icon = vbExclamation
            message = "数据验证失败：" & vbCrLf & errorInfo.ErrorDescription & vbCrLf & vbCrLf & _
                     "请检查输入的数据是否符合要求。"
            
        Case etTemplateError
            title = "模板错误"
            icon = vbExclamation
            message = "模板操作失败：" & vbCrLf & errorInfo.ErrorDescription & vbCrLf & vbCrLf & _
                     "请检查模板设置是否正确。"
            
        Case etUserInterface
            title = "界面错误"
            icon = vbExclamation
            message = "用户界面错误：" & vbCrLf & errorInfo.ErrorDescription
            
        Case etSystemError
            title = "系统错误"
            icon = vbCritical
            message = "系统错误：" & vbCrLf & errorInfo.ErrorDescription & vbCrLf & vbCrLf & _
                     "这可能是一个严重的问题，建议重启Excel或联系技术支持。"
            
        Case Else
            title = "未知错误"
            icon = vbCritical
            message = "发生了未知错误：" & vbCrLf & errorInfo.ErrorDescription
    End Select
    
    ' 添加错误详细信息（调试模式下）
    If IsDebugMode() Then
        message = message & vbCrLf & vbCrLf & "调试信息：" & vbCrLf & _
                 "模块: " & errorInfo.ModuleName & vbCrLf & _
                 "过程: " & errorInfo.ProcedureName & vbCrLf & _
                 "错误号: " & errorInfo.ErrorNumber
    End If
    
    ' 显示错误消息
    MsgBox message, icon, ADDIN_NAME & " - " & title
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 错误恢复函数
' ===============================================================================

' 从错误中恢复
Private Sub RecoverFromError(errorInfo As ErrorInfo)
    On Error Resume Next
    
    Select Case errorInfo.ErrorType
        Case etFileAccess
            ' 文件访问错误恢复
            Call RecoverFromFileError(errorInfo)
            
        Case etRangeError
            ' 区域错误恢复
            Call RecoverFromRangeError(errorInfo)
            
        Case etSystemError
            ' 系统错误恢复
            Call RecoverFromSystemError(errorInfo)
            
        Case Else
            ' 一般错误恢复
            Call RecoverFromGeneralError(errorInfo)
    End Select
    
    On Error GoTo 0
End Sub

' 从文件错误中恢复
Private Sub RecoverFromFileError(errorInfo As ErrorInfo)
    On Error Resume Next
    
    ' 尝试创建必要的目录
    If Dir(g_ConfigPath, vbDirectory) = "" Then MkDir g_ConfigPath
    If Dir(g_LogPath, vbDirectory) = "" Then MkDir g_LogPath
    
    ' 清理可能损坏的临时文件
    Call CleanupTempFiles
    
    On Error GoTo 0
End Sub

' 从区域错误中恢复
Private Sub RecoverFromRangeError(errorInfo As ErrorInfo)
    On Error Resume Next
    
    ' 重置选择
    If Not ActiveSheet Is Nothing Then
        ActiveSheet.Range("A1").Select
    End If
    
    On Error GoTo 0
End Sub

' 从系统错误中恢复
Private Sub RecoverFromSystemError(errorInfo As ErrorInfo)
    On Error Resume Next
    
    ' 重置应用程序状态
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ' 清理资源
    Call CleanupResources
    
    On Error GoTo 0
End Sub

' 从一般错误中恢复
Private Sub RecoverFromGeneralError(errorInfo As ErrorInfo)
    On Error Resume Next
    
    ' 基本恢复操作
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 辅助函数
' ===============================================================================

' 获取错误类型名称
Private Function GetErrorTypeName(errorType As ErrorType) As String
    Select Case errorType
        Case etGeneral: GetErrorTypeName = "一般错误"
        Case etFileAccess: GetErrorTypeName = "文件访问错误"
        Case etRangeError: GetErrorTypeName = "区域错误"
        Case etDataValidation: GetErrorTypeName = "数据验证错误"
        Case etTemplateError: GetErrorTypeName = "模板错误"
        Case etUserInterface: GetErrorTypeName = "用户界面错误"
        Case etSystemError: GetErrorTypeName = "系统错误"
        Case Else: GetErrorTypeName = "未知错误"
    End Select
End Function

' 获取基本系统信息
Private Function GetBasicSystemInfo() As String
    On Error Resume Next
    
    GetBasicSystemInfo = "Excel " & Application.Version & " on " & Application.OperatingSystem
    
    On Error GoTo 0
End Function

' 检查是否为调试模式
Private Function IsDebugMode() As Boolean
    On Error Resume Next
    
    ' 可以通过注册表或配置文件设置调试模式
    ' 这里简单检查是否存在调试标志文件
    IsDebugMode = (Dir(g_ConfigPath & "debug.flag") <> "")
    
    On Error GoTo 0
End Function

' ===============================================================================
' 错误日志管理
' ===============================================================================

' 获取错误日志
Public Function GetErrorLog() As ErrorInfo()
    On Error Resume Next
    
    If m_ErrorCount = 0 Then
        Dim emptyLog(0) As ErrorInfo
        GetErrorLog = emptyLog
        Exit Function
    End If
    
    Dim result() As ErrorInfo
    ReDim result(0 To m_ErrorCount - 1)
    
    Dim i As Long
    For i = 0 To m_ErrorCount - 1
        result(i) = m_ErrorLog(i)
    Next i
    
    GetErrorLog = result
    
    On Error GoTo 0
End Function

' 清除错误日志
Public Sub ClearErrorLog()
    On Error Resume Next
    
    m_ErrorCount = 0
    ReDim m_ErrorLog(0 To m_MaxLogEntries - 1)
    
    On Error GoTo 0
End Sub

' 获取错误统计信息
Public Function GetErrorStatistics() As String
    On Error Resume Next
    
    Dim stats As String
    stats = "错误统计信息：" & vbCrLf
    stats = stats & "总错误数: " & m_ErrorCount & vbCrLf
    
    ' 按类型统计
    Dim typeCounts(0 To 6) As Long
    Dim i As Long
    
    For i = 0 To m_ErrorCount - 1
        typeCounts(m_ErrorLog(i).ErrorType) = typeCounts(m_ErrorLog(i).ErrorType) + 1
    Next i
    
    stats = stats & "一般错误: " & typeCounts(etGeneral) & vbCrLf
    stats = stats & "文件访问错误: " & typeCounts(etFileAccess) & vbCrLf
    stats = stats & "区域错误: " & typeCounts(etRangeError) & vbCrLf
    stats = stats & "数据验证错误: " & typeCounts(etDataValidation) & vbCrLf
    stats = stats & "模板错误: " & typeCounts(etTemplateError) & vbCrLf
    stats = stats & "用户界面错误: " & typeCounts(etUserInterface) & vbCrLf
    stats = stats & "系统错误: " & typeCounts(etSystemError)
    
    GetErrorStatistics = stats
    
    On Error GoTo 0
End Function

' 导出错误日志到文件
Public Function ExportErrorLog(filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ExportErrorLog = False
    
    If m_ErrorCount = 0 Then
        MsgBox "没有错误日志可以导出。", vbInformation, ADDIN_NAME
        Exit Function
    End If
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    
    Print #fileNum, "批量批注插件错误日志导出"
    Print #fileNum, "导出时间: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #fileNum, "总错误数: " & m_ErrorCount
    Print #fileNum, String(50, "=")
    Print #fileNum, ""
    
    Dim i As Long
    For i = 0 To m_ErrorCount - 1
        With m_ErrorLog(i)
            Print #fileNum, "错误 #" & (i + 1)
            Print #fileNum, "时间: " & Format(.Timestamp, "yyyy-mm-dd hh:mm:ss")
            Print #fileNum, "类型: " & GetErrorTypeName(.ErrorType)
            Print #fileNum, "编号: " & .ErrorNumber
            Print #fileNum, "描述: " & .ErrorDescription
            Print #fileNum, "模块: " & .ModuleName
            Print #fileNum, "过程: " & .ProcedureName
            If .UserAction <> "" Then Print #fileNum, "用户操作: " & .UserAction
            Print #fileNum, "系统信息: " & .SystemInfo
            Print #fileNum, String(30, "-")
            Print #fileNum, ""
        End With
    Next i
    
    Close #fileNum
    ExportErrorLog = True
    
    MsgBox "错误日志已成功导出到：" & vbCrLf & filePath, vbInformation, ADDIN_NAME
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "导出错误日志失败：" & Err.Description, vbCritical, ADDIN_NAME
    ExportErrorLog = False
End Function

' ===============================================================================
' 错误处理设置
' ===============================================================================

' 启用错误日志记录
Public Sub EnableErrorLogging()
    m_IsLoggingEnabled = True
End Sub

' 禁用错误日志记录
Public Sub DisableErrorLogging()
    m_IsLoggingEnabled = False
End Sub

' 设置最大日志条目数
Public Sub SetMaxLogEntries(maxEntries As Long)
    On Error Resume Next
    
    If maxEntries > 0 And maxEntries <= 10000 Then
        m_MaxLogEntries = maxEntries
        
        ' 重新调整数组大小
        If m_ErrorCount > maxEntries Then
            m_ErrorCount = maxEntries
        End If
        
        ReDim Preserve m_ErrorLog(0 To m_MaxLogEntries - 1)
    End If
    
    On Error GoTo 0
End Sub

' 获取错误处理器状态
Public Function GetErrorHandlerStatus() As String
    Dim status As String
    status = "错误处理器状态：" & vbCrLf
    status = status & "日志记录: " & IIf(m_IsLoggingEnabled, "启用", "禁用") & vbCrLf
    status = status & "最大日志条目: " & m_MaxLogEntries & vbCrLf
    status = status & "当前错误数: " & m_ErrorCount & vbCrLf
    status = status & "日志路径: " & g_LogPath
    
    GetErrorHandlerStatus = status
End Function