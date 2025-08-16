Attribute VB_Name = "BatchCommentAddin"
' ===============================================================================
' 模块: BatchCommentAddin
' 描述: Excel批量批注插件主模块
' 版本: 1.0.0
' 作者: OneDay Team
' 创建: 2024-01-01
' ===============================================================================

Option Explicit

' 常量定义
Public Const ADDIN_NAME As String = "批量批注助手"
Public Const ADDIN_VERSION As String = "1.0.0"
Public Const ADDIN_AUTHOR As String = "OneDay Team"
Public Const CONFIG_PATH As String = "\BatchCommentAddin\"

' 全局变量
Public g_IsInitialized As Boolean
Public g_ConfigPath As String
Public g_LogPath As String

' ===============================================================================
' 插件生命周期管理
' ===============================================================================

' 插件初始化 - Excel启动时自动调用
Public Sub Auto_Open()
    On Error GoTo ErrorHandler
    
    Call InitializeAddin
    Call CreateRibbonInterface
    Call LogOperation("STARTUP", "插件启动成功 v" & ADDIN_VERSION)
    
    g_IsInitialized = True
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "BatchCommentAddin", "Auto_Open")
End Sub

' 插件卸载 - Excel关闭时自动调用
Public Sub Auto_Close()
    On Error Resume Next
    
    Call SaveUserSettings
    Call CleanupResources
    Call LogOperation("SHUTDOWN", "插件正常关闭")
    
    g_IsInitialized = False
End Sub

' 初始化插件
Private Sub InitializeAddin()
    ' 设置路径
    g_ConfigPath = Environ("APPDATA") & CONFIG_PATH
    g_LogPath = g_ConfigPath & "logs\"
    
    ' 创建必要的目录
    Call CreateDirectoryStructure
    
    ' 初始化组件
    Call InitializeTemplateManager
    Call LoadUserSettings
    
    ' 设置应用程序属性
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' 创建目录结构
Private Sub CreateDirectoryStructure()
    On Error Resume Next
    
    If Dir(g_ConfigPath, vbDirectory) = "" Then MkDir g_ConfigPath
    If Dir(g_LogPath, vbDirectory) = "" Then MkDir g_LogPath
    If Dir(g_ConfigPath & "templates\", vbDirectory) = "" Then MkDir g_ConfigPath & "templates\"
    If Dir(g_ConfigPath & "temp\", vbDirectory) = "" Then MkDir g_ConfigPath & "temp\"
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 用户界面管理
' ===============================================================================

' 创建功能区界面
Private Sub CreateRibbonInterface()
    On Error Resume Next
    
    ' 添加到快速访问工具栏
    Dim qat As CommandBar
    Set qat = Application.CommandBars("Quick Access Toolbar")
    
    Dim btn As CommandBarButton
    Set btn = qat.Controls.Add(Type:=msoControlButton)
    
    With btn
        .Caption = "批量批注"
        .OnAction = "ShowBatchCommentDialog"
        .FaceId = 1695
        .Tag = "BatchComment_QAT"
        .TooltipText = "打开批量批注工具"
    End With
    
    ' 添加到开发工具菜单
    Call AddToDevMenu
    
    On Error GoTo 0
End Sub

' 添加到开发工具菜单
Private Sub AddToDevMenu()
    On Error Resume Next
    
    Dim devMenu As CommandBarPopup
    Set devMenu = Application.CommandBars("Worksheet Menu Bar").FindControl(Tag:="Developer")
    
    If Not devMenu Is Nothing Then
        Dim btn As CommandBarButton
        Set btn = devMenu.Controls.Add(Type:=msoControlButton)
        
        With btn
            .Caption = "批量批注工具(&B)"
            .OnAction = "ShowBatchCommentDialog"
            .FaceId = 1695
            .Tag = "BatchComment_Dev"
        End With
    End If
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 主要功能入口
' ===============================================================================

' 显示批量批注对话框
Public Sub ShowBatchCommentDialog()
    On Error GoTo ErrorHandler
    
    If Not g_IsInitialized Then
        Call InitializeAddin
    End If
    
    ' 检查Excel状态
    If Not IsExcelReady() Then
        Call ShowWarning("Excel当前状态不允许执行此操作，请稍后再试。")
        Exit Sub
    End If
    
    ' 显示主对话框
    Load BatchCommentForm
    BatchCommentForm.Show vbModal
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "BatchCommentAddin", "ShowBatchCommentDialog")
End Sub

' 显示模板管理器
Public Sub ShowTemplateManager()
    On Error GoTo ErrorHandler
    
    Load TemplateManagerForm
    TemplateManagerForm.Show vbModal
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "BatchCommentAddin", "ShowTemplateManager")
End Sub

' ===============================================================================
' 工具函数
' ===============================================================================

' 检查Excel是否准备就绪
Private Function IsExcelReady() As Boolean
    On Error GoTo ErrorHandler
    
    IsExcelReady = False
    
    ' 检查是否有活动工作簿
    If Application.Workbooks.Count = 0 Then Exit Function
    
    ' 检查是否有活动工作表
    If ActiveSheet Is Nothing Then Exit Function
    
    ' 检查是否在编辑模式
    If Application.Interactive = False Then Exit Function
    
    IsExcelReady = True
    Exit Function
    
ErrorHandler:
    IsExcelReady = False
End Function

' 加载用户设置
Private Sub LoadUserSettings()
    On Error Resume Next
    
    Dim settingsFile As String
    settingsFile = g_ConfigPath & "settings.ini"
    
    If Dir(settingsFile) <> "" Then
        ' 这里可以添加INI文件读取逻辑
        ' 暂时使用默认设置
    End If
    
    On Error GoTo 0
End Sub

' 保存用户设置
Private Sub SaveUserSettings()
    On Error Resume Next
    
    Dim settingsFile As String
    settingsFile = g_ConfigPath & "settings.ini"
    
    ' 这里可以添加INI文件写入逻辑
    
    On Error GoTo 0
End Sub

' 清理资源
Private Sub CleanupResources()
    On Error Resume Next
    
    ' 清理临时文件
    Call CleanupTempFiles
    
    ' 移除菜单项
    Application.CommandBars("Quick Access Toolbar").Controls("批量批注").Delete
    
    ' 卸载窗体
    Unload BatchCommentForm
    Unload TemplateManagerForm
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 版本信息和帮助
' ===============================================================================

' 显示关于对话框
Public Sub ShowAboutDialog()
    Dim aboutText As String
    aboutText = ADDIN_NAME & " v" & ADDIN_VERSION & vbCrLf & vbCrLf
    aboutText = aboutText & "作者: " & ADDIN_AUTHOR & vbCrLf
    aboutText = aboutText & "版权所有 © 2024" & vbCrLf & vbCrLf
    aboutText = aboutText & "一个专业的Excel批量批注工具，支持多种批注来源、" & vbCrLf
    aboutText = aboutText & "格式设置和模板管理功能。" & vbCrLf & vbCrLf
    aboutText = aboutText & "GitHub: https://github.com/yourusername/BatchCommentAddin"
    
    MsgBox aboutText, vbInformation, "关于 " & ADDIN_NAME
End Sub

' 获取版本信息
Public Function GetVersion() As String
    GetVersion = ADDIN_VERSION
End Function

' 检查更新
Public Sub CheckForUpdates()
    ' 这里可以添加在线更新检查逻辑
    MsgBox "当前版本: " & ADDIN_VERSION & vbCrLf & "您使用的是最新版本。", vbInformation, "检查更新"
End Sub