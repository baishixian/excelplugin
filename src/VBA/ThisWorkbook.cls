VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ===============================================================================
' 类: ThisWorkbook
' 描述: 工作簿事件处理
' 版本: 1.0.0
' 作者: OneDay Team
' 创建: 2024-01-01
' ===============================================================================

Option Explicit

' ===============================================================================
' 工作簿事件
' ===============================================================================

' 工作簿打开事件
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    
    ' 插件自动初始化在Auto_Open中处理
    ' 这里可以添加工作簿特定的初始化逻辑
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "ThisWorkbook", "Workbook_Open")
End Sub

' 工作簿关闭前事件
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' 清理资源
    Call CleanupResources
    
    On Error GoTo 0
End Sub

' 工作簿保存前事件
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    
    ' 记录保存操作
    Call LogOperation("WORKBOOK_SAVE", "工作簿保存: " & Me.Name)
    
    On Error GoTo 0
End Sub

' 工作表激活事件
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    
    ' 可以在这里添加工作表切换时的逻辑
    
    On Error GoTo 0
End Sub

' 工作表选择变化事件
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    
    ' 可以在这里添加选择变化时的逻辑
    ' 例如：更新批注工具的目标区域
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 加载项特定事件
' ===============================================================================

' 加载项安装事件
Private Sub Workbook_AddinInstall()
    On Error GoTo ErrorHandler
    
    Call LogOperation("ADDIN_INSTALL", "插件安装完成")
    
    ' 显示欢迎消息
    MsgBox "欢迎使用批量批注插件！" & vbCrLf & vbCrLf & _
           "您可以在快速访问工具栏或开发工具菜单中找到批量批注功能。" & vbCrLf & vbCrLf & _
           "如需帮助，请点击工具中的帮助按钮。", _
           vbInformation, ADDIN_NAME
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "ThisWorkbook", "Workbook_AddinInstall")
End Sub

' 加载项卸载事件
Private Sub Workbook_AddinUninstall()
    On Error Resume Next
    
    Call LogOperation("ADDIN_UNINSTALL", "插件卸载")
    
    ' 清理所有资源
    Call CleanupResources
    
    ' 显示卸载消息
    MsgBox "批量批注插件已成功卸载。" & vbCrLf & "感谢您的使用！", _
           vbInformation, ADDIN_NAME
    
    On Error GoTo 0
End Sub