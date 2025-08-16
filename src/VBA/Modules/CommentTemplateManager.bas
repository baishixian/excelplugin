Attribute VB_Name = "CommentTemplateManager"
' ===============================================================================
' 模块: CommentTemplateManager
' 描述: 批注模板管理器
' 版本: 1.0.0
' 作者: OneDay Team
' 创建: 2024-01-01
' ===============================================================================

Option Explicit

' 模板结构定义
Public Type CommentTemplate
    Name As String
    FontName As String
    FontSize As Integer
    FontColor As Long
    IsBold As Boolean
    IsItalic As Boolean
    DefaultText As String
    BackgroundColor As Long
    Width As Integer
    Height As Integer
    IsAutoSize As Boolean
End Type

' 全局变量
Private m_Templates() As CommentTemplate
Private m_TemplateCount As Long
Private m_IsInitialized As Boolean

' ===============================================================================
' 模板管理器初始化
' ===============================================================================

' 初始化模板管理器
Public Sub InitializeTemplateManager()
    On Error GoTo ErrorHandler
    
    If m_IsInitialized Then Exit Sub
    
    ' 重置模板数组
    ReDim m_Templates(0 To 9) ' 预分配10个模板空间
    m_TemplateCount = 0
    
    ' 加载默认模板
    Call LoadDefaultTemplates
    
    ' 加载用户自定义模板
    Call LoadUserTemplates
    
    m_IsInitialized = True
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "CommentTemplateManager", "InitializeTemplateManager")
End Sub

' 加载默认模板
Private Sub LoadDefaultTemplates()
    On Error Resume Next
    
    ' 默认模板
    Call AddTemplate("默认", "微软雅黑", 9, 1, False, False, "", &HE0E0E0, 200, 100, True)
    
    ' 重要提示模板
    Call AddTemplate("重要提示", "微软雅黑", 10, 3, True, False, "重要：", &HFFCCCC, 250, 120, True)
    
    ' 说明文档模板
    Call AddTemplate("说明文档", "宋体", 9, 1, False, False, "说明：", &HE0FFE0, 300, 150, True)
    
    ' 警告模板
    Call AddTemplate("警告", "微软雅黑", 10, 46, True, True, "警告：", &HCCFFFF, 280, 130, True)
    
    ' 数据来源模板
    Call AddTemplate("数据来源", "Calibri", 8, 5, False, True, "数据来源：", &HF0F0F0, 220, 80, True)
    
    On Error GoTo 0
End Sub

' 加载用户自定义模板
Private Sub LoadUserTemplates()
    On Error Resume Next
    
    Dim templateFile As String
    templateFile = g_ConfigPath & "templates\user_templates.txt"
    
    If Dir(templateFile) <> "" Then
        ' 这里可以添加从文件加载用户模板的逻辑
        ' 暂时跳过
    End If
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 模板操作
' ===============================================================================

' 添加模板
Public Function AddTemplate(templateName As String, fontName As String, fontSize As Integer, _
                           fontColor As Long, isBold As Boolean, isItalic As Boolean, _
                           defaultText As String, backgroundColor As Long, _
                           width As Integer, height As Integer, isAutoSize As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    AddTemplate = False
    
    ' 检查模板名称是否已存在
    If GetTemplateIndex(templateName) >= 0 Then
        Exit Function ' 模板已存在
    End If
    
    ' 扩展数组如果需要
    If m_TemplateCount >= UBound(m_Templates) Then
        ReDim Preserve m_Templates(0 To UBound(m_Templates) + 10)
    End If
    
    ' 添加新模板
    With m_Templates(m_TemplateCount)
        .Name = templateName
        .FontName = fontName
        .FontSize = fontSize
        .FontColor = fontColor
        .IsBold = isBold
        .IsItalic = isItalic
        .DefaultText = defaultText
        .BackgroundColor = backgroundColor
        .Width = width
        .Height = height
        .IsAutoSize = isAutoSize
    End With
    
    m_TemplateCount = m_TemplateCount + 1
    AddTemplate = True
    
    Exit Function
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "CommentTemplateManager", "AddTemplate")
    AddTemplate = False
End Function

' 删除模板
Public Function DeleteTemplate(templateName As String) As Boolean
    On Error GoTo ErrorHandler
    
    DeleteTemplate = False
    
    Dim index As Long
    index = GetTemplateIndex(templateName)
    
    If index < 0 Then Exit Function ' 模板不存在
    
    ' 移动数组元素
    Dim i As Long
    For i = index To m_TemplateCount - 2
        m_Templates(i) = m_Templates(i + 1)
    Next i
    
    m_TemplateCount = m_TemplateCount - 1
    DeleteTemplate = True
    
    ' 保存到文件
    Call SaveUserTemplates
    
    Exit Function
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "CommentTemplateManager", "DeleteTemplate")
    DeleteTemplate = False
End Function

' 更新模板
Public Function UpdateTemplate(oldName As String, newTemplate As CommentTemplate) As Boolean
    On Error GoTo ErrorHandler
    
    UpdateTemplate = False
    
    Dim index As Long
    index = GetTemplateIndex(oldName)
    
    If index < 0 Then Exit Function ' 模板不存在
    
    ' 更新模板
    m_Templates(index) = newTemplate
    UpdateTemplate = True
    
    ' 保存到文件
    Call SaveUserTemplates
    
    Exit Function
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "CommentTemplateManager", "UpdateTemplate")
    UpdateTemplate = False
End Function

' ===============================================================================
' 模板查询
' ===============================================================================

' 获取模板
Public Function GetTemplate(templateName As String) As CommentTemplate
    On Error GoTo ErrorHandler
    
    Dim index As Long
    index = GetTemplateIndex(templateName)
    
    If index >= 0 Then
        GetTemplate = m_Templates(index)
    Else
        ' 返回默认模板
        GetTemplate = m_Templates(0)
    End If
    
    Exit Function
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "CommentTemplateManager", "GetTemplate")
    ' 返回空模板
    Dim emptyTemplate As CommentTemplate
    GetTemplate = emptyTemplate
End Function

' 获取模板索引
Private Function GetTemplateIndex(templateName As String) As Long
    On Error Resume Next
    
    GetTemplateIndex = -1
    
    Dim i As Long
    For i = 0 To m_TemplateCount - 1
        If m_Templates(i).Name = templateName Then
            GetTemplateIndex = i
            Exit Function
        End If
    Next i
    
    On Error GoTo 0
End Function

' 获取所有模板名称
Public Function GetTemplateNames() As String()
    On Error GoTo ErrorHandler
    
    If m_TemplateCount = 0 Then
        Dim emptyArray(0) As String
        GetTemplateNames = emptyArray
        Exit Function
    End If
    
    Dim names() As String
    ReDim names(0 To m_TemplateCount - 1)
    
    Dim i As Long
    For i = 0 To m_TemplateCount - 1
        names(i) = m_Templates(i).Name
    Next i
    
    GetTemplateNames = names
    Exit Function
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "CommentTemplateManager", "GetTemplateNames")
    Dim errorArray(0) As String
    GetTemplateNames = errorArray
End Function

' 获取模板数量
Public Function GetTemplateCount() As Long
    GetTemplateCount = m_TemplateCount
End Function

' 模板是否存在
Public Function TemplateExists(templateName As String) As Boolean
    TemplateExists = (GetTemplateIndex(templateName) >= 0)
End Function

' ===============================================================================
' 模板持久化
' ===============================================================================

' 保存用户模板到文件
Public Sub SaveUserTemplates()
    On Error GoTo ErrorHandler
    
    Dim templateFile As String
    templateFile = g_ConfigPath & "templates\user_templates.txt"
    
    ' 创建模板目录
    If Dir(g_ConfigPath & "templates\", vbDirectory) = "" Then
        MkDir g_ConfigPath & "templates\"
    End If
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open templateFile For Output As #fileNum
    
    ' 写入模板数据（跳过前5个默认模板）
    Dim i As Long
    For i = 5 To m_TemplateCount - 1
        With m_Templates(i)
            Print #fileNum, .Name & "|" & .FontName & "|" & .FontSize & "|" & .FontColor & "|" & _
                           .IsBold & "|" & .IsItalic & "|" & .DefaultText & "|" & .BackgroundColor & "|" & _
                           .Width & "|" & .Height & "|" & .IsAutoSize
        End With
    Next i
    
    Close #fileNum
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call HandleError(etFileAccess, Err.Number, Err.Description, "CommentTemplateManager", "SaveUserTemplates")
End Sub

' 从文件加载用户模板
Public Sub LoadUserTemplatesFromFile()
    On Error GoTo ErrorHandler
    
    Dim templateFile As String
    templateFile = g_ConfigPath & "templates\user_templates.txt"
    
    If Dir(templateFile) = "" Then Exit Sub
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open templateFile For Input As #fileNum
    
    Dim line As String
    Dim parts() As String
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        
        If Trim(line) <> "" Then
            parts = Split(line, "|")
            
            If UBound(parts) >= 10 Then
                Call AddTemplate(parts(0), parts(1), CInt(parts(2)), CLng(parts(3)), _
                               CBool(parts(4)), CBool(parts(5)), parts(6), CLng(parts(7)), _
                               CInt(parts(8)), CInt(parts(9)), CBool(parts(10)))
            End If
        End If
    Loop
    
    Close #fileNum
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call HandleError(etFileAccess, Err.Number, Err.Description, "CommentTemplateManager", "LoadUserTemplatesFromFile")
End Sub

' ===============================================================================
' 模板导入导出
' ===============================================================================

' 导出模板到文件
Public Function ExportTemplate(templateName As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ExportTemplate = False
    
    Dim template As CommentTemplate
    template = GetTemplate(templateName)
    
    If template.Name = "" Then Exit Function
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    
    ' 写入模板信息
    Print #fileNum, "[BatchCommentTemplate]"
    Print #fileNum, "Name=" & template.Name
    Print #fileNum, "FontName=" & template.FontName
    Print #fileNum, "FontSize=" & template.FontSize
    Print #fileNum, "FontColor=" & template.FontColor
    Print #fileNum, "IsBold=" & template.IsBold
    Print #fileNum, "IsItalic=" & template.IsItalic
    Print #fileNum, "DefaultText=" & template.DefaultText
    Print #fileNum, "BackgroundColor=" & template.BackgroundColor
    Print #fileNum, "Width=" & template.Width
    Print #fileNum, "Height=" & template.Height
    Print #fileNum, "IsAutoSize=" & template.IsAutoSize
    
    Close #fileNum
    ExportTemplate = True
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call HandleError(etFileAccess, Err.Number, Err.Description, "CommentTemplateManager", "ExportTemplate")
    ExportTemplate = False
End Function

' 从文件导入模板
Public Function ImportTemplate(filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ImportTemplate = False
    
    If Dir(filePath) = "" Then Exit Function
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    
    Dim line As String
    Dim template As CommentTemplate
    Dim isValidTemplate As Boolean
    
    ' 读取模板信息
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        
        If InStr(line, "[BatchCommentTemplate]") > 0 Then
            isValidTemplate = True
        ElseIf InStr(line, "Name=") = 1 Then
            template.Name = Mid(line, 6)
        ElseIf InStr(line, "FontName=") = 1 Then
            template.FontName = Mid(line, 10)
        ElseIf InStr(line, "FontSize=") = 1 Then
            template.FontSize = CInt(Mid(line, 10))
        ElseIf InStr(line, "FontColor=") = 1 Then
            template.FontColor = CLng(Mid(line, 11))
        ElseIf InStr(line, "IsBold=") = 1 Then
            template.IsBold = CBool(Mid(line, 8))
        ElseIf InStr(line, "IsItalic=") = 1 Then
            template.IsItalic = CBool(Mid(line, 10))
        ElseIf InStr(line, "DefaultText=") = 1 Then
            template.DefaultText = Mid(line, 13)
        ElseIf InStr(line, "BackgroundColor=") = 1 Then
            template.BackgroundColor = CLng(Mid(line, 17))
        ElseIf InStr(line, "Width=") = 1 Then
            template.Width = CInt(Mid(line, 7))
        ElseIf InStr(line, "Height=") = 1 Then
            template.Height = CInt(Mid(line, 8))
        ElseIf InStr(line, "IsAutoSize=") = 1 Then
            template.IsAutoSize = CBool(Mid(line, 12))
        End If
    Loop
    
    Close #fileNum
    
    ' 验证并添加模板
    If isValidTemplate And template.Name <> "" Then
        ImportTemplate = AddTemplate(template.Name, template.FontName, template.FontSize, _
                                   template.FontColor, template.IsBold, template.IsItalic, _
                                   template.DefaultText, template.BackgroundColor, _
                                   template.Width, template.Height, template.IsAutoSize)
        
        If ImportTemplate Then
            Call SaveUserTemplates
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call HandleError(etFileAccess, Err.Number, Err.Description, "CommentTemplateManager", "ImportTemplate")
    ImportTemplate = False
End Function

' ===============================================================================
' 模板应用
' ===============================================================================

' 应用模板到批注
Public Sub ApplyTemplateToComment(comment As comment, templateName As String)
    On Error Resume Next
    
    Dim template As CommentTemplate
    template = GetTemplate(templateName)
    
    If template.Name = "" Then Exit Sub
    
    ' 应用字体设置
    With comment.Shape.TextFrame.Characters.Font
        .Name = template.FontName
        .Size = template.FontSize
        .ColorIndex = template.FontColor
        .Bold = template.IsBold
        .Italic = template.IsItalic
    End With
    
    ' 应用大小设置
    With comment.Shape
        If template.IsAutoSize Then
            .TextFrame.AutoSize = True
        Else
            .Width = template.Width
            .Height = template.Height
            .TextFrame.AutoSize = False
        End If
        
        ' 设置背景色
        .Fill.ForeColor.RGB = template.BackgroundColor
    End With
    
    On Error GoTo 0
End Sub

' 获取模板预览信息
Public Function GetTemplatePreview(templateName As String) As String
    On Error Resume Next
    
    Dim template As CommentTemplate
    template = GetTemplate(templateName)
    
    If template.Name = "" Then
        GetTemplatePreview = "模板不存在"
        Exit Function
    End If
    
    Dim preview As String
    preview = "模板: " & template.Name & vbCrLf
    preview = preview & "字体: " & template.FontName & " " & template.FontSize & "pt" & vbCrLf
    preview = preview & "样式: "
    
    If template.IsBold Then preview = preview & "粗体 "
    If template.IsItalic Then preview = preview & "斜体 "
    
    preview = preview & vbCrLf & "大小: "
    If template.IsAutoSize Then
        preview = preview & "自动调整"
    Else
        preview = preview & template.Width & "x" & template.Height
    End If
    
    If template.DefaultText <> "" Then
        preview = preview & vbCrLf & "默认文本: " & template.DefaultText
    End If
    
    GetTemplatePreview = preview
    
    On Error GoTo 0
End Function