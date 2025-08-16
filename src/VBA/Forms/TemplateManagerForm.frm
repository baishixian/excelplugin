VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateManagerForm 
   Caption         =   "模板管理器"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   OleObjectBlob   =   "TemplateManagerForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "TemplateManagerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ===============================================================================
' 窗体: TemplateManagerForm
' 描述: 批注模板管理界面
' 版本: 1.0.0
' ===============================================================================

Option Explicit

' 私有变量
Private m_selectedTemplate As String
Private m_isEditing As Boolean
Private m_originalTemplate As CommentTemplate

' 控件声明
Private WithEvents lstTemplates As MSForms.ListBox
Private WithEvents txtTemplateName As MSForms.TextBox
Private WithEvents cmbFontName As MSForms.ComboBox
Private WithEvents cmbFontSize As MSForms.ComboBox
Private WithEvents cmbFontColor As MSForms.ComboBox
Private WithEvents chkBold As MSForms.CheckBox
Private WithEvents chkItalic As MSForms.CheckBox
Private WithEvents txtDefaultText As MSForms.TextBox
Private WithEvents txtWidth As MSForms.TextBox
Private WithEvents txtHeight As MSForms.TextBox
Private WithEvents chkAutoSize As MSForms.CheckBox
Private WithEvents lblPreview As MSForms.Label
Private WithEvents btnNew As MSForms.CommandButton
Private WithEvents btnEdit As MSForms.CommandButton
Private WithEvents btnDelete As MSForms.CommandButton
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private WithEvents btnImport As MSForms.CommandButton
Private WithEvents btnExport As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

' ===============================================================================
' 窗体事件
' ===============================================================================

' 窗体初始化
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Call InitializeControls
    Call LoadTemplateList
    Call SetEditMode(False)
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etUserInterface, Err.Number, Err.Description, "TemplateManagerForm", "UserForm_Initialize")
End Sub

' ===============================================================================
' 控件初始化
' ===============================================================================

' 初始化控件
Private Sub InitializeControls()
    ' 设置窗体属性
    Me.Caption = "模板管理器 v" & GetVersion()
    Me.Width = 640
    Me.Height = 440
    
    ' 创建控件
    Call CreateControls
    Call SetControlProperties
    Call InitializeComboBoxes
End Sub

' 创建控件
Private Sub CreateControls()
    ' 模板列表
    Set lstTemplates = Me.Controls.Add("Forms.ListBox.1", "lstTemplates")
    
    ' 模板属性编辑
    Set txtTemplateName = Me.Controls.Add("Forms.TextBox.1", "txtTemplateName")
    Set cmbFontName = Me.Controls.Add("Forms.ComboBox.1", "cmbFontName")
    Set cmbFontSize = Me.Controls.Add("Forms.ComboBox.1", "cmbFontSize")
    Set cmbFontColor = Me.Controls.Add("Forms.ComboBox.1", "cmbFontColor")
    Set chkBold = Me.Controls.Add("Forms.CheckBox.1", "chkBold")
    Set chkItalic = Me.Controls.Add("Forms.CheckBox.1", "chkItalic")
    Set txtDefaultText = Me.Controls.Add("Forms.TextBox.1", "txtDefaultText")
    Set txtWidth = Me.Controls.Add("Forms.TextBox.1", "txtWidth")
    Set txtHeight = Me.Controls.Add("Forms.TextBox.1", "txtHeight")
    Set chkAutoSize = Me.Controls.Add("Forms.CheckBox.1", "chkAutoSize")
    
    ' 预览
    Set lblPreview = Me.Controls.Add("Forms.Label.1", "lblPreview")
    
    ' 按钮
    Set btnNew = Me.Controls.Add("Forms.CommandButton.1", "btnNew")
    Set btnEdit = Me.Controls.Add("Forms.CommandButton.1", "btnEdit")
    Set btnDelete = Me.Controls.Add("Forms.CommandButton.1", "btnDelete")
    Set btnSave = Me.Controls.Add("Forms.CommandButton.1", "btnSave")
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    Set btnImport = Me.Controls.Add("Forms.CommandButton.1", "btnImport")
    Set btnExport = Me.Controls.Add("Forms.CommandButton.1", "btnExport")
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
End Sub

' 设置控件属性
Private Sub SetControlProperties()
    ' 模板列表
    With lstTemplates
        .Left = 20: .Top = 30: .Width = 180: .Height = 300
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    ' 模板名称
    With txtTemplateName
        .Left = 220: .Top = 30: .Width = 200: .Height = 18
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    ' 字体设置
    With cmbFontName
        .Left = 220: .Top = 60: .Width = 120: .Height = 18
        .Style = fmStyleDropDownCombo
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With cmbFontSize
        .Left = 350: .Top = 60: .Width = 70: .Height = 18
        .Style = fmStyleDropDownCombo
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With cmbFontColor
        .Left = 430: .Top = 60: .Width = 80: .Height = 18
        .Style = fmStyleDropDownList
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With chkBold
        .Left = 220: .Top = 90: .Width = 60: .Height = 18
        .Caption = "粗体": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With chkItalic
        .Left = 290: .Top = 90: .Width = 60: .Height = 18
        .Caption = "斜体": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    ' 默认文本
    With txtDefaultText
        .Left = 220: .Top = 120: .Width = 290: .Height = 40
        .MultiLine = True: .ScrollBars = fmScrollBarsVertical
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    ' 大小设置
    With txtWidth
        .Left = 220: .Top = 170: .Width = 60: .Height = 18
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With txtHeight
        .Left = 290: .Top = 170: .Width = 60: .Height = 18
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With chkAutoSize
        .Left = 360: .Top = 170: .Width = 80: .Height = 18
        .Caption = "自动调整": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    ' 预览
    With lblPreview
        .Left = 220: .Top = 200: .Width = 290: .Height = 80
        .BorderStyle = fmBorderStyleSingle
        .BackColor = &HF0F0F0
        .Font.Name = "微软雅黑": .Font.Size = 8
        .Caption = "预览区域"
    End With
    
    ' 按钮组
    With btnNew
        .Left = 20: .Top = 340: .Width = 60: .Height = 25
        .Caption = "新建": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnEdit
        .Left = 90: .Top = 340: .Width = 60: .Height = 25
        .Caption = "编辑": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnDelete
        .Left = 160: .Top = 340: .Width = 60: .Height = 25
        .Caption = "删除": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnSave
        .Left = 220: .Top = 340: .Width = 60: .Height = 25
        .Caption = "保存": .Font.Name = "微软雅黑": .Font.Size = 9
        .Enabled = False
    End With
    
    With btnCancel
        .Left = 290: .Top = 340: .Width = 60: .Height = 25
        .Caption = "取消": .Font.Name = "微软雅黑": .Font.Size = 9
        .Enabled = False
    End With
    
    With btnImport
        .Left = 370: .Top = 340: .Width = 60: .Height = 25
        .Caption = "导入": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnExport
        .Left = 440: .Top = 340: .Width = 60: .Height = 25
        .Caption = "导出": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnClose
        .Left = 510: .Top = 340: .Width = 60: .Height = 25
        .Caption = "关闭": .Font.Name = "微软雅黑": .Font.Size = 9
        .Cancel = True
    End With
End Sub

' 初始化下拉框
Private Sub InitializeComboBoxes()
    ' 字体列表
    cmbFontName.Clear
    cmbFontName.AddItem "微软雅黑"
    cmbFontName.AddItem "宋体"
    cmbFontName.AddItem "Arial"
    cmbFontName.AddItem "Times New Roman"
    cmbFontName.AddItem "Calibri"
    cmbFontName.AddItem "Courier New"
    cmbFontName.AddItem "Verdana"
    
    ' 字体大小
    cmbFontSize.Clear
    Dim sizes As Variant
    sizes = Array("8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32")
    
    Dim i As Long
    For i = 0 To UBound(sizes)
        cmbFontSize.AddItem sizes(i)
    Next i
    
    ' 颜色列表
    cmbFontColor.Clear
    cmbFontColor.AddItem "黑色"
    cmbFontColor.AddItem "红色"
    cmbFontColor.AddItem "蓝色"
    cmbFontColor.AddItem "绿色"
    cmbFontColor.AddItem "紫色"
    cmbFontColor.AddItem "橙色"
    cmbFontColor.AddItem "棕色"
    cmbFontColor.AddItem "灰色"
End Sub

' ===============================================================================
' 控件事件处理
' ===============================================================================

' 模板列表选择变化
Private Sub lstTemplates_Click()
    If lstTemplates.ListIndex >= 0 And Not m_isEditing Then
        m_selectedTemplate = lstTemplates.Value
        Call LoadTemplateDetails(m_selectedTemplate)
        Call UpdatePreview
    End If
End Sub

' 新建模板
Private Sub btnNew_Click()
    On Error GoTo ErrorHandler
    
    Call ClearTemplateDetails
    Call SetEditMode(True)
    txtTemplateName.SetFocus
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etTemplateError, Err.Number, Err.Description, "TemplateManagerForm", "btnNew_Click")
End Sub

' 编辑模板
Private Sub btnEdit_Click()
    On Error GoTo ErrorHandler
    
    If lstTemplates.ListIndex < 0 Then
        Call ShowWarning("请先选择要编辑的模板！")
        Exit Sub
    End If
    
    ' 保存原始模板信息
    m_originalTemplate = GetTemplate(m_selectedTemplate)
    
    Call SetEditMode(True)
    txtTemplateName.SetFocus
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etTemplateError, Err.Number, Err.Description, "TemplateManagerForm", "btnEdit_Click")
End Sub

' 删除模板
Private Sub btnDelete_Click()
    On Error GoTo ErrorHandler
    
    If lstTemplates.ListIndex < 0 Then
        Call ShowWarning("请先选择要删除的模板！")
        Exit Sub
    End If
    
    If ShowConfirm("确定要删除模板 '" & m_selectedTemplate & "' 吗？") Then
        If DeleteTemplate(m_selectedTemplate) Then
            Call LoadTemplateList
            Call ClearTemplateDetails
            Call ShowInfo("模板删除成功！")
        Else
            Call ShowError("删除模板失败！")
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etTemplateError, Err.Number, Err.Description, "TemplateManagerForm", "btnDelete_Click")
End Sub

' 保存模板
Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
    
    If Not ValidateTemplateInput() Then Exit Sub
    
    Dim newTemplate As CommentTemplate
    newTemplate = CreateTemplateFromInput()
    
    Dim success As Boolean
    If m_isEditing And m_selectedTemplate <> "" Then
        ' 更新现有模板
        success = UpdateTemplate(m_selectedTemplate, newTemplate)
    Else
        ' 添加新模板
        success = AddTemplate(newTemplate.Name, newTemplate.FontName, newTemplate.FontSize, _
                            newTemplate.FontColor, newTemplate.IsBold, newTemplate.IsItalic, _
                            newTemplate.DefaultText, newTemplate.BackgroundColor, _
                            newTemplate.Width, newTemplate.Height, newTemplate.IsAutoSize)
    End If
    
    If success Then
        Call LoadTemplateList
        Call SetEditMode(False)
        Call ShowInfo("模板保存成功！")
        
        ' 选中新保存的模板
        Dim i As Long
        For i = 0 To lstTemplates.ListCount - 1
            If lstTemplates.List(i) = newTemplate.Name Then
                lstTemplates.ListIndex = i
                Exit For
            End If
        Next i
    Else
        Call ShowError("保存模板失败！")
    End If
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etTemplateError, Err.Number, Err.Description, "TemplateManagerForm", "btnSave_Click")
End Sub

' 取消编辑
Private Sub btnCancel_Click()
    If m_isEditing Then
        Call SetEditMode(False)
        
        If m_selectedTemplate <> "" Then
            Call LoadTemplateDetails(m_selectedTemplate)
        Else
            Call ClearTemplateDetails
        End If
    End If
End Sub

' 导入模板
Private Sub btnImport_Click()
    On Error GoTo ErrorHandler
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "选择模板文件"
        .Filters.Clear
        .Filters.Add "模板文件", "*.bct"
        .Filters.Add "所有文件", "*.*"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            If ImportTemplate(.SelectedItems(1)) Then
                Call LoadTemplateList
                Call ShowInfo("模板导入成功！")
            Else
                Call ShowError("导入模板失败！")
            End If
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etFileAccess, Err.Number, Err.Description, "TemplateManagerForm", "btnImport_Click")
End Sub

' 导出模板
Private Sub btnExport_Click()
    On Error GoTo ErrorHandler
    
    If lstTemplates.ListIndex < 0 Then
        Call ShowWarning("请先选择要导出的模板！")
        Exit Sub
    End If
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    With fd
        .Title = "保存模板文件"
        .FilterIndex = 1
        .InitialFileName = m_selectedTemplate & ".bct"
        
        If .Show = -1 Then
            If ExportTemplate(m_selectedTemplate, .SelectedItems(1)) Then
                Call ShowInfo("模板导出成功！")
            Else
                Call ShowError("导出模板失败！")
            End If
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etFileAccess, Err.Number, Err.Description, "TemplateManagerForm", "btnExport_Click")
End Sub

' 关闭窗体
Private Sub btnClose_Click()
    Unload Me
End Sub

' ===============================================================================
' 业务逻辑
' ===============================================================================

' 加载模板列表
Private Sub LoadTemplateList()
    On Error Resume Next
    
    lstTemplates.Clear
    
    Dim templateNames() As String
    templateNames = GetTemplateNames()
    
    Dim i As Long
    For i = 0 To UBound(templateNames)
        lstTemplates.AddItem templateNames(i)
    Next i
    
    On Error GoTo 0
End Sub

' 加载模板详细信息
Private Sub LoadTemplateDetails(templateName As String)
    On Error Resume Next
    
    Dim template As CommentTemplate
    template = GetTemplate(templateName)
    
    With template
        txtTemplateName.Value = .Name
        cmbFontName.Value = .FontName
        cmbFontSize.Value = CStr(.FontSize)
        chkBold.Value = .IsBold
        chkItalic.Value = .IsItalic
        txtDefaultText.Value = .DefaultText
        txtWidth.Value = CStr(.Width)
        txtHeight.Value = CStr(.Height)
        chkAutoSize.Value = .IsAutoSize
        
        ' 设置颜色
        Select Case .FontColor
            Case 1: cmbFontColor.Value = "黑色"
            Case 3: cmbFontColor.Value = "红色"
            Case 5: cmbFontColor.Value = "蓝色"
            Case 10: cmbFontColor.Value = "绿色"
            Case 13: cmbFontColor.Value = "紫色"
            Case 46: cmbFontColor.Value = "橙色"
            Case Else: cmbFontColor.Value = "黑色"
        End Select
    End With
    
    Call UpdatePreview
    
    On Error GoTo 0
End Sub

' 清空模板详细信息
Private Sub ClearTemplateDetails()
    txtTemplateName.Value = ""
    cmbFontName.Value = "微软雅黑"
    cmbFontSize.Value = "9"
    cmbFontColor.Value = "黑色"
    chkBold.Value = False
    chkItalic.Value = False
    txtDefaultText.Value = ""
    txtWidth.Value = "200"
    txtHeight.Value = "100"
    chkAutoSize.Value = True
    
    Call UpdatePreview
End Sub

' 设置编辑模式
Private Sub SetEditMode(isEditing As Boolean)
    m_isEditing = isEditing
    
    ' 控制编辑控件的启用状态
    txtTemplateName.Enabled = isEditing
    cmbFontName.Enabled = isEditing
    cmbFontSize.Enabled = isEditing
    cmbFontColor.Enabled = isEditing
    chkBold.Enabled = isEditing
    chkItalic.Enabled = isEditing
    txtDefaultText.Enabled = isEditing
    txtWidth.Enabled = isEditing
    txtHeight.Enabled = isEditing
    chkAutoSize.Enabled = isEditing
    
    ' 控制按钮的启用状态
    btnNew.Enabled = Not isEditing
    btnEdit.Enabled = Not isEditing And lstTemplates.ListIndex >= 0
    btnDelete.Enabled = Not isEditing And lstTemplates.ListIndex >= 0
    btnSave.Enabled = isEditing
    btnCancel.Enabled = isEditing
    btnImport.Enabled = Not isEditing
    btnExport.Enabled = Not isEditing And lstTemplates.ListIndex >= 0
    
    lstTemplates.Enabled = Not isEditing
End Sub

' 验证模板输入
Private Function ValidateTemplateInput() As Boolean
    ValidateTemplateInput = False
    
    ' 验证模板名称
    If SafeTrim(txtTemplateName.Value) = "" Then
        Call ShowWarning("请输入模板名称！")
        txtTemplateName.SetFocus
        Exit Function
    End If
    
    ' 检查模板名称是否重复（新建时）
    If Not m_isEditing Or txtTemplateName.Value <> m_selectedTemplate Then
        If TemplateExists(txtTemplateName.Value) Then
            Call ShowWarning("模板名称已存在，请使用其他名称！")
            txtTemplateName.SetFocus
            Exit Function
        End If
    End If
    
    ' 验证字体设置
    If SafeTrim(cmbFontName.Value) = "" Then
        Call ShowWarning("请选择字体名称！")
        cmbFontName.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(cmbFontSize.Value) Or CLng(cmbFontSize.Value) < 6 Or CLng(cmbFontSize.Value) > 72 Then
        Call ShowWarning("字体大小必须在6-72之间！")
        cmbFontSize.SetFocus
        Exit Function
    End If
    
    ' 验证大小设置（如果不是自动调整）
    If Not chkAutoSize.Value Then
        If Not IsNumeric(txtWidth.Value) Or CLng(txtWidth.Value) < 50 Or CLng(txtWidth.Value) > 500 Then
            Call ShowWarning("宽度必须在50-500之间！")
            txtWidth.SetFocus
            Exit Function
        End If
        
        If Not IsNumeric(txtHeight.Value) Or CLng(txtHeight.Value) < 30 Or CLng(txtHeight.Value) > 300 Then
            Call ShowWarning("高度必须在30-300之间！")
            txtHeight.SetFocus
            Exit Function
        End If
    End If
    
    ValidateTemplateInput = True
End Function

' 从输入创建模板对象
Private Function CreateTemplateFromInput() As CommentTemplate
    Dim template As CommentTemplate
    
    With template
        .Name = SafeTrim(txtTemplateName.Value)
        .FontName = cmbFontName.Value
        .FontSize = CLng(cmbFontSize.Value)
        .FontColor = GetColorIndexFromName(cmbFontColor.Value)
        .IsBold = chkBold.Value
        .IsItalic = chkItalic.Value
        .DefaultText = txtDefaultText.Value
        .BackgroundColor = &HE0E0E0 ' 默认背景色
        .Width = IIf(IsNumeric(txtWidth.Value), CLng(txtWidth.Value), 200)
        .Height = IIf(IsNumeric(txtHeight.Value), CLng(txtHeight.Value), 100)
        .IsAutoSize = chkAutoSize.Value
    End With
    
    CreateTemplateFromInput = template
End Function

' 从颜色名称获取颜色索引
Private Function GetColorIndexFromName(colorName As String) As Long
    Select Case colorName
        Case "黑色": GetColorIndexFromName = 1
        Case "红色": GetColorIndexFromName = 3
        Case "蓝色": GetColorIndexFromName = 5
        Case "绿色": GetColorIndexFromName = 10
        Case "紫色": GetColorIndexFromName = 13
        Case "橙色": GetColorIndexFromName = 46
        Case "棕色": GetColorIndexFromName = 53
        Case "灰色": GetColorIndexFromName = 15
        Case Else: GetColorIndexFromName = 1
    End Select
End Function

' 更新预览
Private Sub UpdatePreview()
    On Error Resume Next
    
    Dim previewText As String
    previewText = "模板预览" & vbCrLf
    previewText = previewText & "字体: " & cmbFontName.Value & " " & cmbFontSize.Value & "pt" & vbCrLf
    previewText = previewText & "颜色: " & cmbFontColor.Value & vbCrLf
    previewText = previewText & "样式: "
    
    If chkBold.Value Then previewText = previewText & "粗体 "
    If chkItalic.Value Then previewText = previewText & "斜体 "
    
    previewText = previewText & vbCrLf & "大小: "
    If chkAutoSize.Value Then
        previewText = previewText & "自动调整"
    Else
        previewText = previewText & txtWidth.Value & "x" & txtHeight.Value
    End If
    
    If SafeTrim(txtDefaultText.Value) <> "" Then
        previewText = previewText & vbCrLf & "默认文本: " & Left(txtDefaultText.Value, 20) & "..."
    End If
    
    lblPreview.Caption = previewText
    
    ' 设置预览样式
    With lblPreview.Font
        .Name = cmbFontName.Value
        .Size = IIf(IsNumeric(cmbFontSize.Value), CLng(cmbFontSize.Value), 9)
        .Bold = chkBold.Value
        .Italic = chkItalic.Value
    End With
    
    On Error GoTo 0
End Sub

' 字体设置变化时更新预览
Private Sub cmbFontName_Change()
    Call UpdatePreview
End Sub

Private Sub cmbFontSize_Change()
    Call UpdatePreview
End Sub

Private Sub cmbFontColor_Change()
    Call UpdatePreview
End Sub

Private Sub chkBold_Click()
    Call UpdatePreview
End Sub

Private Sub chkItalic_Click()
    Call UpdatePreview
End Sub

Private Sub txtDefaultText_Change()
    Call UpdatePreview
End Sub

Private Sub chkAutoSize_Click()
    txtWidth.Enabled = Not chkAutoSize.Value
    txtHeight.Enabled = Not chkAutoSize.Value
    Call UpdatePreview
End Sub