VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BatchCommentForm 
   Caption         =   "批量批注工具"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   OleObjectBlob   =   "BatchCommentForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "BatchCommentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ===============================================================================
' 窗体: BatchCommentForm
' 描述: 批量批注主界面
' 版本: 1.0.0
' ===============================================================================

Option Explicit

' 私有变量
Private m_cancelled As Boolean
Private m_processing As Boolean
Private m_targetRange As Range
Private m_sourceData As Variant
Private m_currentTemplate As CommentTemplate

' 控件声明
Private WithEvents txtTargetRange As MSForms.TextBox
Private WithEvents btnSelectRange As MSForms.CommandButton
Private WithEvents optFromCell As MSForms.OptionButton
Private WithEvents optFromText As MSForms.OptionButton
Private WithEvents optFromFile As MSForms.OptionButton
Private WithEvents txtCommentSource As MSForms.TextBox
Private WithEvents btnSelectSource As MSForms.CommandButton
Private WithEvents txtFixedText As MSForms.TextBox
Private WithEvents cmbTemplate As MSForms.ComboBox
Private WithEvents cmbFontName As MSForms.ComboBox
Private WithEvents cmbFontSize As MSForms.ComboBox
Private WithEvents cmbFontColor As MSForms.ComboBox
Private WithEvents chkBold As MSForms.CheckBox
Private WithEvents chkItalic As MSForms.CheckBox
Private WithEvents progressBar As MSForms.Label
Private WithEvents lblProgress As MSForms.Label
Private WithEvents btnOK As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private WithEvents btnHelp As MSForms.CommandButton
Private WithEvents btnTemplateManager As MSForms.CommandButton

' ===============================================================================
' 窗体事件
' ===============================================================================

' 窗体初始化
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Call InitializeControls
    Call LoadTemplates
    Call LoadSettings
    Call SetDefaultValues
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "BatchCommentForm", "UserForm_Initialize")
End Sub

' 窗体激活
Private Sub UserForm_Activate()
    ' 检查是否有选中的区域
    If Not Selection Is Nothing And TypeName(Selection) = "Range" Then
        If txtTargetRange.Value = "" Then
            txtTargetRange.Value = Selection.Address
        End If
    End If
End Sub

' 窗体关闭查询
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If m_processing Then
        Cancel = True
        Call ShowWarning("正在处理中，请先取消操作或等待完成！")
    Else
        Call SaveSettings
    End If
End Sub

' ===============================================================================
' 控件初始化
' ===============================================================================

' 初始化控件
Private Sub InitializeControls()
    ' 设置窗体属性
    Me.Caption = "批量批注工具 v" & GetVersion()
    Me.Width = 480
    Me.Height = 420
    
    ' 创建控件
    Call CreateControls
    Call SetControlProperties
    Call SetTabOrder
End Sub

' 创建控件
Private Sub CreateControls()
    ' 目标区域组
    Set txtTargetRange = Me.Controls.Add("Forms.TextBox.1", "txtTargetRange")
    Set btnSelectRange = Me.Controls.Add("Forms.CommandButton.1", "btnSelectRange")
    
    ' 批注来源组
    Set optFromCell = Me.Controls.Add("Forms.OptionButton.1", "optFromCell")
    Set optFromText = Me.Controls.Add("Forms.OptionButton.1", "optFromText")
    Set optFromFile = Me.Controls.Add("Forms.OptionButton.1", "optFromFile")
    Set txtCommentSource = Me.Controls.Add("Forms.TextBox.1", "txtCommentSource")
    Set btnSelectSource = Me.Controls.Add("Forms.CommandButton.1", "btnSelectSource")
    Set txtFixedText = Me.Controls.Add("Forms.TextBox.1", "txtFixedText")
    
    ' 格式设置组
    Set cmbTemplate = Me.Controls.Add("Forms.ComboBox.1", "cmbTemplate")
    Set cmbFontName = Me.Controls.Add("Forms.ComboBox.1", "cmbFontName")
    Set cmbFontSize = Me.Controls.Add("Forms.ComboBox.1", "cmbFontSize")
    Set cmbFontColor = Me.Controls.Add("Forms.ComboBox.1", "cmbFontColor")
    Set chkBold = Me.Controls.Add("Forms.CheckBox.1", "chkBold")
    Set chkItalic = Me.Controls.Add("Forms.CheckBox.1", "chkItalic")
    
    ' 进度显示
    Set progressBar = Me.Controls.Add("Forms.Label.1", "progressBar")
    Set lblProgress = Me.Controls.Add("Forms.Label.1", "lblProgress")
    
    ' 按钮组
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    Set btnHelp = Me.Controls.Add("Forms.CommandButton.1", "btnHelp")
    Set btnTemplateManager = Me.Controls.Add("Forms.CommandButton.1", "btnTemplateManager")
End Sub

' 设置控件属性
Private Sub SetControlProperties()
    ' 目标区域
    With txtTargetRange
        .Left = 120: .Top = 30: .Width = 240: .Height = 18
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnSelectRange
        .Left = 370: .Top = 28: .Width = 60: .Height = 22
        .Caption = "选择": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    ' 批注来源选项
    With optFromCell
        .Left = 20: .Top = 70: .Width = 100: .Height = 18
        .Caption = "来自单元格": .Value = True
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With optFromText
        .Left = 20: .Top = 95: .Width = 100: .Height = 18
        .Caption = "固定文本": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With optFromFile
        .Left = 20: .Top = 120: .Width = 100: .Height = 18
        .Caption = "从文件导入": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With txtCommentSource
        .Left = 120: .Top = 68: .Width = 240: .Height = 18
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnSelectSource
        .Left = 370: .Top = 66: .Width = 60: .Height = 22
        .Caption = "选择": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With txtFixedText
        .Left = 120: .Top = 93: .Width = 310: .Height = 40
        .MultiLine = True: .ScrollBars = fmScrollBarsVertical
        .Font.Name = "微软雅黑": .Font.Size = 9
        .Enabled = False
    End With
    
    ' 格式设置
    With cmbTemplate
        .Left = 120: .Top = 150: .Width = 150: .Height = 18
        .Style = fmStyleDropDownCombo
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With cmbFontName
        .Left = 120: .Top = 180: .Width = 120: .Height = 18
        .Style = fmStyleDropDownCombo
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With cmbFontSize
        .Left = 250: .Top = 180: .Width = 60: .Height = 18
        .Style = fmStyleDropDownCombo
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With cmbFontColor
        .Left = 320: .Top = 180: .Width = 80: .Height = 18
        .Style = fmStyleDropDownList
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With chkBold
        .Left = 120: .Top = 210: .Width = 60: .Height = 18
        .Caption = "粗体": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With chkItalic
        .Left = 190: .Top = 210: .Width = 60: .Height = 18
        .Caption = "斜体": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    ' 进度显示
    With progressBar
        .Left = 20: .Top = 250: .Width = 410: .Height = 20
        .BackColor = &HE0E0E0: .BorderStyle = fmBorderStyleSingle
        .Visible = False
    End With
    
    With lblProgress
        .Left = 20: .Top = 275: .Width = 410: .Height = 15
        .Caption = "准备就绪": .Font.Name = "微软雅黑": .Font.Size = 8
        .Visible = False
    End With
    
    ' 按钮
    With btnOK
        .Left = 200: .Top = 310: .Width = 70: .Height = 25
        .Caption = "确定": .Default = True
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnCancel
        .Left = 280: .Top = 310: .Width = 70: .Height = 25
        .Caption = "取消": .Cancel = True
        .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnHelp
        .Left = 360: .Top = 310: .Width = 70: .Height = 25
        .Caption = "帮助": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
    
    With btnTemplateManager
        .Left = 280: .Top = 148: .Width = 80: .Height = 22
        .Caption = "模板管理": .Font.Name = "微软雅黑": .Font.Size = 9
    End With
End Sub

' 设置Tab顺序
Private Sub SetTabOrder()
    txtTargetRange.TabIndex = 0
    btnSelectRange.TabIndex = 1
    optFromCell.TabIndex = 2
    optFromText.TabIndex = 3
    optFromFile.TabIndex = 4
    txtCommentSource.TabIndex = 5
    btnSelectSource.TabIndex = 6
    txtFixedText.TabIndex = 7
    cmbTemplate.TabIndex = 8
    btnTemplateManager.TabIndex = 9
    cmbFontName.TabIndex = 10
    cmbFontSize.TabIndex = 11
    cmbFontColor.TabIndex = 12
    chkBold.TabIndex = 13
    chkItalic.TabIndex = 14
    btnOK.TabIndex = 15
    btnCancel.TabIndex = 16
    btnHelp.TabIndex = 17
End Sub

' ===============================================================================
' 控件事件处理
' ===============================================================================

' 选择目标区域
Private Sub btnSelectRange_Click()
    On Error GoTo ErrorHandler
    
    Me.Hide
    
    Dim selectedRange As Range
    Set selectedRange = Application.InputBox("请选择要插入批注的区域:", "选择区域", txtTargetRange.Value, Type:=8)
    
    If Not selectedRange Is Nothing Then
        txtTargetRange.Value = selectedRange.Address
        Set m_targetRange = selectedRange
    End If
    
    Me.Show
    Exit Sub
    
ErrorHandler:
    Me.Show
    If Err.Number <> 424 Then ' 用户取消
        Call HandleError(etRangeError, Err.Number, Err.Description, "BatchCommentForm", "btnSelectRange_Click")
    End If
End Sub

' 选择批注来源
Private Sub btnSelectSource_Click()
    On Error GoTo ErrorHandler
    
    If optFromCell.Value Then
        Call SelectSourceRange
    ElseIf optFromFile.Value Then
        Call SelectSourceFile
    End If
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etFileAccess, Err.Number, Err.Description, "BatchCommentForm", "btnSelectSource_Click")
End Sub

' 选择来源区域
Private Sub SelectSourceRange()
    Me.Hide
    
    Dim sourceRange As Range
    Set sourceRange = Application.InputBox("请选择批注来源区域:", "选择来源", txtCommentSource.Value, Type:=8)
    
    If Not sourceRange Is Nothing Then
        txtCommentSource.Value = sourceRange.Address
    End If
    
    Me.Show
End Sub

' 选择来源文件
Private Sub SelectSourceFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "选择批注文件"
        .Filters.Clear
        .Filters.Add "文本文件", "*.txt"
        .Filters.Add "Excel文件", "*.xlsx;*.xls"
        .Filters.Add "CSV文件", "*.csv"
        .Filters.Add "所有文件", "*.*"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            txtCommentSource.Value = .SelectedItems(1)
        End If
    End With
End Sub

' 批注来源选项变化
Private Sub optFromCell_Click()
    Call UpdateSourceControls
End Sub

Private Sub optFromText_Click()
    Call UpdateSourceControls
End Sub

Private Sub optFromFile_Click()
    Call UpdateSourceControls
End Sub

' 更新来源控件状态
Private Sub UpdateSourceControls()
    If optFromCell.Value Then
        txtCommentSource.Enabled = True
        btnSelectSource.Enabled = True
        txtFixedText.Enabled = False
        btnSelectSource.Caption = "选择"
    ElseIf optFromText.Value Then
        txtCommentSource.Enabled = False
        btnSelectSource.Enabled = False
        txtFixedText.Enabled = True
    ElseIf optFromFile.Value Then
        txtCommentSource.Enabled = True
        btnSelectSource.Enabled = True
        txtFixedText.Enabled = False
        btnSelectSource.Caption = "浏览"
    End If
End Sub

' 模板选择变化
Private Sub cmbTemplate_Change()
    If cmbTemplate.ListIndex >= 0 Then
        Call ApplyTemplate(cmbTemplate.Value)
    End If
End Sub

' 确定按钮
Private Sub btnOK_Click()
    On Error GoTo ErrorHandler
    
    If ValidateInput() Then
        Call ProcessBatchComment
    End If
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "BatchCommentForm", "btnOK_Click")
End Sub

' 取消按钮
Private Sub btnCancel_Click()
    If m_processing Then
        m_cancelled = True
        Call ShowInfo("正在取消操作，请稍候...")
    Else
        Unload Me
    End If
End Sub

' 帮助按钮
Private Sub btnHelp_Click()
    Call ShowHelp
End Sub

' 模板管理按钮
Private Sub btnTemplateManager_Click()
    Call ShowTemplateManager
    Call LoadTemplates ' 重新加载模板列表
End Sub

' ===============================================================================
' 业务逻辑
' ===============================================================================

' 加载模板
Private Sub LoadTemplates()
    On Error Resume Next
    
    cmbTemplate.Clear
    
    Dim templateNames() As String
    templateNames = GetTemplateNames()
    
    Dim i As Long
    For i = 0 To UBound(templateNames)
        cmbTemplate.AddItem templateNames(i)
    Next i
    
    If cmbTemplate.ListCount > 0 Then
        cmbTemplate.ListIndex = 0
    End If
    
    On Error GoTo 0
End Sub

' 设置默认值
Private Sub SetDefaultValues()
    ' 字体列表
    cmbFontName.Clear
    cmbFontName.AddItem "微软雅黑"
    cmbFontName.AddItem "宋体"
    cmbFontName.AddItem "Arial"
    cmbFontName.AddItem "Times New Roman"
    cmbFontName.AddItem "Calibri"
    cmbFontName.Value = "微软雅黑"
    
    ' 字体大小
    cmbFontSize.Clear
    Dim sizes As Variant
    sizes = Array("8", "9", "10", "11", "12", "14", "16", "18", "20", "24")
    
    Dim i As Long
    For i = 0 To UBound(sizes)
        cmbFontSize.AddItem sizes(i)
    Next i
    cmbFontSize.Value = "9"
    
    ' 颜色列表
    cmbFontColor.Clear
    cmbFontColor.AddItem "黑色"
    cmbFontColor.AddItem "红色"
    cmbFontColor.AddItem "蓝色"
    cmbFontColor.AddItem "绿色"
    cmbFontColor.AddItem "紫色"
    cmbFontColor.AddItem "橙色"
    cmbFontColor.Value = "黑色"
    
    ' 默认选项
    optFromCell.Value = True
    Call UpdateSourceControls
End Sub

' 应用模板
Private Sub ApplyTemplate(templateName As String)
    On Error Resume Next
    
    m_currentTemplate = GetTemplate(templateName)
    
    With m_currentTemplate
        cmbFontName.Value = .FontName
        cmbFontSize.Value = CStr(.FontSize)
        chkBold.Value = .IsBold
        chkItalic.Value = .IsItalic
        
        ' 设置颜色
        Select Case .FontColor
            Case 1: cmbFontColor.Value = "黑色"
            Case 3: cmbFontColor.Value = "红色"
            Case 5: cmbFontColor.Value = "蓝色"
            Case 10: cmbFontColor.Value = "绿色"
            Case Else: cmbFontColor.Value = "黑色"
        End Select
        
        ' 如果有默认文本且选择了固定文本选项
        If .DefaultText <> "" And optFromText.Value Then
            txtFixedText.Value = .DefaultText
        End If
    End With
    
    On Error GoTo 0
End Sub

' 验证输入
Private Function ValidateInput() As Boolean
    ValidateInput = False
    
    ' 验证目标区域
    If Trim(txtTargetRange.Value) = "" Then
        Call ShowWarning("请选择要插入批注的区域！")
        txtTargetRange.SetFocus
        Exit Function
    End If
    
    ' 验证区域有效性
    If Not IsValidRange(txtTargetRange.Value) Then
        Call ShowWarning("目标区域格式不正确！")
        txtTargetRange.SetFocus
        Exit Function
    End If
    
    ' 验证批注来源
    If optFromCell.Value Then
        If Trim(txtCommentSource.Value) = "" Then
            Call ShowWarning("请选择批注来源区域！")
            txtCommentSource.SetFocus
            Exit Function
        End If
        
        If Not IsValidRange(txtCommentSource.Value) Then
            Call ShowWarning("来源区域格式不正确！")
            txtCommentSource.SetFocus
            Exit Function
        End If
        
    ElseIf optFromText.Value Then
        If Trim(txtFixedText.Value) = "" Then
            Call ShowWarning("请输入固定文本内容！")
            txtFixedText.SetFocus
            Exit Function
        End If
        
    ElseIf optFromFile.Value Then
        If Trim(txtCommentSource.Value) = "" Then
            Call ShowWarning("请选择批注文件！")
            Exit Function
        End If
        
        If Dir(txtCommentSource.Value) = "" Then
            Call ShowWarning("指定的文件不存在！")
            Exit Function
        End If
    End If
    
    ValidateInput = True
End Function

' 处理批量批注
Private Sub ProcessBatchComment()
    On Error GoTo ErrorHandler
    
    m_processing = True
    m_cancelled = False
    
    ' 显示进度
    Call ShowProgress(True)
    
    ' 获取目标区域
    Set m_targetRange = Range(txtTargetRange.Value)
    
    ' 获取批注数据
    m_sourceData = GetCommentData()
    
    If IsEmpty(m_sourceData) Then
        Call ShowWarning("无法获取批注数据！")
        GoTo Cleanup
    End If
    
    ' 创建备份
    If MsgBox("是否创建工作表备份？", vbYesNo + vbQuestion, "确认") = vbYes Then
        Call CreateBackup(ActiveSheet)
    End If
    
    ' 执行批量批注
    Call ApplyBatchComments
    
    If Not m_cancelled Then
        Call ShowInfo("批量批注完成！共处理 " & m_targetRange.Cells.Count & " 个单元格。")
        Call LogOperation("BATCH_COMMENT", "成功处理 " & m_targetRange.Cells.Count & " 个批注")
    End If
    
Cleanup:
    Call ShowProgress(False)
    m_processing = False
    
    If Not m_cancelled Then
        Unload Me
    End If
    
    Exit Sub
    
ErrorHandler:
    Call HandleError(etGeneral, Err.Number, Err.Description, "BatchCommentForm", "ProcessBatchComment")
    GoTo Cleanup
End Sub

' 获取批注数据
Private Function GetCommentData() As Variant
    On Error GoTo ErrorHandler
    
    If optFromCell.Value Then
        ' 从单元格获取
        Dim sourceRange As Range
        Set sourceRange = Range(txtCommentSource.Value)
        GetCommentData = sourceRange.Value
        
    ElseIf optFromText.Value Then
        ' 固定文本
        GetCommentData = txtFixedText.Value
        
    ElseIf optFromFile.Value Then
        ' 从文件获取
        GetCommentData = LoadCommentFromFile(txtCommentSource.Value)
    End If
    
    Exit Function
    
ErrorHandler:
    GetCommentData = Empty
End Function

' 应用批量批注
Private Sub ApplyBatchComments()
    On Error Resume Next
    
    Dim cell As Range
    Dim commentText As String
    Dim totalCells As Long
    Dim processedCells As Long
    Dim startTime As Double
    
    totalCells = m_targetRange.Cells.Count
    processedCells = 0
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    For Each cell In m_targetRange
        If m_cancelled Then Exit For
        
        processedCells = processedCells + 1
        
        ' 更新进度（每50个单元格更新一次）
        If processedCells Mod 50 = 0 Or processedCells = totalCells Then
            Call UpdateProgress(processedCells, totalCells, startTime)
            DoEvents
        End If
        
        ' 获取批注文本
        commentText = GetCommentTextForCell(cell)
        
        If Trim(commentText) <> "" Then
            ' 删除现有批注
            If Not cell.Comment Is Nothing Then
                cell.Comment.Delete
            End If
            
            ' 添加新批注
            cell.AddComment commentText
            
            ' 设置批注格式
            Call FormatComment(cell.Comment)
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 最终进度更新
    Call UpdateProgress(processedCells, totalCells, startTime)
End Sub

' 获取单元格对应的批注文本
Private Function GetCommentTextForCell(cell As Range) As String
    On Error Resume Next
    
    If IsArray(m_sourceData) Then
        ' 数组数据 - 按位置对应
        Dim rowOffset As Long
        Dim colOffset As Long
        
        rowOffset = cell.Row - m_targetRange.Row + 1
        colOffset = cell.Column - m_targetRange.Column + 1
        
        ' 检查数组边界
        If rowOffset <= UBound(m_sourceData, 1) And colOffset <= UBound(m_sourceData, 2) Then
            GetCommentTextForCell = CStr(m_sourceData(rowOffset, colOffset))
        End If
    Else
        ' 固定文本
        GetCommentTextForCell = CStr(m_sourceData)
    End If
    
    ' 替换占位符
    GetCommentTextForCell = ReplacePlaceholders(GetCommentTextForCell, cell)
End Function

' 设置批注格式
Private Sub FormatComment(comment As Comment)
    On Error Resume Next
    
    With comment.Shape.TextFrame.Characters.Font
        .Name = cmbFontName.Value
        .Size = CLng(cmbFontSize.Value)
        .ColorIndex = GetSelectedColorIndex()
        .Bold = chkBold.Value
        .Italic = chkItalic.Value
    End With
    
    ' 设置批注大小
    With comment.Shape
        .Width = 200
        .Height = 100
        
        ' 自动调整大小
        .TextFrame.AutoSize = True
        If .Width > 300 Then .Width = 300
        If .Height > 150 Then .Height = 150
    End With
    
    On Error GoTo 0
End Sub

' 获取选中的颜色索引
Private Function GetSelectedColorIndex() As Long
    Select Case cmbFontColor.Value
        Case "黑色": GetSelectedColorIndex = 1
        Case "红色": GetSelectedColorIndex = 3
        Case "蓝色": GetSelectedColorIndex = 5
        Case "绿色": GetSelectedColorIndex = 10
        Case "紫色": GetSelectedColorIndex = 13
        Case "橙色": GetSelectedColorIndex = 46
        Case Else: GetSelectedColorIndex = 1
    End Select
End Function

' ===============================================================================
' 进度显示
' ===============================================================================

' 显示/隐藏进度
Private Sub ShowProgress(show As Boolean)
    progressBar.Visible = show
    lblProgress.Visible = show
    
    If show Then
        progressBar.Width = 0
        progressBar.BackColor = &H8000000F ' 系统按钮面颜色
        lblProgress.Caption = "准备处理..."
    End If
End Sub

' 更新进度
Private Sub UpdateProgress(processed As Long, total As Long, startTime As Double)
    Dim percentage As Double
    Dim elapsed As Double
    Dim estimated As Double
    Dim remaining As Double
    
    percentage = (processed / total) * 100
    elapsed = Timer - startTime
    
    If processed > 0 Then
        estimated = elapsed * total / processed
        remaining = estimated - elapsed
    End If
    
    ' 更新进度条
    progressBar.Width = (410 * percentage) / 100
    progressBar.BackColor = &H8000000D ' 系统高亮颜色
    
    ' 更新进度文本
    lblProgress.Caption = "正在处理: " & processed & "/" & total & _
                         " (" & Format(percentage, "0.0") & "%) " & _
                         "剩余时间: " & Format(remaining, "0") & "秒"
End Sub

' ===============================================================================
' 设置管理
' ===============================================================================

' 加载设置
Private Sub LoadSettings()
    On Error Resume Next
    
    ' 这里可以从注册表或配置文件加载用户设置
    ' 暂时使用默认值
    
    On Error GoTo 0
End Sub

' 保存设置
Private Sub SaveSettings()
    On Error Resume Next
    
    ' 这里可以保存用户设置到注册表或配置文件
    
    On Error GoTo 0
End Sub

' ===============================================================================
' 帮助功能
' ===============================================================================

' 显示帮助
Private Sub ShowHelp()
    Dim helpText As String
    helpText = "批量批注工具 v" & GetVersion() & vbCrLf & vbCrLf
    helpText = helpText & "使用说明:" & vbCrLf
    helpText = helpText & "1. 选择目标区域：要添加批注的单元格区域" & vbCrLf
    helpText = helpText & "2. 选择批注来源：" & vbCrLf
    helpText = helpText & "   • 来自单元格：从指定区域获取批注内容" & vbCrLf
    helpText = helpText & "   • 固定文本：所有单元格使用相同批注" & vbCrLf
    helpText = helpText & "   • 从文件导入：从文本或Excel文件导入" & vbCrLf
    helpText = helpText & "3. 设置格式：字体、大小、颜色等" & vbCrLf
    helpText = helpText & "4. 点击确定开始处理" & vbCrLf & vbCrLf
    helpText = helpText & "占位符支持:" & vbCrLf
    helpText = helpText & "{CELL} - 单元格地址" & vbCrLf
    helpText = helpText & "{VALUE} - 单元格值" & vbCrLf
    helpText = helpText & "{ROW} - 行号" & vbCrLf
    helpText = helpText & "{COLUMN} - 列号" & vbCrLf
    helpText = helpText & "{DATE} - 当前日期" & vbCrLf
    helpText = helpText & "{TIME} - 当前时间" & vbCrLf
    helpText = helpText & "{USER} - 用户名" & vbCrLf & vbCrLf
    helpText = helpText & "注意：现有批注将被替换！"
    
    MsgBox helpText, vbInformation, "帮助"
End Sub