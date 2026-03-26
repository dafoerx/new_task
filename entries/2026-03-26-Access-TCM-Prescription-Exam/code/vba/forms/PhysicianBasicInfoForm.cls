VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPhysicianBasicInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 医师基本信息管理表单
' 表单名称：frmPhysicianBasicInfo
' 按考试要求设计：
' 1. 对话框样式，无导航按钮、记录选择器、分割线、滚动条
' 2. 控制框无最大化/最小化按钮
' 3. 标题：医师基本信息管理
' 4. 标题标签：隶书22号加粗红色居中
' 5. 显示8个字段
' 6. 细节区域红色实线边框，宽度3pt
' 7. 5个导航按钮：首记录、上一条、下一条、末记录、关闭表单

Option Explicit

' 表单加载事件
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' 设置表单为对话框样式
    Call UIUtils.SetFormAsDialog(Me)
    
    ' 设置表单标题
    Me.Caption = "医师基本信息管理"
    
    ' 设置标题标签样式
    Call UIUtils.SetLabelFontStyle(Me.lblTitle, UIUtils.FONT_LISHU, 22, True, UIUtils.COLOR_RED, True)
    
    ' 设置细节区域边框
    Call UIUtils.SetDetailSectionBorder(Me, Me.Detail)
    
    ' 设置字段控件属性
    SetFieldProperties
    
    ' 检查字体
    Call UIUtils.CheckRequiredFonts
    
    ' 移动到第一条记录
    If Not (Me.Recordset.EOF And Me.Recordset.BOF) Then
        Me.Recordset.MoveFirst
    End If
    
    ' 更新按钮状态
    UpdateButtonState
    
    Exit Sub
    
ErrorHandler:
    MsgBox "表单加载错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置字段控件属性
Private Sub SetFieldProperties()
    On Error Resume Next
    
    ' 设置文本框边框
    Call UIUtils.SetTextBoxBorder(Me.txtPhysician_ID, UIUtils.COLOR_BLACK, 1)
    Call UIUtils.SetTextBoxBorder(Me.txtDepartment_Name, UIUtils.COLOR_BLACK, 1)
    Call UIUtils.SetTextBoxBorder(Me.txtPhysician_Name, UIUtils.COLOR_BLACK, 1)
    Call UIUtils.SetTextBoxBorder(Me.cboPhysician_Gender, UIUtils.COLOR_BLACK, 1)
    Call UIUtils.SetTextBoxBorder(Me.txtPhysician_Age, UIUtils.COLOR_BLACK, 1)
    Call UIUtils.SetTextBoxBorder(Me.imgPhysician_Photo, UIUtils.COLOR_BLACK, 1)
    Call UIUtils.SetTextBoxBorder(Me.txtPhysician_Education, UIUtils.COLOR_BLACK, 1)
    Call UIUtils.SetTextBoxBorder(Me.txtPhysician_Title, UIUtils.COLOR_BLACK, 1)
    
    ' 设置字段提示文本
    Call UIUtils.SetControlTooltip(Me.txtPhysician_ID, "医师编号，10位字符")
    Call UIUtils.SetControlTooltip(Me.txtDepartment_Name, "所属科室名称")
    Call UIUtils.SetControlTooltip(Me.txtPhysician_Name, "医师姓名")
    Call UIUtils.SetControlTooltip(Me.cboPhysician_Gender, "选择医师性别")
    Call UIUtils.SetControlTooltip(Me.txtPhysician_Age, "医师年龄")
    Call UIUtils.SetControlTooltip(Me.imgPhysician_Photo, "医师照片，支持JPG、BMP等格式")
    Call UIUtils.SetControlTooltip(Me.txtPhysician_Education, "医师学历")
    Call UIUtils.SetControlTooltip(Me.txtPhysician_Title, "医师职称")
    
    ' 设置性别下拉框选项
    Me.cboPhysician_Gender.RowSourceType = "Value List"
    Me.cboPhysician_Gender.RowSource = "男;女"
    
    ' 设置年龄字段输入验证
    Me.txtPhysician_Age.InputMask = "###"
End Sub

' 首记录按钮点击事件
Private Sub cmdFirst_Click()
    On Error GoTo ErrorHandler
    
    If Not (Me.Recordset.EOF And Me.Recordset.BOF) Then
        Me.Recordset.MoveFirst
        UpdateButtonState
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "移动到首记录错误：" & Err.Description, vbCritical, "错误"
End Sub

' 上一条按钮点击事件
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorHandler
    
    If Not Me.Recordset.BOF Then
        Me.Recordset.MovePrevious
        If Me.Recordset.BOF Then
            Me.Recordset.MoveFirst
        End If
        UpdateButtonState
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "移动到上一条记录错误：" & Err.Description, vbCritical, "错误"
End Sub

' 下一条按钮点击事件
Private Sub cmdNext_Click()
    On Error GoTo ErrorHandler
    
    If Not Me.Recordset.EOF Then
        Me.Recordset.MoveNext
        If Me.Recordset.EOF Then
            Me.Recordset.MoveLast
        End If
        UpdateButtonState
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "移动到下一条记录错误：" & Err.Description, vbCritical, "错误"
End Sub

' 末记录按钮点击事件
Private Sub cmdLast_Click()
    On Error GoTo ErrorHandler
    
    If Not (Me.Recordset.EOF And Me.Recordset.BOF) Then
        Me.Recordset.MoveLast
        UpdateButtonState
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "移动到末记录错误：" & Err.Description, vbCritical, "错误"
End Sub

' 关闭表单按钮点击事件
Private Sub cmdClose_Click()
    On Error Resume Next
    
    DoCmd.Close acForm, Me.Name
End Sub

' 更新按钮状态（根据记录位置）
Private Sub UpdateButtonState()
    On Error Resume Next
    
    Dim isAtFirst As Boolean
    Dim isAtLast As Boolean
    Dim isEmpty As Boolean
    
    isEmpty = (Me.Recordset.EOF And Me.Recordset.BOF)
    isAtFirst = isEmpty Or Me.Recordset.BOF
    isAtLast = isEmpty Or Me.Recordset.EOF
    
    ' 设置按钮可用状态
    Me.cmdFirst.Enabled = Not (isEmpty Or isAtFirst)
    Me.cmdPrevious.Enabled = Not (isEmpty Or isAtFirst)
    Me.cmdNext.Enabled = Not (isEmpty Or isAtLast)
    Me.cmdLast.Enabled = Not (isEmpty Or isAtLast)
    
    ' 更新状态栏
    If isEmpty Then
        Me.lblStatus.Caption = "无记录"
    Else
        Me.lblStatus.Caption = "记录 " & Me.Recordset.AbsolutePosition + 1 & " / " & Me.Recordset.RecordCount
    End If
End Sub

' 当前记录改变时更新按钮状态
Private Sub Form_Current()
    On Error Resume Next
    UpdateButtonState
End Sub

' 表单激活事件
Private Sub Form_Activate()
    On Error Resume Next
    UpdateButtonState
End Sub

' 双击照片字段打开文件对话框
Private Sub imgPhysician_Photo_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    Dim fd As FileDialog
    Dim selectedFile As Variant
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "选择医师照片"
        .Filters.Clear
        .Filters.Add "图像文件", "*.jpg;*.jpeg;*.bmp;*.png;*.gif"
        .Filters.Add "所有文件", "*.*"
        .AllowMultiSelect = False
        
        If .Show = True Then
            For Each selectedFile In .SelectedItems
                ' 在实际应用中，这里需要将图像文件加载到OLE字段
                ' 简化处理：显示文件路径
                Me.imgPhysician_Photo.ControlTipText = "照片文件：" & selectedFile
                MsgBox "已选择照片文件：" & selectedFile & vbCrLf & _
                       "注意：实际应用中需要将图像数据保存到数据库。", vbInformation, "照片选择"
            Next selectedFile
        End If
    End With
    
    Set fd = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "选择照片文件错误：" & Err.Description, vbCritical, "错误"
End Sub

' 验证数据
Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    Dim validationError As String
    
    validationError = ""
    
    ' 验证必填字段
    If Nz(Me.txtPhysician_ID.Value, "") = "" Then
        validationError = validationError & "• 医师编号不能为空" & vbCrLf
    End If
    
    If Nz(Me.txtPhysician_Name.Value, "") = "" Then
        validationError = validationError & "• 医师姓名不能为空" & vbCrLf
    End If
    
    If Nz(Me.cboPhysician_Gender.Value, "") = "" Then
        validationError = validationError & "• 请选择医师性别" & vbCrLf
    End If
    
    ' 验证年龄范围
    If Nz(Me.txtPhysician_Age.Value, "") <> "" Then
        If Not IsNumeric(Me.txtPhysician_Age.Value) Then
            validationError = validationError & "• 年龄必须为数字" & vbCrLf
        ElseIf CInt(Me.txtPhysician_Age.Value) < 18 Or CInt(Me.txtPhysician_Age.Value) > 100 Then
            validationError = validationError & "• 年龄应在18-100岁之间" & vbCrLf
        End If
    End If
    
    ' 如果有验证错误，取消更新并显示消息
    If validationError <> "" Then
        Cancel = True
        MsgBox "数据验证错误：" & vbCrLf & vbCrLf & validationError, vbExclamation, "验证错误"
        ' 将焦点设置到第一个错误字段
        If Nz(Me.txtPhysician_ID.Value, "") = "" Then
            Me.txtPhysician_ID.SetFocus
        ElseIf Nz(Me.txtPhysician_Name.Value, "") = "" Then
            Me.txtPhysician_Name.SetFocus
        ElseIf Nz(Me.cboPhysician_Gender.Value, "") = "" Then
            Me.cboPhysician_Gender.SetFocus
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "数据验证错误：" & Err.Description, vbCritical, "错误"
    Cancel = True
End Sub

' 表单卸载事件
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    ' 保存最后的状态
    If MsgBox("是否保存更改？", vbQuestion + vbYesNo, "确认") = vbYes Then
        If Me.Dirty Then
            Me.Dirty = False
        End If
    Else
        If Me.Dirty Then
            Me.Undo
        End If
    End If
End Sub

' 快捷键处理
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyEscape ' ESC键关闭表单
            If Shift = 0 Then
                cmdClose_Click
            End If
            
        Case vbKeyHome ' Home键首记录
            If Shift = 0 Then
                cmdFirst_Click
            End If
            
        Case vbKeyEnd ' End键末记录
            If Shift = 0 Then
                cmdLast_Click
            End If
            
        Case vbKeyLeft ' 左箭头上一记录
            If Shift = 0 Then
                cmdPrevious_Click
            End If
            
        Case vbKeyRight ' 右箭头下一记录
            If Shift = 0 Then
                cmdNext_Click
            End If
    End Select
End Sub