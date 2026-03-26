VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBrowsePrintPrescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 浏览与打印处方信息表单
' 表单名称：frmBrowsePrintPrescription
' 按考试要求设计：
' 1. 对话框样式，无导航按钮、记录选择器、分割线、滚动条
' 2. 控制框无最大化/最小化按钮
' 3. 标题：浏览与打印处方信息
' 4. 标题标签：STXinwei 24号加粗蓝色居中
' 5. 设计师标签：楷体12号加粗红色，显示"设计师：[考生姓名]"
' 6. 选择医师下拉框，刷新按钮
' 7. 子窗体仅显示5个字段（患者姓名、临床诊断、方剂名称、处方日期、处方金额）
' 8. 预览医师处方列表按钮，打开报表打印预览

Option Explicit

' 当前选中的医师ID
Private m_SelectedPhysicianID As String
Private m_SelectedPhysicianName As String
Private m_SelectedPhysicianTitle As String

' 表单加载事件
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' 设置表单为对话框样式
    Call UIUtils.SetFormAsDialog(Me)
    
    ' 设置表单标题
    Me.Caption = "浏览与打印处方信息"
    
    ' 设置标题标签样式
    Call UIUtils.SetLabelFontStyle(Me.lblTitle, UIUtils.FONT_STXINWEI, 24, True, UIUtils.COLOR_BLUE, True)
    
    ' 设置设计师标签样式
    Call UIUtils.SetLabelFontStyle(Me.lblDesigner, UIUtils.FONT_KAITI, 12, True, UIUtils.COLOR_RED, False)
    
    ' 初始化控件
    InitializeControls
    
    ' 加载医师列表
    LoadPhysicianList
    
    ' 检查字体
    Call UIUtils.CheckRequiredFonts
    
    ' 默认选择第一个医师
    If Me.cboPhysician.ListCount > 0 Then
        Me.cboPhysician = Me.cboPhysician.ItemData(0)
        m_SelectedPhysicianID = Me.cboPhysician.ItemData(0)
        m_SelectedPhysicianName = Me.cboPhysician.Column(1, 0)
        m_SelectedPhysicianTitle = Nz(DLookup("Physician_Title", "Doctor_Table", "Physician_ID = '" & m_SelectedPhysicianID & "'"), "")
        
        ' 刷新子窗体数据
        RefreshSubformData
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "表单加载错误：" & Err.Description, vbCritical, "错误"
End Sub

' 初始化控件
Private Sub InitializeControls()
    On Error Resume Next
    
    ' 设置按钮样式
    Me.cmdRefresh.Caption = "刷新"
    Me.cmdPreview.Caption = "预览医师处方列表"
    
    ' 设置下拉框样式
    Me.cboPhysician.RowSourceType = "Table/Query"
    Me.cboPhysician.ColumnCount = 2
    Me.cboPhysician.ColumnWidths = "0cm;3cm" ' 隐藏ID列，显示姓名列
    Me.cboPhysician.BoundColumn = 1 ' 绑定到医师ID
    
    ' 设置子窗体属性
    SetupSubform
    
    ' 设置控件提示文本
    Call UIUtils.SetControlTooltip(Me.cboPhysician, "选择要查看处方的医师")
    Call UIUtils.SetControlTooltip(Me.cmdRefresh, "刷新选中医师的处方列表")
    Call UIUtils.SetControlTooltip(Me.cmdPreview, "预览并打印当前医师的处方列表报表")
    Call UIUtils.SetControlTooltip(Me.lblDesigner, "设计师：考生姓名（请替换为实际姓名）")
End Sub

' 设置子窗体
Private Sub SetupSubform()
    On Error GoTo ErrorHandler
    
    ' 设置子窗体数据源（显示5个字段的查询）
    ' 注意：需要在Access中创建名为"qryPrescriptionFiveFields"的查询
    ' 查询应包含：Patient_Name, Clinical_Diagnosis, Formula_Name, Prescription_Date, Prescription_Amount, Physician_ID
    
    Me.subPrescriptions.SourceObject = "Query.qryPrescriptionFiveFields"
    Me.subPrescriptions.LinkMasterFields = "cboPhysician"
    Me.subPrescriptions.LinkChildFields = "Physician_ID"
    
    ' 设置子窗体格式
    With Me.subPrescriptions.Form
        .NavigationButtons = False
        .RecordSelectors = False
        .ScrollBars = 2 ' acScrollBarsVertical
        .BorderStyle = 1 ' acFixedSingle
    End With
    
    Exit Sub
    
ErrorHandler:
    ' 如果查询不存在，显示提示
    MsgBox "子窗体查询 'qryPrescriptionFiveFields' 不存在。" & vbCrLf & _
           "请在Access中创建包含以下字段的查询：" & vbCrLf & _
           "Patient_Name, Clinical_Diagnosis, Formula_Name, Prescription_Date, Prescription_Amount, Physician_ID", _
           vbExclamation, "查询缺失"
End Sub

' 加载医师列表
Private Sub LoadPhysicianList()
    On Error GoTo ErrorHandler
    
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' 构建SQL查询，获取所有医师
    sql = "SELECT Physician_ID, Physician_Name FROM Doctor_Table ORDER BY Physician_Name"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    ' 清空下拉框
    Me.cboPhysician.RowSource = ""
    
    If Not rs.EOF Then
        ' 构建行来源字符串
        Dim rowSource As String
        rowSource = ""
        
        Do While Not rs.EOF
            If rowSource <> "" Then rowSource = rowSource & ";"
            rowSource = rowSource & rs!Physician_ID & ";" & rs!Physician_Name
            rs.MoveNext
        Loop
        
        Me.cboPhysician.RowSource = rowSource
    Else
        MsgBox "医师表中没有数据，请先添加医师信息。", vbExclamation, "数据缺失"
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "加载医师列表错误：" & Err.Description, vbCritical, "错误"
End Sub

' 医师选择改变事件
Private Sub cboPhysician_AfterUpdate()
    On Error GoTo ErrorHandler
    
    If Nz(Me.cboPhysician.Value, "") <> "" Then
        m_SelectedPhysicianID = Me.cboPhysician.Value
        m_SelectedPhysicianName = Me.cboPhysician.Column(1)
        
        ' 获取医师职称
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset( _
            "SELECT Physician_Title FROM Doctor_Table WHERE Physician_ID = '" & m_SelectedPhysicianID & "'")
        
        If Not rs.EOF Then
            m_SelectedPhysicianTitle = Nz(rs!Physician_Title, "")
        Else
            m_SelectedPhysicianTitle = ""
        End If
        
        rs.Close
        Set rs = Nothing
        
        ' 刷新子窗体数据
        RefreshSubformData
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "医师选择更新错误：" & Err.Description, vbCritical, "错误"
End Sub

' 刷新按钮点击事件
Private Sub cmdRefresh_Click()
    On Error GoTo ErrorHandler
    
    If Nz(Me.cboPhysician.Value, "") = "" Then
        MsgBox "请先选择医师。", vbExclamation, "提示"
        Me.cboPhysician.SetFocus
        Exit Sub
    End If
    
    RefreshSubformData
    
    MsgBox "处方数据已刷新。", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "刷新数据错误：" & Err.Description, vbCritical, "错误"
End Sub

' 刷新子窗体数据
Private Sub RefreshSubformData()
    On Error GoTo ErrorHandler
    
    ' 更新子窗体筛选条件
    If Nz(m_SelectedPhysicianID, "") <> "" Then
        Me.subPrescriptions.Form.Filter = "Physician_ID = '" & m_SelectedPhysicianID & "'"
        Me.subPrescriptions.Form.FilterOn = True
        
        ' 更新状态标签
        Me.lblStatus.Caption = "当前医师：" & m_SelectedPhysicianName & "，处方数量：" & _
                               GetPrescriptionCount(m_SelectedPhysicianID)
    Else
        Me.subPrescriptions.Form.Filter = ""
        Me.subPrescriptions.Form.FilterOn = False
        Me.lblStatus.Caption = "请选择医师"
    End If
    
    ' 刷新子窗体
    Me.subPrescriptions.Form.Requery
    
    Exit Sub
    
ErrorHandler:
    MsgBox "刷新子窗体数据错误：" & Err.Description, vbCritical, "错误"
End Sub

' 获取医师处方数量
Private Function GetPrescriptionCount(physicianID As String) As Long
    On Error GoTo ErrorHandler
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT COUNT(*) AS PrescriptionCount FROM TCM_Prescription_Table WHERE Physician_ID = '" & physicianID & "'")
    
    If Not rs.EOF Then
        GetPrescriptionCount = Nz(rs!PrescriptionCount, 0)
    Else
        GetPrescriptionCount = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
ErrorHandler:
    GetPrescriptionCount = -1
End Function

' 预览按钮点击事件
Private Sub cmdPreview_Click()
    On Error GoTo ErrorHandler
    
    If Nz(Me.cboPhysician.Value, "") = "" Then
        MsgBox "请先选择医师。", vbExclamation, "提示"
        Me.cboPhysician.SetFocus
        Exit Sub
    End If
    
    ' 检查是否有处方数据
    If GetPrescriptionCount(m_SelectedPhysicianID) = 0 Then
        MsgBox "选中的医师没有处方记录，无法生成报表。", vbExclamation, "无数据"
        Exit Sub
    End If
    
    ' 打开报表（需要传递医师ID参数）
    OpenPrescriptionReport
    
    Exit Sub
    
ErrorHandler:
    MsgBox "预览报表错误：" & Err.Description, vbCritical, "错误"
End Sub

' 打开处方报表
Private Sub OpenPrescriptionReport()
    On Error GoTo ErrorHandler
    
    ' 方法1：使用OpenReport带Where条件
    ' DoCmd.OpenReport "rptPhysicianPrescriptionList", acViewPreview, , "Physician_ID = '" & m_SelectedPhysicianID & "'"
    
    ' 方法2：使用临时表或全局变量传递参数
    ' 这里使用全局变量传递医师信息（需要在模块中定义）
    SetReportParameters
    
    ' 打开报表
    DoCmd.OpenReport "rptPhysicianPrescriptionList", acViewPreview
    
    Exit Sub
    
ErrorHandler:
    MsgBox "打开报表错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置报表参数
Private Sub SetReportParameters()
    On Error Resume Next
    
    ' 使用全局变量传递参数（需要在标准模块中定义）
    ' 示例：Global g_ReportPhysicianID As String
    '        Global g_ReportPhysicianName As String
    '        Global g_ReportPhysicianTitle As String
    
    ' 如果全局变量已定义，则设置值
    ' g_ReportPhysicianID = m_SelectedPhysicianID
    ' g_ReportPhysicianName = m_SelectedPhysicianName
    ' g_ReportPhysicianTitle = m_SelectedPhysicianTitle
    
    ' 或者使用OpenArgs传递参数
    ' DoCmd.OpenReport "rptPhysicianPrescriptionList", acViewPreview, , , , m_SelectedPhysicianID & "|" & m_SelectedPhysicianName & "|" & m_SelectedPhysicianTitle
    
    ' 实际应用中，报表应基于查询自动筛选数据
    ' 这里简化处理，显示消息
    MsgBox "将打开医师 '" & m_SelectedPhysicianName & "' 的处方列表报表。" & vbCrLf & _
           "在实际应用中，报表应自动筛选并显示该医师的所有处方。", vbInformation, "报表预览"
End Sub

' 关闭表单按钮点击事件
Private Sub cmdClose_Click()
    On Error Resume Next
    
    DoCmd.Close acForm, Me.Name
End Sub

' 表单激活事件
Private Sub Form_Activate()
    On Error Resume Next
    
    ' 刷新数据
    If Nz(m_SelectedPhysicianID, "") <> "" Then
        RefreshSubformData
    End If
End Sub

' 表单卸载事件
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    ' 清理资源
    m_SelectedPhysicianID = ""
    m_SelectedPhysicianName = ""
    m_SelectedPhysicianTitle = ""
End Sub

' 快捷键处理
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyEscape ' ESC键关闭表单
            If Shift = 0 Then
                cmdClose_Click
            End If
            
        Case vbKeyF5 ' F5键刷新
            If Shift = 0 Then
                cmdRefresh_Click
            End If
            
        Case vbKeyP ' Ctrl+P打印预览
            If Shift = vbCtrlMask Then
                cmdPreview_Click
            End If
    End Select
End Sub

' 双击子窗体记录打开详细信息
Private Sub subPrescriptions_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    ' 获取当前处方ID
    Dim prescriptionID As String
    prescriptionID = Nz(Me.subPrescriptions.Form!Prescription_ID, "")
    
    If prescriptionID <> "" Then
        ' 打开处方详细信息表单
        ' DoCmd.OpenForm "frmPrescriptionDetail", acNormal, , "Prescription_ID = '" & prescriptionID & "'", , acDialog
        MsgBox "处方ID：" & prescriptionID & vbCrLf & _
               "在实际应用中，这里可以打开处方详细信息表单。", vbInformation, "处方详情"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "打开处方详情错误：" & Err.Description, vbCritical, "错误"
End Sub

' 更新设计师姓名（示例函数）
Public Sub UpdateDesignerName(designerName As String)
    On Error Resume Next
    
    Me.lblDesigner.Caption = "设计师：" & designerName
End Sub

' 导出当前医师处方数据
Public Sub ExportCurrentPrescriptions()
    On Error GoTo ErrorHandler
    
    If Nz(m_SelectedPhysicianID, "") = "" Then
        MsgBox "请先选择医师。", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 导出数据到Excel或文本文件
    Dim exportPath As String
    exportPath = CurrentProject.Path & "\" & m_SelectedPhysicianName & "_处方列表_" & Format(Date, "yyyy-mm-dd") & ".xlsx"
    
    ' 使用TransferSpreadsheet方法导出
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
        "qryPrescriptionFiveFields", exportPath, True, "Physician_ID = '" & m_SelectedPhysicianID & "'"
    
    MsgBox "处方数据已导出到：" & vbCrLf & exportPath, vbInformation, "导出完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "导出数据错误：" & Err.Description, vbCritical, "错误"
End Sub