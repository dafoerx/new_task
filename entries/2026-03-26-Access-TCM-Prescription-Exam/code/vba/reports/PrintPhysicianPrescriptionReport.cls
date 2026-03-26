VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptPhysicianPrescriptionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 打印医师处方列表报表
' 报表名称：rptPhysicianPrescriptionList
' 按考试要求设计：
' 1. 报表标题包含医师姓名和职称，宋体24号加粗红色
' 2. 报表页脚统计处方数量和总金额，宋体14号加粗红色
' 3. 每条处方信息在蓝色边框矩形框内，边框宽度3pt
' 4. 处方ID：宋体18号黑色，其他字段宋体12号黑色
' 5. 患者年龄显示为"X岁"，处方金额显示为"¥X元"
' 6. 处方各区域用点划线分隔，宽度1磅黑色
' 7. 总计：xx处方，累计处方金额：xx元

Option Explicit

' 报表打开事件
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    ' 检查是否传递了医师ID参数
    Dim physicianID As String
    Dim physicianName As String
    Dim physicianTitle As String
    
    ' 从OpenArgs获取参数
    If Not IsMissing(Reports!rptPhysicianPrescriptionList.OpenArgs) Then
        Dim args As String
        args = Reports!rptPhysicianPrescriptionList.OpenArgs
        
        If args <> "" Then
            Dim parts() As String
            parts = Split(args, "|")
            
            If UBound(parts) >= 2 Then
                physicianID = parts(0)
                physicianName = parts(1)
                physicianTitle = parts(2)
            End If
        End If
    End If
    
    ' 如果没有参数，使用默认值或提示
    If physicianID = "" Then
        ' 这里可以设置默认医师或提示用户
        ' 简化处理：使用第一个医师
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 Physician_ID, Physician_Name, Physician_Title FROM Doctor_Table")
        
        If Not rs.EOF Then
            physicianID = Nz(rs!Physician_ID, "")
            physicianName = Nz(rs!Physician_Name, "")
            physicianTitle = Nz(rs!Physician_Title, "")
        End If
        
        rs.Close
        Set rs = Nothing
    End If
    
    ' 设置报表筛选条件
    If physicianID <> "" Then
        Me.Filter = "Physician_ID = '" & physicianID & "'"
        Me.FilterOn = True
        
        ' 设置报表标题
        SetReportTitle physicianName, physicianTitle
    Else
        Me.Filter = ""
        Me.FilterOn = False
        SetReportTitle "所有医师", ""
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "报表打开错误：" & Err.Description, vbCritical, "错误"
    Cancel = True
End Sub

' 设置报表标题
Private Sub SetReportTitle(physicianName As String, physicianTitle As String)
    On Error GoTo ErrorHandler
    
    Dim titleText As String
    
    If physicianTitle <> "" Then
        titleText = physicianName & " " & physicianTitle & " 处方列表"
    Else
        titleText = physicianName & " 处方列表"
    End If
    
    ' 设置报表标题标签
    Call UIUtils.SetLabelFontStyle(Me.lblReportTitle, UIUtils.FONT_SIMSUN, 24, True, UIUtils.COLOR_RED, True)
    Me.lblReportTitle.Caption = titleText
    
    ' 设置页面页眉标题
    Call UIUtils.SetLabelFontStyle(Me.lblPageTitle, UIUtils.FONT_SIMSUN, 18, True, UIUtils.COLOR_BLACK, True)
    Me.lblPageTitle.Caption = "中医处方管理数据库 - 处方列表"
    
    Exit Sub
    
ErrorHandler:
    Me.lblReportTitle.Caption = "处方列表"
    Me.lblReportTitle.FontName = "宋体"
    Me.lblReportTitle.FontSize = 24
    Me.lblReportTitle.FontBold = True
    Me.lblReportTitle.ForeColor = vbRed
    Me.lblReportTitle.TextAlign = 2
End Sub

' 报表格式化事件（每个处方记录）
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    On Error GoTo ErrorHandler
    
    ' 设置处方信息框样式
    Call UIUtils.SetPrescriptionBoxStyle(Me, Me.boxPrescription)
    
    ' 设置处方ID字段样式
    Call UIUtils.SetFieldDisplayFormat(Me.txtPrescription_ID, "Prescription_ID")
    
    ' 设置其他字段样式
    Call UIUtils.SetFieldDisplayFormat(Me.txtPatient_Name, "Other_Fields")
    Call UIUtils.SetFieldDisplayFormat(Me.txtPatient_Gender, "Other_Fields")
    Call UIUtils.SetFieldDisplayFormat(Me.txtPatient_Age, "Patient_Age")
    Call UIUtils.SetFieldDisplayFormat(Me.txtClinical_Diagnosis, "Other_Fields")
    Call UIUtils.SetFieldDisplayFormat(Me.txtFormula_Name, "Other_Fields")
    Call UIUtils.SetFieldDisplayFormat(Me.txtFormula_Composition, "Other_Fields")
    Call UIUtils.SetFieldDisplayFormat(Me.txtPrescription_Date, "Other_Fields")
    Call UIUtils.SetFieldDisplayFormat(Me.txtPrescription_Amount, "Prescription_Amount")
    
    ' 设置分隔线样式
    Call UIUtils.SetSectionDivider(Me.lineSection1)
    Call UIUtils.SetSectionDivider(Me.lineSection2)
    
    ' 设置字段标签样式
    SetFieldLabelStyles
    
    Exit Sub
    
ErrorHandler:
    ' 简化处理：使用默认样式
    Resume Next
End Sub

' 设置字段标签样式
Private Sub SetFieldLabelStyles()
    On Error Resume Next
    
    Dim ctrl As Control
    
    For Each ctrl In Me.Detail.Controls
        If TypeOf ctrl Is Label Then
            With ctrl
                .FontName = UIUtils.FONT_SIMSUN
                .FontSize = 10
                .FontBold = True
                .ForeColor = UIUtils.COLOR_BLACK
            End With
        End If
    Next ctrl
End Sub

' 报表页脚格式化事件
Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)
    On Error GoTo ErrorHandler
    
    ' 计算处方数量和总金额
    Dim prescriptionCount As Long
    Dim totalAmount As Currency
    
    prescriptionCount = GetPrescriptionCount
    totalAmount = GetTotalAmount
    
    ' 设置统计信息
    Me.lblPrescriptionCount.Caption = "总计：" & prescriptionCount & "处方"
    Me.lblTotalAmount.Caption = "累计处方金额：" & Format(totalAmount, "¥0.00元")
    
    ' 设置统计信息样式
    Call UIUtils.SetReportSummaryStyle(Me.lblPrescriptionCount)
    Call UIUtils.SetReportSummaryStyle(Me.lblTotalAmount)
    
    ' 设置报表日期
    Me.lblReportDate.Caption = "报表生成日期：" & Format(Date, "yyyy年mm月dd日")
    Me.lblReportDate.FontName = UIUtils.FONT_SIMSUN
    Me.lblReportDate.FontSize = 10
    Me.lblReportDate.ForeColor = UIUtils.COLOR_BLACK
    
    Exit Sub
    
ErrorHandler:
    Me.lblPrescriptionCount.Caption = "总计：0处方"
    Me.lblTotalAmount.Caption = "累计处方金额：¥0.00元"
End Sub

' 获取当前筛选条件下的处方数量
Private Function GetPrescriptionCount() As Long
    On Error GoTo ErrorHandler
    
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' 构建SQL查询，考虑筛选条件
    sql = "SELECT COUNT(*) AS PrescriptionCount FROM TCM_Prescription_Table"
    
    ' 添加筛选条件
    If Me.FilterOn And Me.Filter <> "" Then
        ' 从Filter中提取条件
        sql = sql & " WHERE " & Me.Filter
    End If
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        GetPrescriptionCount = Nz(rs!PrescriptionCount, 0)
    Else
        GetPrescriptionCount = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
ErrorHandler:
    GetPrescriptionCount = 0
End Function

' 获取当前筛选条件下的总金额
Private Function GetTotalAmount() As Currency
    On Error GoTo ErrorHandler
    
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' 构建SQL查询，考虑筛选条件
    sql = "SELECT SUM(Prescription_Amount) AS TotalAmount FROM TCM_Prescription_Table"
    
    ' 添加筛选条件
    If Me.FilterOn And Me.Filter <> "" Then
        ' 从Filter中提取条件
        sql = sql & " WHERE " & Me.Filter
    End If
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        GetTotalAmount = Nz(rs!TotalAmount, 0)
    Else
        GetTotalAmount = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
ErrorHandler:
    GetTotalAmount = 0
End Function

' 报表激活事件
Private Sub Report_Activate()
    On Error Resume Next
    
    ' 更新报表标题和统计信息
    Report_Open False
End Sub

' 报表关闭事件
Private Sub Report_Close()
    On Error Resume Next
    
    ' 清理资源
    Me.Filter = ""
    Me.FilterOn = False
End Sub

' 页面页眉格式化事件
Private Sub PageHeader_Format(Cancel As Integer, FormatCount As Integer)
    On Error Resume Next
    
    ' 设置页面信息
    Me.lblPageInfo.Caption = "第 " & Me.Page & " 页"
    Me.lblPageInfo.FontName = UIUtils.FONT_SIMSUN
    Me.lblPageInfo.FontSize = 10
    Me.lblPageInfo.ForeColor = UIUtils.COLOR_BLACK
    Me.lblPageInfo.TextAlign = 2 ' 居中
End Sub

' 报表无数据事件
Private Sub Report_NoData(Cancel As Integer)
    On Error Resume Next
    
    ' 如果没有数据，显示提示信息
    Dim response As Integer
    
    response = MsgBox("没有找到处方数据。是否显示所有医师的处方？", vbQuestion + vbYesNo, "无数据")
    
    If response = vbYes Then
        ' 移除筛选条件
        Me.Filter = ""
        Me.FilterOn = False
        SetReportTitle "所有医师", ""
        Cancel = False ' 继续显示报表
    Else
        Cancel = True ' 取消报表
        MsgBox "报表已取消。", vbInformation, "取消"
    End If
End Sub

' 设置打印选项
Public Sub SetPrintOptions()
    On Error GoTo ErrorHandler
    
    ' 设置报表打印属性
    With Me
        .PrintSection = True
        .MoveLayout = True
        .NextRecord = False
        .PrintCount = 1
    End With
    
    ' 设置页面设置
    With Me
        .LeftMargin = 1440 ' 1英寸
        .RightMargin = 1440
        .TopMargin = 1440
        .BottomMargin = 1440
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置打印选项错误：" & Err.Description, vbCritical, "错误"
End Sub

' 导出报表为PDF
Public Sub ExportToPDF()
    On Error GoTo ErrorHandler
    
    Dim exportPath As String
    Dim physicianName As String
    
    ' 获取医师姓名（从标题中提取）
    physicianName = Replace(Me.lblReportTitle.Caption, " 处方列表", "")
    physicianName = Replace(physicianName, " 主任医师", "")
    physicianName = Replace(physicianName, " 副主任医师", "")
    physicianName = Replace(physicianName, " 主治医师", "")
    
    exportPath = CurrentProject.Path & "\" & physicianName & "_处方列表_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    
    ' 导出为PDF
    DoCmd.OutputTo acOutputReport, Me.Name, acFormatPDF, exportPath
    
    MsgBox "报表已导出为PDF：" & vbCrLf & exportPath, vbInformation, "导出完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "导出PDF错误：" & Err.Description, vbCritical, "错误"
End Sub

' 预览报表（公开方法）
Public Sub PreviewReport(physicianID As String, Optional physicianName As String = "", Optional physicianTitle As String = "")
    On Error GoTo ErrorHandler
    
    If physicianID = "" Then
        MsgBox "请提供医师ID。", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 构建OpenArgs参数
    Dim openArgs As String
    openArgs = physicianID & "|" & physicianName & "|" & physicianTitle
    
    ' 打开报表预览
    DoCmd.OpenReport Me.Name, acViewPreview, , , , openArgs
    
    Exit Sub
    
ErrorHandler:
    MsgBox "预览报表错误：" & Err.Description, vbCritical, "错误"
End Sub

' 打印报表（直接打印）
Public Sub PrintReport(physicianID As String, Optional copies As Integer = 1)
    On Error GoTo ErrorHandler
    
    If physicianID = "" Then
        MsgBox "请提供医师ID。", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 设置筛选条件
    Me.Filter = "Physician_ID = '" & physicianID & "'"
    Me.FilterOn = True
    
    ' 打印报表
    DoCmd.OpenReport Me.Name, acViewNormal
    
    ' 设置打印份数
    ' 注意：Access的打印份数设置有限制，可能需要使用API
    ' 这里简化处理
    
    Exit Sub
    
ErrorHandler:
    MsgBox "打印报表错误：" & Err.Description, vbCritical, "错误"
End Sub