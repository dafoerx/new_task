Attribute VB_Name = "UIUtils"
' 中医处方管理数据库 - UI工具模块
' 提供表单和报表的UI相关函数，严格遵循考试要求

Option Explicit

' 字体常量定义（按考试要求）
Public Const FONT_LISHU As String = "隶书"
Public Const FONT_STXINWEI As String = "STXinwei"
Public Const FONT_KAITI As String = "KaiTi"
Public Const FONT_SIMSUN As String = "SimSun"

' 颜色常量定义
Public Const COLOR_RED As Long = vbRed
Public Const COLOR_BLUE As Long = vbBlue
Public Const COLOR_BLACK As Long = vbBlack

' 设置表单为对话框样式（按考试要求）
Public Sub SetFormAsDialog(frm As Form)
    On Error GoTo ErrorHandler
    
    With frm
        ' 边框样式：对话框
        .BorderStyle = 3 ' acDialogBorder
        
        ' 无导航按钮、记录选择器、分割线、滚动条
        .NavigationButtons = False
        .RecordSelectors = False
        .DividingLines = False
        .ScrollBars = 0 ' acScrollBarsNone
        
        ' 控制框无最大化/最小化按钮
        .MinMaxButtons = 0 ' acMinMaxButtonNone
        .CloseButton = True ' 保留关闭按钮
        
        ' 弹出模式
        .Modal = True
        .PopUp = True
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置表单对话框样式错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置标签字体样式（按考试要求）
Public Sub SetLabelFontStyle(lbl As Label, fontName As String, fontSize As Integer, _
                            Optional isBold As Boolean = True, Optional textColor As Long = vbBlack, _
                            Optional isCentered As Boolean = True)
    On Error GoTo ErrorHandler
    
    With lbl
        .FontName = fontName
        .FontSize = fontSize
        .FontBold = isBold
        .ForeColor = textColor
        
        If isCentered Then
            .TextAlign = 2 ' acTextAlignCenter
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    ' 如果字体不存在，使用默认字体
    lbl.FontName = "宋体"
    MsgBox "字体 '" & fontName & "' 可能未安装，已使用默认字体。", vbExclamation, "字体警告"
End Sub

' 设置文本框边框样式
Public Sub SetTextBoxBorder(txt As TextBox, borderColor As Long, borderWidth As Integer, _
                           Optional borderStyle As Integer = 1) ' 1=实线
    On Error Resume Next ' 某些控件可能不支持所有边框属性
    
    With txt
        .BorderColor = borderColor
        .BorderWidth = borderWidth
        .BorderStyle = borderStyle
    End With
End Sub

' 设置细节区域边框（按考试要求：红色实线边框，宽度3pt）
Public Sub SetDetailSectionBorder(frm As Form, section As Section)
    On Error GoTo ErrorHandler
    
    With section
        .BorderColor = COLOR_RED
        .BorderWidth = 3 ' 3pt
        .BorderStyle = 1 ' 实线
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置细节区域边框错误：" & Err.Description, vbCritical, "错误"
End Sub

' 创建导航按钮（按考试要求：5个按钮）
Public Sub CreateNavigationButtons(frm As Form, buttonsContainer As Control)
    On Error GoTo ErrorHandler
    
    Dim btnFirst As CommandButton
    Dim btnPrevious As CommandButton
    Dim btnNext As CommandButton
    Dim btnLast As CommandButton
    Dim btnClose As CommandButton
    
    ' 创建按钮（在实际表单设计中应通过设计器创建）
    ' 这里提供创建代码和事件处理
    
    ' 设置按钮属性
    SetButtonProperties btnFirst, "首记录", 100, 100
    SetButtonProperties btnPrevious, "上一条", 200, 100
    SetButtonProperties btnNext, "下一条", 300, 100
    SetButtonProperties btnLast, "末记录", 400, 100
    SetButtonProperties btnClose, "关闭表单", 500, 100
    
    ' 绑定事件（需要在类模块中实现）
    
    Exit Sub
    
ErrorHandler:
    MsgBox "创建导航按钮错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置按钮属性
Private Sub SetButtonProperties(btn As CommandButton, caption As String, left As Integer, top As Integer)
    On Error Resume Next
    
    With btn
        .Caption = caption
        .Left = left
        .Top = top
        .Width = 1200
        .Height = 400
        .FontName = "宋体"
        .FontSize = 10
    End With
End Sub

' 设置子窗体显示字段（按考试要求：仅显示5个字段）
Public Sub SetSubformFields(subfrm As SubForm, fieldNames As Variant)
    On Error GoTo ErrorHandler
    
    ' fieldNames应包含5个字段名的数组
    ' 例如：Array("Patient_Name", "Clinical_Diagnosis", "Formula_Name", "Prescription_Date", "Prescription_Amount")
    
    ' 在实际应用中，这通常通过设置子窗体的源对象和链接字段实现
    ' 这里提供概念性代码
    
    Dim i As Integer
    Dim fieldList As String
    
    ' 构建字段列表
    For i = LBound(fieldNames) To UBound(fieldNames)
        If i > LBound(fieldNames) Then fieldList = fieldList & ", "
        fieldList = fieldList & fieldNames(i)
    Next i
    
    ' 设置子窗体数据源（示例）
    ' subfrm.SourceObject = "查询.仅五字段查询"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置子窗体字段错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置报表标题样式（按考试要求：宋体24号加粗红色）
Public Sub SetReportTitleStyle(rpt As Report, titleLabel As Label)
    On Error GoTo ErrorHandler
    
    SetLabelFontStyle titleLabel, FONT_SIMSUN, 24, True, COLOR_RED, True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置报表标题样式错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置处方信息框样式（按考试要求：蓝色边框，宽度3pt）
Public Sub SetPrescriptionBoxStyle(rpt As Report, boxControl As Control)
    On Error GoTo ErrorHandler
    
    With boxControl
        .BorderColor = COLOR_BLUE
        .BorderWidth = 3
        .BorderStyle = 1 ' 实线
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置处方信息框样式错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置字段显示格式
Public Sub SetFieldDisplayFormat(fld As Control, fieldType As String)
    On Error GoTo ErrorHandler
    
    Select Case fieldType
        Case "Patient_Age"
            ' 显示为"X岁"
            fld.ControlSource = "Format([Patient_Age], '0') & '岁'"
            fld.FontName = FONT_SIMSUN
            fld.FontSize = 12
            fld.ForeColor = COLOR_BLACK
            
        Case "Prescription_Amount"
            ' 显示为"¥X元"
            fld.ControlSource = "'¥' & Format([Prescription_Amount], '0.00') & '元'"
            fld.FontName = FONT_SIMSUN
            fld.FontSize = 12
            fld.ForeColor = COLOR_BLACK
            
        Case "Prescription_ID"
            ' 处方ID：宋体18号黑色
            fld.FontName = FONT_SIMSUN
            fld.FontSize = 18
            fld.ForeColor = COLOR_BLACK
            
        Case "Other_Fields"
            ' 其他字段：宋体12号黑色
            fld.FontName = FONT_SIMSUN
            fld.FontSize = 12
            fld.ForeColor = COLOR_BLACK
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置字段显示格式错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置区域分隔线（按考试要求：点划线，宽度1pt，黑色）
Public Sub SetSectionDivider(lineControl As Line)
    On Error GoTo ErrorHandler
    
    With lineControl
        .BorderStyle = 3 ' acBorderStyleDashDot
        .BorderWidth = 1
        .BorderColor = COLOR_BLACK
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置区域分隔线错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置报表统计信息样式（按考试要求：宋体14号加粗红色）
Public Sub SetReportSummaryStyle(lbl As Label)
    On Error GoTo ErrorHandler
    
    SetLabelFontStyle lbl, FONT_SIMSUN, 14, True, COLOR_RED, True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置报表统计信息样式错误：" & Err.Description, vbCritical, "错误"
End Sub

' 验证字体是否安装
Public Function IsFontInstalled(fontName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim screenFont As Screen
    Set screenFont = Screen
    
    ' 尝试设置字体名称，如果错误则字体不存在
    screenFont.FontName = fontName
    
    IsFontInstalled = True
    Exit Function
    
ErrorHandler:
    IsFontInstalled = False
End Function

' 检查所有必需字体
Public Sub CheckRequiredFonts()
    On Error GoTo ErrorHandler
    
    Dim missingFonts As String
    missingFonts = ""
    
    If Not IsFontInstalled(FONT_LISHU) Then
        missingFonts = missingFonts & "• 隶书 (LiShu)" & vbCrLf
    End If
    
    If Not IsFontInstalled(FONT_STXINWEI) Then
        missingFonts = missingFonts & "• STXinwei" & vbCrLf
    End If
    
    If Not IsFontInstalled(FONT_KAITI) Then
        missingFonts = missingFonts & "• 楷体 (KaiTi)" & vbCrLf
    End If
    
    If Not IsFontInstalled(FONT_SIMSUN) Then
        missingFonts = missingFonts & "• 宋体 (SimSun)" & vbCrLf
    End If
    
    If missingFonts <> "" Then
        MsgBox "以下字体未安装，可能影响界面显示效果：" & vbCrLf & vbCrLf & _
               missingFonts & vbCrLf & _
               "请安装所需字体或使用字体替代方案。", vbExclamation, "字体检查"
    Else
        MsgBox "所有必需字体均已安装。", vbInformation, "字体检查"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "检查字体错误：" & Err.Description, vbCritical, "错误"
End Sub

' 设置控件提示文本
Public Sub SetControlTooltip(ctrl As Control, tooltipText As String)
    On Error Resume Next ' 某些Access版本可能不支持ControlTipText
    
    ctrl.ControlTipText = tooltipText
End Sub

' 高亮显示必填字段
Public Sub HighlightRequiredField(txt As TextBox, isRequired As Boolean)
    On Error Resume Next
    
    If isRequired Then
        txt.BackColor = 13434879 ' 浅黄色
        SetControlTooltip txt, "此字段为必填项"
    Else
        txt.BackColor = vbWhite
        txt.ControlTipText = ""
    End If
End Sub

' 启用/禁用控件组
Public Sub EnableControlGroup(ctrlArray As Variant, enabled As Boolean)
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    
    For i = LBound(ctrlArray) To UBound(ctrlArray)
        ctrlArray(i).Enabled = enabled
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "启用/禁用控件组错误：" & Err.Description, vbCritical, "错误"
End Sub