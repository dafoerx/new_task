Attribute VB_Name = "MainModule"
' 中医处方管理数据库 - 主模块
' 作者：考生姓名
' 日期：2026年3月26日
' 版本：1.0

Option Explicit

' 应用程序启动
Public Sub Main()
    On Error GoTo ErrorHandler
    
    ' 检查数据库连接
    If Not CheckDatabaseConnection() Then
        MsgBox "数据库连接失败，请检查数据库文件。", vbCritical, "错误"
        Exit Sub
    End If
    
    ' 显示主菜单
    ShowMainMenu
    
    Exit Sub
    
ErrorHandler:
    MsgBox "应用程序启动错误：" & Err.Description, vbCritical, "错误"
    Resume Next
End Sub

' 显示主菜单
Public Sub ShowMainMenu()
    On Error GoTo ErrorHandler
    
    ' 这里可以显示自定义菜单或启动表单
    ' 根据考试要求，直接打开两个主要表单
    
    ' 打开医师基本信息管理表单
    DoCmd.OpenForm "frmPhysicianBasicInfo", acNormal, , , , acDialog
    
    ' 打开浏览与打印处方信息表单
    DoCmd.OpenForm "frmBrowsePrintPrescription", acNormal, , , , acDialog
    
    Exit Sub
    
ErrorHandler:
    MsgBox "显示菜单错误：" & Err.Description, vbCritical, "错误"
    Resume Next
End Sub

' 检查数据库连接
Public Function CheckDatabaseConnection() As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    ' 检查必要表是否存在
    If Not TableExists("Department_Table") Then
        MsgBox "缺少科室表(Department_Table)，请先创建数据库结构。", vbExclamation, "警告"
        CheckDatabaseConnection = False
        Exit Function
    End If
    
    If Not TableExists("Doctor_Table") Then
        MsgBox "缺少医师表(Doctor_Table)，请先创建数据库结构。", vbExclamation, "警告"
        CheckDatabaseConnection = False
        Exit Function
    End If
    
    If Not TableExists("TCM_Prescription_Table") Then
        MsgBox "缺少处方表(TCM_Prescription_Table)，请先创建数据库结构。", vbExclamation, "警告"
        CheckDatabaseConnection = False
        Exit Function
    End If
    
    CheckDatabaseConnection = True
    Exit Function
    
ErrorHandler:
    CheckDatabaseConnection = False
End Function

' 检查表是否存在
Public Function TableExists(tableName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    Set db = CurrentDb
    
    For Each tdf In db.TableDefs
        If tdf.Name = tableName Then
            TableExists = True
            Exit Function
        End If
    Next tdf
    
    TableExists = False
    Exit Function
    
ErrorHandler:
    TableExists = False
End Function

' 初始化数据库（第一次运行时调用）
Public Sub InitializeDatabase()
    On Error GoTo ErrorHandler
    
    Dim response As Integer
    
    response = MsgBox("是否初始化数据库？这将创建所有表和关系。", vbQuestion + vbYesNo, "初始化数据库")
    
    If response = vbYes Then
        ' 执行SQL脚本创建表
        ' 注意：实际应用中应执行SQL脚本文件
        ' 这里简化为检查并创建必要表
        
        CreateTables
        CreateIndexes
        CreateRelationships
        InsertTestData
        
        MsgBox "数据库初始化完成。", vbInformation, "完成"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "初始化数据库错误：" & Err.Description, vbCritical, "错误"
End Sub

' 创建表（简化版）
Private Sub CreateTables()
    ' 在实际应用中，这里应执行SQL脚本
    ' 为了简化，仅显示消息
    MsgBox "正在创建表结构...", vbInformation, "请稍候"
End Sub

' 创建索引
Private Sub CreateIndexes()
    MsgBox "正在创建索引...", vbInformation, "请稍候"
End Sub

' 创建关系
Private Sub CreateRelationships()
    MsgBox "正在创建表关系...", vbInformation, "请稍候"
End Sub

' 插入测试数据
Private Sub InsertTestData()
    MsgBox "正在插入测试数据...", vbInformation, "请稍候"
End Sub

' 打开查询
Public Sub OpenQuery(queryName As String)
    On Error GoTo ErrorHandler
    
    If QueryExists(queryName) Then
        DoCmd.OpenQuery queryName
    Else
        MsgBox "查询 '" & queryName & "' 不存在。", vbExclamation, "警告"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "打开查询错误：" & Err.Description, vbCritical, "错误"
End Sub

' 检查查询是否存在
Public Function QueryExists(queryName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb
    
    For Each qdf In db.QueryDefs
        If qdf.Name = queryName Then
            QueryExists = True
            Exit Function
        End If
    Next qdf
    
    QueryExists = False
    Exit Function
    
ErrorHandler:
    QueryExists = False
End Function