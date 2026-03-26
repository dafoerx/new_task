Attribute VB_Name = "DatabaseUtils"
' 中医处方管理数据库 - 数据库工具模块
' 提供常用的数据库操作函数

Option Explicit

' 执行SQL语句
Public Function ExecuteSQL(sqlStatement As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    db.Execute sqlStatement, dbFailOnError
    
    ExecuteSQL = True
    Exit Function
    
ErrorHandler:
    MsgBox "执行SQL错误：" & Err.Description & vbCrLf & "SQL: " & sqlStatement, vbCritical, "错误"
    ExecuteSQL = False
End Function

' 执行查询并返回记录集
Public Function GetRecordset(sqlStatement As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sqlStatement, dbOpenDynaset)
    
    Set GetRecordset = rs
    Exit Function
    
ErrorHandler:
    MsgBox "获取记录集错误：" & Err.Description, vbCritical, "错误"
    Set GetRecordset = Nothing
End Function

' 获取表中记录数量
Public Function GetRecordCount(tableName As String, Optional whereClause As String = "") As Long
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb
    
    sql = "SELECT COUNT(*) AS RecordCount FROM " & tableName
    
    If whereClause <> "" Then
        sql = sql & " WHERE " & whereClause
    End If
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        GetRecordCount = rs!RecordCount
    Else
        GetRecordCount = 0
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErrorHandler:
    GetRecordCount = -1
End Function

' 清空表数据
Public Function ClearTable(tableName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim response As Integer
    
    response = MsgBox("确定要清空表 '" & tableName & "' 的所有数据吗？", vbQuestion + vbYesNo, "确认")
    
    If response = vbYes Then
        ExecuteSQL "DELETE FROM " & tableName
        ClearTable = True
    Else
        ClearTable = False
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "清空表错误：" & Err.Description, vbCritical, "错误"
    ClearTable = False
End Function

' 导入文本文件到表
Public Function ImportTextFile(filePath As String, tableName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim specName As String
    
    ' 使用导入规范（需要预先创建）
    specName = "TCM_Prescription_Import_Spec"
    
    DoCmd.TransferText acImportDelim, specName, tableName, filePath, True
    
    ImportTextFile = True
    Exit Function
    
ErrorHandler:
    MsgBox "导入文本文件错误：" & Err.Description, vbCritical, "错误"
    ImportTextFile = False
End Function

' 导出表到文本文件
Public Function ExportTableToText(tableName As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    DoCmd.TransferText acExportDelim, , tableName, filePath, True
    
    ExportTableToText = True
    Exit Function
    
ErrorHandler:
    MsgBox "导出表错误：" & Err.Description, vbCritical, "错误"
    ExportTableToText = False
End Function

' 备份数据库
Public Sub BackupDatabase(backupPath As String)
    On Error GoTo ErrorHandler
    
    Dim dbPath As String
    Dim fso As Object
    
    dbPath = CurrentDb.Name
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(dbPath) Then
        fso.CopyFile dbPath, backupPath, True
        MsgBox "数据库备份已保存到：" & backupPath, vbInformation, "备份完成"
    Else
        MsgBox "数据库文件不存在。", vbExclamation, "错误"
    End If
    
    Set fso = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "备份数据库错误：" & Err.Description, vbCritical, "错误"
End Sub

' 验证数据完整性
Public Function ValidateDataIntegrity() As Boolean
    On Error GoTo ErrorHandler
    
    Dim issues As String
    Dim hasIssues As Boolean
    
    hasIssues = False
    issues = "数据完整性检查结果：" & vbCrLf & vbCrLf
    
    ' 检查外键约束
    ' 1. 检查处方表中的医师ID是否都在医师表中存在
    Dim sql As String
    Dim rs As DAO.Recordset
    
    sql = "SELECT P.Prescription_ID, P.Physician_ID " & _
          "FROM TCM_Prescription_Table AS P " & _
          "LEFT JOIN Doctor_Table AS D ON P.Physician_ID = D.Physician_ID " & _
          "WHERE D.Physician_ID IS NULL"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        issues = issues & "发现无效的医师ID（在处方表中存在但在医师表中不存在）：" & vbCrLf
        Do While Not rs.EOF
            issues = issues & "  处方ID: " & rs!Prescription_ID & ", 医师ID: " & rs!Physician_ID & vbCrLf
            rs.MoveNext
        Loop
        issues = issues & vbCrLf
        hasIssues = True
    End If
    
    rs.Close
    
    ' 2. 检查医师表中的科室ID是否都在科室表中存在
    sql = "SELECT D.Physician_ID, D.Department_ID " & _
          "FROM Doctor_Table AS D " & _
          "LEFT JOIN Department_Table AS Dept ON D.Department_ID = Dept.Department_ID " & _
          "WHERE Dept.Department_ID IS NULL"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        issues = issues & "发现无效的科室ID（在医师表中存在但在科室表中不存在）：" & vbCrLf
        Do While Not rs.EOF
            issues = issues & "  医师ID: " & rs!Physician_ID & ", 科室ID: " & rs!Department_ID & vbCrLf
            rs.MoveNext
        Loop
        issues = issues & vbCrLf
        hasIssues = True
    End If
    
    rs.Close
    Set rs = Nothing
    
    If hasIssues Then
        MsgBox issues, vbExclamation, "数据完整性警告"
        ValidateDataIntegrity = False
    Else
        MsgBox "数据完整性检查通过，未发现问题。", vbInformation, "检查完成"
        ValidateDataIntegrity = True
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "验证数据完整性错误：" & Err.Description, vbCritical, "错误"
    ValidateDataIntegrity = False
End Function

' 获取数据库信息
Public Sub ShowDatabaseInfo()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    Dim info As String
    Dim tableCount As Integer, queryCount As Integer
    
    Set db = CurrentDb
    
    tableCount = 0
    queryCount = 0
    
    info = "数据库信息：" & vbCrLf & vbCrLf
    info = info & "数据库路径：" & db.Name & vbCrLf
    info = info & "创建日期：" & Format(FileDateTime(db.Name), "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
    
    info = info & "表列表：" & vbCrLf
    For Each tdf In db.TableDefs
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 4) <> "~TMP" Then
            info = info & "  • " & tdf.Name & " (" & tdf.RecordCount & " 条记录)" & vbCrLf
            tableCount = tableCount + 1
        End If
    Next tdf
    
    info = info & vbCrLf & "查询列表：" & vbCrLf
    For Each qdf In db.QueryDefs
        If Left(qdf.Name, 4) <> "~TMP" Then
            info = info & "  • " & qdf.Name & vbCrLf
            queryCount = queryCount + 1
        End If
    Next qdf
    
    info = info & vbCrLf & "统计：" & vbCrLf
    info = info & "  表数量：" & tableCount & vbCrLf
    info = info & "  查询数量：" & queryCount & vbCrLf
    
    MsgBox info, vbInformation, "数据库信息"
    
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "获取数据库信息错误：" & Err.Description, vbCritical, "错误"
End Sub