Attribute VB_Name = "VibramSalesSQLWorkflow Mod"
Option Compare Database
Option Explicit

Public Function VibramSalesSQLWorkflowCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Function ReplaceDatesInSQL(sql, ByVal newFirstDate As String, Optional ByVal newSecondDate = Null) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Global = False
    regex.IgnoreCase = True
    regex.pattern = "\#\d{1,2}\/\d{1,2}\/\d{4}\#" ' Matches date in the format #m/d/yyyy#
    
    ' Find the first match and replace it with newFirstDate
    Dim firstMatch As Object
    Set firstMatch = regex.Execute(sql)
    If firstMatch.count > 0 Then
        sql = Left$(sql, firstMatch.item(0).FirstIndex) & "#" & newFirstDate & "#" & Mid$(sql, firstMatch.item(0).FirstIndex + firstMatch.item(0).length + 1)
    End If
    
    If isFalse(newSecondDate) Then
        ReplaceDatesInSQL = sql
        Exit Function
    End If
    
    ' Reset the regex object to find the last match
    regex.Global = True
    regex.pattern = "\#\d{1,2}\/\d{1,2}\/\d{4}\#"
    
    ' Find the last match and replace it with newSecondDate
    Dim allMatches As Object
    Set allMatches = regex.Execute(sql)
    If allMatches.count > 0 Then
        Dim lastMatch As Object
        Set lastMatch = allMatches.item(allMatches.count - 1)
        sql = Left$(sql, lastMatch.FirstIndex) & "#" & newSecondDate & "#" & Mid$(sql, lastMatch.FirstIndex + lastMatch.length + 1)
    End If
    
    ReplaceDatesInSQL = sql
End Function

Sub TestReplaceDates()
    Dim sql As String
    sql = "SELECT [Sales-Main].* FROM [Sales-Main] WHERE ((([Sales-Main].Sales_Date) Between #1/1/2024# And #1/31/2024#));"
    
    ' Replace the first date with "2/1/2024" and the second date with "2/28/2024"
    Debug.Print ReplaceDatesInSQL(sql, "2/1/2024", "2/28/2024")
End Sub

Public Function RunVibramSalesDataSQLWorkflow(frm As Form)

    Dim VibramSalesSQLWorkflowID: VibramSalesSQLWorkflowID = frm("VibramSalesSQLWorkflowID")
    Dim StartDate: StartDate = frm("StartDate"): If ExitIfTrue(isFalse(StartDate), "Start Date is empty..") Then Exit Function
    Dim EndDate: EndDate = frm("EndDate"): If ExitIfTrue(isFalse(EndDate), "End Date is empty..") Then Exit Function
    Dim SourceDatabaseFile: SourceDatabaseFile = frm("SourceDatabaseFile"): If ExitIfTrue(isFalse(SourceDatabaseFile), "Source Database File is empty..") Then Exit Function
    Dim SkipMode: SkipMode = frm("SkipMode")
    
    Run_UpdatetblPaymentTotals SourceDatabaseFile
    ''Modify the queries first.
    Call ModifyQueries(SourceDatabaseFile, StartDate, EndDate)
    
    AppendTextToFile GenerateDeleteStatements(StartDate, EndDate), CurrentProject.path & "\vibram-sales.sql", True
    
    GenerateAndAppendSqlStatements SourceDatabaseFile, SkipMode
    
    AppendTextToFile GenerateUpdateSalesQuery(), CurrentProject.path & "\vibram-sales.sql"
    AppendTextToFile GenerateUpdateReturnSlipsQuery(), CurrentProject.path & "\vibram-sales.sql"
    
    OpenFolderLocation CurrentProject.path & "\"
    
End Function

Private Function GenerateUpdateSalesQuery() As String
    Dim sqlQuery As String
    sqlQuery = "UPDATE vibram_sales.sales " & _
               "SET gross_sales = sub.gross_sales, " & _
               "    discount = sub.discount " & _
               "FROM (" & _
               "    SELECT dr_number, SUM(gross_sales) AS gross_sales, SUM(total_discount) AS discount " & _
               "    FROM vibram_sales.sales_items_view " & _
               "    GROUP BY dr_number" & _
               ") AS sub " & _
               "WHERE vibram_sales.sales.dr_number = sub.dr_number;"
    GenerateUpdateSalesQuery = sqlQuery
End Function

Private Function GenerateUpdateReturnSlipsQuery() As String
    Dim sqlQuery As String
    sqlQuery = "UPDATE vibram_sales.return_slips" & _
               " SET gross_amount = sub.gross_amount," & _
               "    discount = sub.discount" & _
               " FROM (" & _
               "    SELECT slip_no, SUM(gross_amount) AS gross_amount, SUM(total_discount) AS discount " & _
               "    FROM vibram_sales.return_slip_items_view " & _
               "    GROUP BY slip_no" & _
               " ) AS sub " & _
               " WHERE vibram_sales.return_slips.slip_no = sub.slip_no;"
    GenerateUpdateReturnSlipsQuery = sqlQuery
End Function

Private Function GenerateDeleteStatements(StartDate, EndDate) As String
    Dim deleteSalesStatement As String
    deleteSalesStatement = "DELETE FROM vibram_sales.sales WHERE sales_date between '" & Format(StartDate, "YYYY-MM-DD") & "' AND '" & Format(EndDate, "YYYY-MM-DD") & "';"
    
    Dim deletePaymentsStatement As String
    deletePaymentsStatement = "DELETE FROM vibram_sales.payments WHERE collection_date between '" & Format(StartDate, "YYYY-MM-DD") & "' AND '" & Format(EndDate, "YYYY-MM-DD") & "';"
    
    GenerateDeleteStatements = deleteSalesStatement & vbCrLf & deletePaymentsStatement
End Function

Private Sub GenerateAndAppendSqlStatements(SourceDatabaseFile, SkipMode)

    Dim frm As Form
    Dim sqlStr: sqlStr = "SELECT * FROM tblVibramSalesSQLWorkflowItems"
    If SkipMode Then
        sqlStr = sqlStr & " WHERE Not Skippable"
    End If
    sqlStr = sqlStr & " ORDER BY StepOrder"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Sub
        DoCmd.OpenForm "frmCsvInsertStatements", , , "TableName = " & Esc(TableName) & " AND SourceDatabaseFile = " & Esc(SourceDatabaseFile)
        Set frm = Forms("frmCsvInsertStatements")
        Dim CsvInsertStatementID: CsvInsertStatementID = frm("CsvInsertStatementID")
        Dim Operations: Operations = rs.fields("Operation"): If ExitIfTrue(isFalse(Operations), "Operations is empty..") Then Exit Sub
        If Operations = "Insert" Or Operations = "Both" Then
            GenerateAndAppendSqlStatementsForType "Insert", frm, CsvInsertStatementID, TableName
        End If

        If Operations = "Update" Or Operations = "Both" Then
            GenerateAndAppendSqlStatementsForType "Update", frm, CsvInsertStatementID, TableName
        End If
        
        Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Sub
        If isPresent("tblSeqModelFields", "Autoincrement And PrimaryKey AND SeqModelID = " & SeqModelID) Then
            AppendTextToFile GetResetSerialAutonumber(frm, SeqModelID), CurrentProject.path & "\vibram-sales.sql"
        End If
        
        rs.MoveNext
    Loop
End Sub

Private Sub GenerateAndAppendSqlStatementsForType(operationType, frm As Form, CsvInsertStatementID, TableName)
    Dim rs As Recordset, i As Integer
    Dim sqlOperation As String

    Select Case operationType
        Case "Insert"
            GenerateInsertSqlStatementsForPostgres frm
            sqlOperation = "INSERT"
        Case "Update"
            GenerateUpdateSqlStatementsForPostgres frm
            sqlOperation = "UPDATE"
        Case Else
            Exit Sub
    End Select

    Set rs = ReturnRecordset("SELECT * FROM tblCsvInsertStatementItems WHERE CsvInsertStatementID = " & CsvInsertStatementID & " ORDER BY CsvInsertStatementItemID")
    i = 1
    Do Until rs.EOF
        AppendTextToFile "--" & TableName & "--" & sqlOperation & "#" & i & vbNewLine & rs.fields("SqlStatement"), CurrentProject.path & "\vibram-sales.sql"
        i = i + 1
        rs.MoveNext
    Loop
End Sub

Private Sub ModifyQueries(SourceDatabaseFile, StartDate, EndDate)
    ''Modify the queries first.
    ''qryFilteredSalesMain
    ModifyTargetQuery SourceDatabaseFile, "qryFilteredSalesMain", StartDate, EndDate
    ''qryPaymentsForExportFiltered
    ModifyTargetQuery SourceDatabaseFile, "qryPaymentsForExportFiltered", StartDate, EndDate
    ''qryFilteredSalesMainEarlierThanFiltered
    ModifyTargetQuery SourceDatabaseFile, "qryFilteredSalesMainEarlierThanFiltered", StartDate
End Sub

Public Function Run_UpdatetblPaymentTotals(SourceDatabaseFile)

    Dim appAccess As Access.Application
    ' Create instance of Access Application object.
    Set appAccess = CreateObject("Access.Application")
    appAccess.Visible = False
    ' Open WizCode database in Microsoft Access window.
    appAccess.OpenCurrentDatabase SourceDatabaseFile, False
    ' Run Sub procedure.
    appAccess.Run "UpdatetblPaymentTotals"
    appAccess.CloseCurrentDatabase
    appAccess.Quit acQuitSaveNone
    Set appAccess = Nothing
 
End Function

Public Function ModifyTargetQuery(SourceDatabaseFile, qryName, StartDate, Optional EndDate = Null)
    
    Dim db As DAO.Database, qDef As DAO.QueryDef
    
    Set db = OpenDatabase(SourceDatabaseFile)
    
    Set qDef = db.QueryDefs(qryName)
    
    Dim sqlStr: sqlStr = qDef.sql
    sqlStr = ReplaceDatesInSQL(sqlStr, StartDate, EndDate)
    qDef.sql = sqlStr
    qDef.Close
    db.Close
    
    Set qDef = Nothing
    Set db = Nothing
    
End Function
