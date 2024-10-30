Attribute VB_Name = "CsvInsertStatement Mod"
Option Compare Database
Option Explicit

Public Function CsvInsertStatementCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function CopyInsertSqlToClipboard(frm As Form)
    
    Dim Start: Start = frm("Start")
    Dim vEnd: vEnd = frm("End")
    
    Dim rs As Recordset: Set rs = frm("subCsvInsertStatementItems").Form.RecordsetClone
    rs.MoveFirst
    
    Dim lines As New clsArray
    
    Dim i As Integer: i = 1
    Do Until rs.EOF
        Dim SqlStatement: SqlStatement = rs.fields("SqlStatement")
        If i > vEnd Then
            Exit Do
        End If
        If i >= Start Then
            lines.Add SqlStatement
        End If
        i = i + 1
        rs.MoveNext
    Loop
    
    Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    If isPresent("tblSeqModelFields", "Autoincrement And PrimaryKey AND SeqModelID = " & SeqModelID) Then
        lines.Add GetResetSerialAutonumber(frm, SeqModelID)
    End If
    
    Dim Text: Text = lines.JoinArr(vbNewLine)
    DoCmd.OpenForm "frmClipboardForms"
    Forms("frmClipboardForms")("Snippet") = Text
    CopyFieldContent Forms("frmClipboardForms"), "Snippet"
    
    OpenSupabaseSqlEditorThenClose_frmClipboardForms frm, SeqModelID
    
End Function



Public Function GenerateUpdateSqlStatementsForPostgres(frm As Object, Optional CsvInsertStatementItemID = "")
    
    RunCommandSaveRecord
    
    Dim CsvInsertStatementID: CsvInsertStatementID = frm("CsvInsertStatementID")
    Dim Cursor
    If Not isFalse(CsvInsertStatementItemID) Then
        
        Cursor = ELookup("tblCsvInsertStatementItems", "CsvInsertStatementItemID = " & CsvInsertStatementItemID, "Cursor")
        
        Dim resp: resp = MsgBox("This will rewrite all the SQL after the cursor: " & Esc(Cursor) & ". Do you want to proceed?", vbYesNo)
        
        If resp = vbNo Then Exit Function
            
        RunSQL "DELETE FROM tblCsvInsertStatementItems WHERE CsvInsertStatementID = " & CsvInsertStatementID & _
            " AND CsvInsertStatementItemID >= " & CsvInsertStatementItemID
    Else
        RunSQL "DELETE FROM tblCsvInsertStatementItems WHERE CsvInsertStatementID = " & CsvInsertStatementID
    End If
    
    Dim SourceDatabaseFile: SourceDatabaseFile = frm("SourceDatabaseFile"): If ExitIfTrue(isFalse(SourceDatabaseFile), "SourceDatabaseFile is empty..") Then Exit Function
    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    Dim Limit: Limit = frm("Limit")
    Dim PrimaryKey: PrimaryKey = frm("PrimaryKey")
    Dim TableName: TableName = frm("TableName")
    Dim TimestampField: TimestampField = frm("TimestampField")
    Dim FetchLimit: FetchLimit = frm("FetchLimit")
    
    Dim TargetFields, UpdateStatements, DatabasePrimaryKey
    
    UpdateUpdateVariables SeqModelID, TableName, TargetFields, TimestampField, DatabasePrimaryKey
    
    GetUpdateValues TableName, TargetFields, PrimaryKey, Limit, CsvInsertStatementID, SeqModelID, TimestampField, FetchLimit, SourceDatabaseFile, Cursor, DatabasePrimaryKey
    ''Dim sqlStr: sqlStr = "INSERT INTO " & TableName & "(" & TargetFields & ") VALUES " & InsertStatements & ";"
    
    frm("subCsvInsertStatementItems").Form.Requery
    
End Function

Private Function GetUpdateValues(TableName, TargetFields, PrimaryKey, Limit, CsvInsertStatementID, SeqModelID, TimestampField, FetchLimit, SourceDatabaseFile, Cursor, DatabasePrimaryKey)
    
    Dim i As Integer: i = 0
    Dim j As Long: j = 0
    
    Dim sqlStr: sqlStr = "SELECT * FROM [" & TableName & "] IN " & Esc(SourceDatabaseFile)
    
    If Not isFalse(Cursor) Then
        Dim isPrimaryKeyAString: isPrimaryKeyAString = Not InStr(PrimaryKey, "ID") > 0
        sqlStr = sqlStr & " WHERE " & PrimaryKey & " > " & IIf(isPrimaryKeyAString, Esc(Cursor), Cursor)
    End If
    
    sqlStr = sqlStr & " ORDER BY " & PrimaryKey
    
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    Dim totalRecords: totalRecords = IIf(isFalse(FetchLimit), rs.recordCount, FetchLimit)
    
    
    Dim rows As New clsArray
    
    Dim SanitizedAppName: SanitizedAppName = ELookup("qrySeqModels", "SeqModelID = " & SeqModelID, "SanitizedAppName")
    Dim BeTableName: BeTableName = ELookup("qrySeqModels", "SeqModelID = " & SeqModelID, "TableName")
    
    Dim cls_ProgressBar As New cls_ProgressBar
    
    cls_ProgressBar.ProgressBar_Show
    cls_ProgressBar.ProgressBar_ProgressOverlay True
    cls_ProgressBar.ProgressBar_Caption "Filtering Data"
    ''cls_ProgressBar.ProgressBar_Resize 3000, 6000
    cls_ProgressBar.ProgressBar_Message1_FontParam 1, "Calibri", 100, 10, vbBlack
    cls_ProgressBar.ProgressBar_ProgressValue_Align 2
    
    Dim setFields
    Do Until rs.EOF
        
        setFields = GenerateRowValuesForUpdate(rs, SeqModelID, TimestampField)
        sqlStr = ReplaceDoubleQuote("UPDATE " & Esc(SanitizedAppName) & "." & Esc(BeTableName) & " SET " & setFields)
        rows.Add sqlStr
        i = i + 1
        j = j + 1
        
        If Not isFalse(FetchLimit) Then
            If j >= FetchLimit Then
                Exit Do
            End If
        End If
        
        cls_ProgressBar.ProgressBar_Progress (j / totalRecords)
        
        If i = Limit Then
            sqlStr = ReplaceDoubleQuote("UPDATE " & Esc(SanitizedAppName) & "." & Esc(BeTableName) & " SET " & setFields)
            ''Insert to tblCsvInsertStatementItems
            RunSQL "INSERT INTO tblCsvInsertStatementItems (Cursor,SqlStatement,CsvInsertStatementID) values (" & _
                Esc(Cursor) & "," & Esc(rows.JoinArr(";") & ";") & "," & CsvInsertStatementID & ")"
            ''Update the cursor
            Cursor = rs.fields(PrimaryKey)
            Set rows = New clsArray
            i = 0
        End If
        rs.MoveNext
    Loop

    If rows.count > 0 Then
        ''Insert to tblCsvInsertStatementItems
        RunSQL "INSERT INTO tblCsvInsertStatementItems (Cursor,SqlStatement,CsvInsertStatementID) values (" & _
            Esc(Cursor) & "," & Esc(rows.JoinArr(";") & ";") & "," & CsvInsertStatementID & ")"
    End If
    
    cls_ProgressBar.ProgressBar_Hide
    
End Function


Private Function UpdateUpdateVariables(SeqModelID, TableName, TargetFields, TimestampField, DatabasePrimaryKey)
    
    Dim fields As New clsArray
     
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT ImportFieldName IS NULL ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName")
        Dim PrimaryKey: PrimaryKey = rs.fields("PrimaryKey")
        If PrimaryKey Then
            DatabasePrimaryKey = DatabaseFieldName
        End If
        fields.Add Esc(DatabaseFieldName)
        rs.MoveNext
    Loop
    
    ''Timestamp
    If Not isFalse(TimestampField) Then
        fields.Add Esc("created_at")
    End If
    
    TargetFields = fields.JoinArr(",")
    ''fields.Add
    
End Function

Public Function RefetchCursor(frm As Form)

    Dim CsvInsertStatementItemID: CsvInsertStatementItemID = frm("subCsvInsertStatementItems").Form("CsvInsertStatementItemID")
    If ExitIfTrue(isFalse(CsvInsertStatementItemID), "Select a row to refetch..") Then Exit Function
    
    GenerateInsertSqlStatementsForPostgres frm, CsvInsertStatementItemID
    
End Function

Public Function GenerateInsertSqlStatementsForPostgres(frm As Object, Optional CsvInsertStatementItemID)
    
    RunCommandSaveRecord
    
    Dim CsvInsertStatementID: CsvInsertStatementID = frm("CsvInsertStatementID")
    Dim Cursor
    If Not isFalse(CsvInsertStatementItemID) Then
        
        Cursor = ELookup("tblCsvInsertStatementItems", "CsvInsertStatementItemID = " & CsvInsertStatementItemID, "Cursor")
        
        Dim resp: resp = MsgBox("This will rewrite all the SQL after the cursor: " & Esc(Cursor) & ". Do you want to proceed?", vbYesNo)
        
        If resp = vbNo Then Exit Function
            
        RunSQL "DELETE FROM tblCsvInsertStatementItems WHERE CsvInsertStatementID = " & CsvInsertStatementID & _
            " AND CsvInsertStatementItemID >= " & CsvInsertStatementItemID
    Else
        RunSQL "DELETE FROM tblCsvInsertStatementItems WHERE CsvInsertStatementID = " & CsvInsertStatementID
    End If
    
    Dim SourceDatabaseFile: SourceDatabaseFile = frm("SourceDatabaseFile"): If ExitIfTrue(isFalse(SourceDatabaseFile), "SourceDatabaseFile is empty..") Then Exit Function
    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    Dim Limit: Limit = frm("Limit")
    Dim PrimaryKey: PrimaryKey = frm("PrimaryKey")
    Dim TableName: TableName = frm("TableName")
    Dim TimestampField: TimestampField = frm("TimestampField")
    Dim FetchLimit: FetchLimit = frm("FetchLimit")
    
    Dim TargetFields, InsertStatements, DatabasePrimaryKey
    
    UpdateVariables SeqModelID, TableName, TargetFields, TimestampField, DatabasePrimaryKey
    
    GetInsertValues TableName, TargetFields, PrimaryKey, Limit, CsvInsertStatementID, SeqModelID, TimestampField, FetchLimit, SourceDatabaseFile, Cursor, DatabasePrimaryKey
    ''Dim sqlStr: sqlStr = "INSERT INTO " & TableName & "(" & TargetFields & ") VALUES " & InsertStatements & ";"
    
    frm("subCsvInsertStatementItems").Form.Requery
   
End Function

Private Function UpdateVariables(SeqModelID, TableName, TargetFields, TimestampField, DatabasePrimaryKey)
    
    Dim fields As New clsArray
     
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT ImportFieldName IS NULL ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName")
        Dim PrimaryKey: PrimaryKey = rs.fields("PrimaryKey")
        If PrimaryKey Then
            DatabasePrimaryKey = DatabaseFieldName
        End If
        fields.Add Esc(DatabaseFieldName)
        rs.MoveNext
    Loop
    
    ''Timestamp
    If Not isFalse(TimestampField) Then
        fields.Add Esc("created_at")
    End If
    
    TargetFields = fields.JoinArr(",")
    ''fields.Add
    
End Function

Private Function GetInsertValues(TableName, TargetFields, PrimaryKey, Limit, CsvInsertStatementID, SeqModelID, TimestampField, FetchLimit, SourceDatabaseFile, Cursor, DatabasePrimaryKey)
    
    Dim i As Integer: i = 0
    Dim j As Long: j = 0
    
    Dim sqlStr: sqlStr = "SELECT * FROM [" & TableName & "] IN " & Esc(SourceDatabaseFile)
    
    If Not isFalse(Cursor) Then
        Dim isPrimaryKeyAString: isPrimaryKeyAString = Not InStr(PrimaryKey, "ID") > 0
        sqlStr = sqlStr & " WHERE " & PrimaryKey & " > " & IIf(isPrimaryKeyAString, Esc(Cursor), Cursor)
    End If
    
    sqlStr = sqlStr & " ORDER BY " & PrimaryKey
    
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    Dim totalRecords: totalRecords = IIf(isFalse(FetchLimit), rs.recordCount, FetchLimit)
    
    
    Dim rows As New clsArray
    
    Dim SanitizedAppName: SanitizedAppName = ELookup("qrySeqModels", "SeqModelID = " & SeqModelID, "SanitizedAppName")
    Dim BeTableName: BeTableName = ELookup("qrySeqModels", "SeqModelID = " & SeqModelID, "TableName")
    
    Dim cls_ProgressBar As New cls_ProgressBar
    
    cls_ProgressBar.ProgressBar_Show
    cls_ProgressBar.ProgressBar_ProgressOverlay True
    cls_ProgressBar.ProgressBar_Caption "Filtering Data"
    ''cls_ProgressBar.ProgressBar_Resize 3000, 6000
    cls_ProgressBar.ProgressBar_Message1_FontParam 1, "Calibri", 100, 10, vbBlack
    cls_ProgressBar.ProgressBar_ProgressValue_Align 2
    
    Do Until rs.EOF
        
        rows.Add GenerateRowValues(rs, SeqModelID, TimestampField)
        
        i = i + 1
        j = j + 1
        
        If Not isFalse(FetchLimit) Then
            If j >= FetchLimit Then
                Exit Do
            End If
        End If
        
        cls_ProgressBar.ProgressBar_Progress (j / totalRecords)
        
        If i = Limit Then
            sqlStr = ReplaceDoubleQuote("INSERT INTO " & Esc(SanitizedAppName) & "." & Esc(BeTableName) & "(" & TargetFields & ") VALUES " & rows.JoinArr(",") & _
            " ON CONFLICT (" & DatabasePrimaryKey & ") DO NOTHING;")
            
            ''Insert to tblCsvInsertStatementItems
            RunSQL "INSERT INTO tblCsvInsertStatementItems (Cursor,SqlStatement,CsvInsertStatementID) values (" & _
                Esc(Cursor) & "," & Esc(sqlStr) & "," & CsvInsertStatementID & ")"
            ''Update the cursor
            Cursor = rs.fields(PrimaryKey)
            Set rows = New clsArray
            i = 0
        End If
        rs.MoveNext
    Loop

    If rows.count > 0 Then
        ''Insert to tblCsvInsertStatementItems
        sqlStr = ReplaceDoubleQuote("INSERT INTO " & Esc(SanitizedAppName) & "." & Esc(BeTableName) & "(" & TargetFields & ") VALUES " & rows.JoinArr(",") & _
            " ON CONFLICT (" & DatabasePrimaryKey & ") DO NOTHING;")
            
            ''Insert to tblCsvInsertStatementItems
            RunSQL "INSERT INTO tblCsvInsertStatementItems (Cursor,SqlStatement,CsvInsertStatementID) values (" & _
                Esc(Cursor) & "," & Esc(sqlStr) & "," & CsvInsertStatementID & ")"
    End If
    
    cls_ProgressBar.ProgressBar_Hide
    
End Function

Private Function GenerateRowValuesForUpdate(rs As Recordset, SeqModelID, TimestampField)

    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT ImportFieldName IS NULL ORDER BY SeqModelFieldID"
    Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
    Dim values As New clsArray, value, sanitizedVal
    Dim whereClause
    
    Do Until rs2.EOF
        Dim AllowNull: AllowNull = rs2.fields("AllowNull")
        Dim ImportFieldName: ImportFieldName = rs2.fields("ImportFieldName")
        Dim DataTypeInterface: DataTypeInterface = rs2.fields("DataTypeInterface")
        Dim DatabaseFieldName: DatabaseFieldName = rs2.fields("DatabaseFieldName")
        Dim PrimaryKey: PrimaryKey = rs2.fields("PrimaryKey")
        
        If Not DoesPropertyExists(rs.fields, ImportFieldName) Then
            GoTo NextField
        End If
        
        value = rs.fields(ImportFieldName)
        
        If AllowNull And IsNull(value) Then
            sanitizedVal = "NULL"
        Else
            If DataTypeInterface = "string" Then
                sanitizedVal = "$$" & value & "$$"
            ElseIf DataTypeInterface = "boolean" Then
                sanitizedVal = IIf(value, "true", "false")
            Else
                sanitizedVal = value
            End If
        End If
        
        If PrimaryKey Then
            whereClause = "WHERE " & DatabaseFieldName & " = " & sanitizedVal
        Else
            values.Add DatabaseFieldName & " = " & sanitizedVal
        End If
NextField:
        rs2.MoveNext
    Loop
        
    ''TimestampField - this is not necessary as this is updated
'    If Not isFalse(TimestampField) Then
'        Dim TsValue: TsValue = rs.fields(TimestampField)
'        If Not isFalse(TsValue) Then
'            values.Add "$$" & TsValue & "$$"
'        Else
'            values.Add "DEFAULT"
'        End If
'    End If
    
    ''Add the where clause here
    GenerateRowValuesForUpdate = values.JoinArr(",") & " " & whereClause
    
End Function

Private Function GenerateRowValues(rs As Recordset, SeqModelID, TimestampField)

    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT ImportFieldName IS NULL AND NOT IsGeneratedField AND Expression IS NULL ORDER BY SeqModelFieldID"
    Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
    Dim values As New clsArray, value, sanitizedVal
    Do Until rs2.EOF
        Dim AllowNull: AllowNull = rs2.fields("AllowNull")
        Dim ImportFieldName: ImportFieldName = rs2.fields("ImportFieldName")
        Dim DataTypeInterface: DataTypeInterface = rs2.fields("DataTypeInterface")
        Dim DataType: DataType = rs2.fields("DataType")
        
        value = rs.fields(ImportFieldName)
        
        If AllowNull And IsNull(value) Then
            sanitizedVal = "NULL"
        Else
        
            If DataType = "DATE" Or DataType = "DATEONLY" Then
                value = SQLDate(value, True)
            End If
            
            If DataTypeInterface = "string" Then
                sanitizedVal = "$$" & value & "$$"
            ElseIf DataTypeInterface = "boolean" Then
                sanitizedVal = IIf(value, "true", "false")
            Else
                sanitizedVal = value
            End If
        End If
        
        values.Add sanitizedVal
        rs2.MoveNext
    Loop
        
    ''TimestampField
    If Not isFalse(TimestampField) Then
        Dim TsValue: TsValue = rs.fields(TimestampField)
        If Not isFalse(TsValue) Then
            values.Add "$$" & TsValue & "$$"
        Else
            values.Add "DEFAULT"
        End If
    End If

    GenerateRowValues = "(" & values.JoinArr(",") & ")"
    
End Function
