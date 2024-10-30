Attribute VB_Name = "ModelCloner Mod"
Option Compare Database
Option Explicit

Public Function ModelClonerCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm("SourceBackendProjectID").AfterUpdate = "=SourceBackendProjectID_AfterUpdate([Form])"
            frm.OnCurrent = "=frmModelCloner_OnCurrent([Form])"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SourceBackendProjectID_AfterUpdate(frm As Form)
    
    Dim SourceBackendProjectID: SourceBackendProjectID = frm("SourceBackendProjectID")
    
    Dim sqlStr: sqlStr = "SELECT SeqModelID, ModelName FROM tblSeqModels"
    
    If Not isFalse(SourceBackendProjectID) Then
        sqlStr = sqlStr & " WHERE BackendProjectID = " & SourceBackendProjectID
    End If
    
    sqlStr = sqlStr & " ORDER BY ModelName ASC"
    
    frm("SourceSeqModelID").RowSource = sqlStr
    
End Function

Public Function frmModelCloner_OnCurrent(frm As Form)
    
    SourceBackendProjectID_AfterUpdate frm
    
End Function

Public Function CloneModel(frm As Object, Optional ModelClonerID = "")

    RunCommandSaveRecord

    If isFalse(ModelClonerID) Then
        ModelClonerID = frm("ModelClonerID")
        If ExitIfTrue(isFalse(ModelClonerID), "ModelClonerID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblModelCloners WHERE ModelClonerID = " & ModelClonerID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SourceSeqModelID: SourceSeqModelID = rs.fields("SourceSeqModelID"): If ExitIfTrue(isFalse(SourceSeqModelID), "SourceSeqModelID is empty..") Then Exit Function
    Dim TargetBackendProjectID: TargetBackendProjectID = rs.fields("TargetBackendProjectID"): If ExitIfTrue(isFalse(TargetBackendProjectID), "TargetBackendProjectID is empty..") Then Exit Function
    
    Dim fields As New clsArray: fields.arr = GetFields("tblSeqModels", "SeqModelID,Timestamp,CreatedBy,RecordImportID,BackendProjectID")
    sqlStr = "SELECT " & fields.JoinArr & "," & TargetBackendProjectID & " AS BackendProjectID  FROM tblSeqModels WHERE SeqModelID = " & SourceSeqModelID
    
    Dim SeqModelID
    Dim sqlObj As New clsSQL
    ''Actually insert the data into the tblSeqModels
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblSeqModels"
          .fields = fields.JoinArr & ",BackendProjectID"
          .insertSQL = sqlStr
          .InsertFilterField = fields.JoinArr & ",BackendProjectID"
          .Run
          SeqModelID = .LastInsertID
    End With
    
    fields.arr = replace(GetFields("tblSeqModelFields", "SeqModelFieldID,Timestamp,CreatedBy,RecordImportID,SeqModelID"), "Unique,", "[Unique],")
    
    sqlStr = "SELECT " & fields.JoinArr & "," & SeqModelID & " AS SeqModelID, SeqModelFieldID As RecordImportID FROM tblSeqModelFields WHERE SeqModelID = " & SourceSeqModelID
    
    ''Actually insert the data into the tblSeqModels
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblSeqModelFields"
          .fields = fields.JoinArr & ",SeqModelID,RecordImportID"
          .insertSQL = sqlStr
          .InsertFilterField = fields.JoinArr & ",SeqModelID,RecordImportID"
          .Run
    End With
    
    If IsFormOpen("frmBackendProjects") Then
        Forms("frmBackendProjects")("subSeqModels").Form.Requery
    End If
    
    fields.arr = replace(GetFields("tblSeqModelFilters", "SeqModelFilterID,Timestamp,CreatedBy,RecordImportID,SeqModelID,SeqModelFieldID,SeqModelRelationshipID,ModelListID", True), "Unique,", "[Unique],")
    
    Dim joinObj As clsJoin, rowsAffected
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblSeqModelFilters"
          .fields = fields.JoinArr & ", tblSeqModelFields.SeqModelFieldID AS SeqModelFieldID," & SeqModelID & " AS SeqModelID"
          .joins.Add GenerateJoinObj("tblSeqModelFields", "SeqModelFieldID", , "RecordImportID", "LEFT")
          .AddFilter "tblSeqModelFilters.SeqModelID = " & SourceSeqModelID
          sqlStr = .sql
    End With
    
    fields.arr = replace(GetFields("tblSeqModelFilters", "SeqModelFilterID,Timestamp,CreatedBy,RecordImportID,SeqModelID,SeqModelFieldID,SeqModelRelationshipID,ModelListID"), "Unique,", "[Unique],")
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblSeqModelFilters"
          .fields = fields.JoinArr & ",SeqModelFieldID,SeqModelID"
          .insertSQL = sqlStr
          .InsertFilterField = fields.JoinArr & ",SeqModelFieldID,SeqModelID"
          .Run
    End With
    
    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND NOT DatabaseFieldName IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    Dim DatabaseFieldName, SeqModelFieldID:
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        DatabaseFieldName = rs.fields("DatabaseFieldName")
        
        ''Get the SeqModelFieldID using a look-up to the SeqModelFields
        SeqModelFieldID = ELookup("tblSeqModelFields", "DatabaseFieldName = " & Esc("DatabaseFieldName") & _
            " AND SeqModelID = " & SeqModelID, "SeqModelFieldID")
        
        RunSQL "UPDATE tblSeqModelFilters SET SeqModelFieldID = " & SeqModelFieldID & " WHERE SeqModelFilterID = " & SeqModelFilterID
        rs.MoveNext
    Loop
    
    ''For the Sorting
    fields.arr = replace(GetFields("tblSeqModelSorts", "SeqModelSortID,Timestamp,CreatedBy,RecordImportID,SeqModelID,SeqModelFieldID", True), "Unique,", "[Unique],")
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblSeqModelSorts"
          .fields = fields.JoinArr & ", tblSeqModelFields.SeqModelFieldID AS SeqModelFieldID," & SeqModelID & " AS SeqModelID"
          .joins.Add GenerateJoinObj("tblSeqModelFields", "SeqModelFieldID", , "RecordImportID", "LEFT")
          .AddFilter "tblSeqModelSorts.SeqModelID = " & SourceSeqModelID
          sqlStr = .sql
    End With
    
    fields.arr = replace(GetFields("tblSeqModelSorts", "SeqModelSortID,Timestamp,CreatedBy,RecordImportID,SeqModelID,SeqModelFieldID"), "Unique,", "[Unique],")
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblSeqModelSorts"
          .fields = fields.JoinArr & ",SeqModelFieldID,SeqModelID"
          .insertSQL = sqlStr
          .InsertFilterField = fields.JoinArr & ",SeqModelFieldID,SeqModelID"
          .Run
    End With
    
    sqlStr = "SELECT * FROM qrySeqModelSorts WHERE SeqModelID = " & SeqModelID & " AND NOT DatabaseFieldName IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelSortID: SeqModelSortID = rs.fields("SeqModelSortID")
        DatabaseFieldName = rs.fields("DatabaseFieldName")
        
        ''Get the SeqModelFieldID using a look-up to the SeqModelFields
        SeqModelFieldID = ELookup("tblSeqModelFields", "DatabaseFieldName = " & Esc("DatabaseFieldName") & _
            " AND SeqModelID = " & SeqModelID, "SeqModelFieldID")
        
        RunSQL "UPDATE tblSeqModelSorts SET SeqModelFieldID = " & SeqModelFieldID & " WHERE SeqModelSortID = " & SeqModelSortID
        rs.MoveNext
    Loop
    
    ''For the Settings
    fields.arr = replace(GetFields("tblSeqModelSettings", "Timestamp,CreatedBy,RecordImportID,SeqModelID", True), "Unique,", "[Unique],")
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblSeqModelSettings"
          .fields = fields.JoinArr & "," & SeqModelID & " AS SeqModelID"
          .AddFilter "tblSeqModelSettings.SeqModelID = " & SourceSeqModelID
          sqlStr = .sql
    End With
    
    fields.arr = replace(GetFields("tblSeqModelSettings", "Timestamp,CreatedBy,RecordImportID,SeqModelID"), "Unique,", "[Unique],")
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblSeqModelSettings"
          .fields = fields.JoinArr & ",SeqModelID"
          .insertSQL = sqlStr
          .InsertFilterField = fields.JoinArr & ",SeqModelID"
          .Run
    End With
    
    MsgBox "Model successfully cloned.."
    
End Function
