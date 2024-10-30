Attribute VB_Name = "ImportModel Mod"
Option Compare Database
Option Explicit

Public Function ImportModelCurrent(frm As Form)

    SetFocusOnForm frm, ""
    
    ''Set the ModelID Rowsource
    frm("ModelID").RowSource = "SELECT ModelID, Model FROM tblModels WHERE ModelID = 0"
    frm("ModelID").ListItemsEditForm = ""
    
End Function

Public Function BrowserSourceDatabaseFile(frm As Form)
    
    Dim fs As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPath, fileName
    With FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .filters.Add "MS Access Database Files", "*.accdb; *.mdb", 1
        'Show the dialog box
        If ExitIfTrue(Not .Show, "Database selection cancelled..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        fileName = fs.GetFileName(fullPath)
        
    End With
    
    frm("SourceDatabaseFile") = fullPath
    ''Select the Models From the database path
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModels IN " & EscapeString(fullPath)
        .AddFilter "IsSystemTable = 0"
        .OrderBy = "ModelID ASC"
        sqlStr = .sql
        Set rs = .Recordset
    End With
    
    frm("ModelID").RowSource = sqlStr

    
End Function

Public Function BuildSetStatement(rsToBeUpdated, rsToBeUsedToUpdate, fieldNames, Optional IncludeRecordImportID As Boolean = False) As String
    
    Dim fieldArr As New clsArray, setStatementArr As New clsArray, fieldArrItem
    fieldArr.arr = fieldNames
    
    If IncludeRecordImportID Then fieldArr.Add "RecordImportID"
    
    For Each fieldArrItem In fieldArr.arr
        setStatementArr.Add concat(rsToBeUpdated, "!", fieldArrItem, " = ", rsToBeUsedToUpdate, "!", fieldArrItem)
    Next fieldArrItem
    
    BuildSetStatement = setStatementArr.JoinArr(",")
    
End Function

Public Function ImportRelatedModelRecords(ModelID, SourceDatabaseFile)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    ''Get all the models (their ModelID specifically) with ParentModelID of Model -> This will enumerate all tables having model as a field
    ''Another interpretation -> Get the ModelID that uses this current ModelID as ParentModelID
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModelFields"
        .AddFilter "ParentModelID = " & ELookup("tblModels", "Model = " & EscapeString("Model"), "ModelID")
        .fields = "ModelID"
        .GroupBy = "ModelID"
        sqlStr = .sql
        Set rs = .Recordset
    End With
    
    Dim tblName, PrimaryKey, db As DAO.Database
    Set db = OpenDatabase(SourceDatabaseFile)
    Do Until rs.EOF
        ''Get the table name of the model and the primary key
        tblName = GetTableNameFromModelID(rs.fields("ModelID"))
        PrimaryKey = concat(ELookup("tblModels", "ModelID = " & rs.fields("ModelID"), "Model"), "ID")
        ''Do this only if the ModelID is existing from the old database
        If DoesPropertyExists(db.TableDefs, tblName) Then
            ImportTableRecords tblName, ModelID, PrimaryKey, SourceDatabaseFile
        End If
        rs.MoveNext
    Loop
    
End Function

Private Function ImportTableRecords(tblName, ModelID, PrimaryKey, SourceDatabaseFile)
    
    ''Get the table from the old database
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = tblName & " IN " & EscapeString(SourceDatabaseFile)
        .AddFilter "ModelID = " & ModelID
        .OrderBy = PrimaryKey & " ASC"
        sqlStr = .sql
        Set rs = .Recordset
    End With
    
    ''Get the fieldnames to be used for the insert statement
    Dim fieldArr As New clsArray, fieldArrItem, fld As DAO.field, fieldNames, isParentModel, isModelField
    
    ''Why is the "isParentModel" important here? --? Also fetch the ModelFieldID, ParentModelID from the main Database
    ''instead of using the raw value
    
    isParentModel = False
    isModelField = False
    ''This are the fields from the old table -> Add all fields except the Timestamp, CreatedBy, RecordImportID, PK and ModelID
    For Each fld In rs.fields
        Select Case fld.Name
            ''Remove the primary key of this table
            Case "Timestamp", "CreatedBy", "RecordImportID", PrimaryKey, "ModelID":
            
            Case Else:
                If fld.Name = "ParentModelID" Then isParentModel = True
                If fld.Name = "ModelFieldID" Then isModelField = True
                fieldArr.Add fld.Name
        End Select
    Next fld
    
    ''Modify the sql -> do proper lookup to the specific tables -> Disregard the previous sqlStr ?
    If isParentModel Or isModelField Then
    
        Dim allFieldExceptParentModelIDArr As New clsArray
        For Each fld In rs.fields
            Select Case fld.Name
                Case "ParentModelID", "ModelFieldID":
                    
                Case Else:
                    allFieldExceptParentModelIDArr.Add "temp." & fld.Name
            End Select
        Next fld
        
        Dim fieldsArr As New clsArray: fieldsArr.arr = allFieldExceptParentModelIDArr.arr
        If isParentModel Then fieldsArr.Add "Clng(tblModels.ModelID) As ParentModelID"
        If isModelField Then fieldsArr.Add "Clng(tblModelFields.ModelFieldID) As ModelFieldID"
        
        If PrimaryKey = "ModelFieldID" Then fieldsArr.Add PrimaryKey
          
        ''SELECT STATEMENT
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = sqlStr
            .fields = fieldsArr.JoinArr
            If isParentModel Then .joins.Add GenerateJoinObj("tblModels", "ParentModelID", , "RecordImportID", "LEFT")
            If isModelField Then .joins.Add GenerateJoinObj("tblModelFields", "ModelFieldID", , "RecordImportID", "LEFT")
            .SourceAlias = "temp"
            sqlStr = .sql
        End With
        
    End If
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = sqlStr
        .MakeTable = "temp"
        .SourceAlias = "temp"
        rowsAffected = .Run
    End With
    
    fieldNames = fieldArr.JoinArr
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "temp"
        .fields = fieldNames & "," & _
                  PrimaryKey & " As RecordImportID," & _
                  ELookup("tblModels", "RecordImportID = " & ModelID, "ModelID") & " AS ModelID"
        sqlStr = .sql
    End With
       
    ''Make sure that the chosen model is not yet present from the current file's own models
    ''or else do an update
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = tblName
        .SetStatement = BuildSetStatement(tblName, "temp", fieldNames & ",ModelID", True)
        .joins.Add GenerateJoinObj(sqlStr, PrimaryKey, "temp", "RecordImportID")
        rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .AddFilter tblName & "." & PrimaryKey & " IS NULL"
        .fields = "temp.*"
        .SourceAlias = "temp"
        .joins.Add GenerateJoinObj(tblName, "RecordImportID", , , "LEFT")
        sqlStr = .sql
    End With
    
     ''Insert the actual table model
     Set sqlObj = New clsSQL
     With sqlObj
         .SQLType = "INSERT"
         .Source = tblName
         .fields = concat(fieldNames, ",RecordImportID,ModelID")
         .insertSQL = sqlStr
         .InsertFilterField = concat(fieldNames, ",RecordImportID,ModelID")
         rowsAffected = .Run
     End With
    
     
End Function

Public Function ImportModelClick(frm As Form)
    
    ''Get the recordset of the ModelID
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set rs = ReturnRecordset(frm("ModelID").RowSource)
    
    ''Get the fieldnames to be used for the insert statement
    Dim fieldArr As New clsArray, fieldArrItem, fld As DAO.field, fieldNames
    For Each fld In rs.fields
        Select Case fld.Name
            Case "Timestamp", "CreatedBy", "RecordImportID", "ModelID":

            Case Else:
                fieldArr.Add fld.Name
        End Select
    Next fld

    fieldNames = fieldArr.JoinArr
    
    Dim SourceDatabaseFile, ModelID
    SourceDatabaseFile = frm("SourceDatabaseFile")
    ModelID = frm("ModelID")
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModels IN " & EscapeString(SourceDatabaseFile)
        .AddFilter "IsSystemTable = 0 AND ModelID = " & ModelID
        .fields = concat(fieldNames, ",ModelID As RecordImportID")
        sqlStr = .sql
        Set rs = .Recordset
        
    End With
    
    ''Make sure that the chosen model is not yet present from the current file's own models
    ''or else do an update
    
    If Not isPresent("tblModels", concat("Model = ", EscapeString(rs.fields("Model")))) Then
        ''Insert the actual table model
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "INSERT"
            .Source = "tblModels"
            .fields = concat(fieldNames, ",RecordImportID")
            .insertSQL = sqlStr
            .InsertFilterField = concat(fieldNames, ",RecordImportID")
            rowsAffected = .Run
        End With
        
    Else
    
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "UPDATE"
            .Source = "tblModels"
            .SetStatement = BuildSetStatement("tblModels", "temp", fieldNames, True)
            .joins.Add GenerateJoinObj(sqlStr, "Model", "temp")
            rowsAffected = .Run
        End With

    End If
    
    ''Also import the related records
    ImportRelatedModelRecords ModelID, SourceDatabaseFile
    
    MsgBox "Model succesfully imported..."

    
End Function
