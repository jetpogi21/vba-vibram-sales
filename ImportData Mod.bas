Attribute VB_Name = "ImportData Mod"
Option Compare Database
Option Explicit

Public Function ChangeToSelectedFile()
    
    ''FileDialog
    ''Update All tblImportData -> SourceDatabaseFile to the selected file
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
    
    UpdateToDatabaseFile fullPath
    
    
End Function

Private Sub UpdateToDatabaseFile(Optional fullPath = "C:\Users\user\Desktop\Databases\Sales Database\Sales_Local.accdb")
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = "tblImportData"
        .SetStatement = "SourceDatabaseFile = " & EscapeString(fullPath)
        rowsAffected = .Run
    End With
    
    Forms!mainImportData.subform.Requery
    
End Sub

Public Function RevertToDefaultFile()

    UpdateToDatabaseFile

End Function

Public Function UpdateDestinationConnector(frm As Form)

    Dim SourceConnector, DestinationConnector
    SourceConnector = frm("SourceConnector")
    DestinationConnector = frm("DestinationConnector")
    
    If IsNull(DestinationConnector) Then frm("DestinationConnector") = SourceConnector
    
End Function

Public Function ImportDataOnLoad(frm As Form)
    
    DefaultFormLoad frm, "ImportDataID"
    frm("ModelID").ListItemsEditForm = ""
    frm("ExternalDBTableID").ListItemsEditForm = ""
    
    
    'frm("subImportDataFields").Form.ModelFieldID.RowSource = "SELECT ModelFieldID, ModelField FROM tblModelFields WHERE ModelFieldID = 0"
    'frm("subImportDataFields").Form.ExternalDBFieldID.RowSource = "SELECT ExternalDBFieldID, FieldName FROM tblExternalDBFields WHERE ExternalDBFieldID = 0"
    
End Function

Public Function ImportDataCurrent(frm As Form)

    SetFocusOnForm frm, ""
    
    Dim ImportDataID, SourceDatabaseFile
    ImportDataID = frm("ImportDataID")
    SourceDatabaseFile = frm("SourceDatabaseFile")
    
    If frm.NewRecord Then
        frm("ExternalDBTableID").RowSource = "SELECT ExternalDBTableID, TableName FROM tblExternalDBTables WHERE ExternalDBTableID = 0"
        frm("subImportDataFields").Form.ModelFieldID.RowSource = "SELECT ModelFieldID, ModelField FROM tblModelFields WHERE ModelFieldID = 0"
        frm("subImportDataFields").Form.ExternalDBFieldID.RowSource = "SELECT ExternalDBFieldID, FieldName FROM tblExternalDBFields WHERE ExternalDBFieldID = 0"
    Else
    
        frm("ExternalDBTableID").RowSource = "SELECT ExternalDBTableID, TableName FROM tblExternalDBTables ORDER BY TableName ASC"
        frm("subImportDataFields").Form.ModelFieldID.RowSource = "SELECT ModelFieldID, ModelField FROM tblModelFields WHERE ImportDataID = " & ImportDataID
        frm("subImportDataFields").Form.ExternalDBFieldID.RowSource = "SELECT ExternalDBFieldID, FieldName FROM tblExternalDBFields WHERE ImportDataID = " & ImportDataID
        
    End If
    
    ''Set the ModelFieldID Rowsource
    'frm("ModelID").RowSource = "SELECT ModelID, Model FROM tblModels WHERE ModelID = 0"
    ''Set the ExternalDBTableID Rowsource
    
End Function

Public Function ExternalDBTableIDAfterUpdate(frm As Form)
    
    Dim ExternalDBTableID, TableName, ImportDataID
    ExternalDBTableID = frm("ExternalDBTableID")
    ImportDataID = frm("ImportDataID")
    
    If IsNull(ExternalDBTableID) Then Exit Function
    TableName = frm("ExternalDBTableID").Column(1)
    
    ''Fetch all the field from this Table
    Dim SourceDatabaseFile
    SourceDatabaseFile = frm("SourceDatabaseFile")
    Dim db As DAO.Database
    Set db = OpenDatabase(SourceDatabaseFile)
    
    Dim fld As DAO.field
    
    ''Remove all the data from tblExternalDBFields
    'RunSQL "DELETE FROM tblExternalDBFields WHERE ImportDataID = " & ImportDataID
    
    Dim rsDef As Object
    If DoesPropertyExists(db.TableDefs, TableName) Then
        Set rsDef = db.TableDefs(TableName)
    Else
        Set rsDef = db.QueryDefs(TableName)
    End If
    
    For Each fld In rsDef.fields
        ''Insert the tableName into the tblExternalDBTables
        If Not isPresent("tblExternalDBFields", "ImportDataID = " & ImportDataID & " And FieldName = " & EscapeString(fld.Name)) Then
            RunSQL "INSERT INTO tblExternalDBFields (ImportDataID,FieldName) VALUES (" & ImportDataID & "," & EscapeString(fld.Name) & ")"
        End If
    Next fld

    frm("subImportDataFields").Form.ExternalDBFieldID.RowSource = "SELECT ExternalDBFieldID, FieldName FROM tblExternalDBFields WHERE " & _
                                                                  "ImportDataID = " & ImportDataID & " ORDER BY FieldName"
    frm("subImportDataFields").Form.ExternalDBFieldID.Requery
    
End Function

Public Function BrowserDataSourceDatabaseFile(frm As Object, Optional dontFetchTables As Boolean = False)
    
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
    
    If (dontFetchTables) Then Exit Function
    frm.Dirty = False
    
    Dim ImportDataID
    ImportDataID = frm("ImportDataID")
    ''Remove all the data from tblExternalDBTables
    'RunSQL "DELETE FROM tblExternalDBTables WHERE SourceDatabaseFile = " & EscapeString(fullPath)
    
    ''Get all the tables from the SourceDatabaseFile = fullPath
    Dim db As DAO.Database
    Set db = OpenDatabase(fullPath)
    
    Dim tblDef As TableDef, qDef As QueryDef
    For Each tblDef In db.TableDefs
        If Not (tblDef.Name Like "*~*" Or tblDef.Name Like "*MSys*") Then
            ''Insert the tableName into the tblExternalDBTables
            If Not isPresent("tblExternalDBTables", "SourceDatabaseFile = " & EscapeString(fullPath) & " And TableName = " & EscapeString(tblDef.Name)) Then
                RunSQL "INSERT INTO tblExternalDBTables (TableName,SourceDatabaseFile) VALUES (" & EscapeString(tblDef.Name) & "," & EscapeString(fullPath) & ")"
            End If
        End If
    Next tblDef
    
    For Each qDef In db.QueryDefs
        If Not qDef.Name Like "*~*" Then
            ''Insert the queryName into the tblExternalDBTables
            If Not isPresent("tblExternalDBTables", "SourceDatabaseFile = " & EscapeString(fullPath) & " And TableName = " & EscapeString(qDef.Name)) Then
                RunSQL "INSERT INTO tblExternalDBTables (TableName,SourceDatabaseFile) VALUES (" & EscapeString(qDef.Name) & "," & EscapeString(fullPath) & ")"
            End If
        End If
    Next qDef
    
    frm("ExternalDBTableID").RowSource = "SELECT ExternalDBTableID, TableName FROM tblExternalDBTables WHERE SourceDatabaseFile = " & EscapeString(fullPath) & _
                                         " ORDER BY TableName ASC"
                                         
    frm("subImportDataFields").Form.Requery
    'frm("SourceConnector") = Null
    'frm("DestinationConnector") = "RecordImportID"
    'frm("ExternalDBTableID") = Null
    'frm("ModelID") = Null

End Function

Private Function GenerateImportDataFields(ImportDataID, sourceTableName, sourceFields, destinationFields, joinObjs As Collection, typecastedFields) As Boolean
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryImportDataFields WHERE ImportDataID = " & ImportDataID & " AND ExternalDBFieldID IS NOT NULL")
    
    Dim ImportDataFieldID, ModelFieldID, LookupTable, LookupField, ReturnField, ExternalDBFieldID, ModelField, fieldName, DestinationConnector, SourceConnector
    Dim sourceFieldArr As New clsArray, destFieldArr As New clsArray, typecastedFieldArr As New clsArray
    
    If rs.EOF Then Exit Function
    Do Until rs.EOF
        ImportDataFieldID = rs.fields("ImportDataFieldID")
        ModelFieldID = rs.fields("ModelFieldID")
        LookupTable = rs.fields("LookupTable")
        LookupField = rs.fields("LookupField")
        ReturnField = rs.fields("ReturnField")
        ExternalDBFieldID = rs.fields("ExternalDBFieldID")
        ModelField = rs.fields("ModelField")
        fieldName = rs.fields("FieldName")
        DestinationConnector = rs.fields("DestinationConnector")
        SourceConnector = rs.fields("SourceConnector")
        
        If IsNull(LookupTable) Then
            sourceFieldArr.Add "[" & sourceTableName & "]." & fieldName & " AS " & ModelField
        Else
            
            If fieldName Like "*_id" Then
            
                typecastedFieldArr.Add "CLng(" & fieldName & ") As TypeCasted" & fieldName
                fieldName = "TypeCasted" & fieldName
                
            End If
            
            joinObjs.Add GenerateJoinObj(LookupTable, fieldName, , LookupField, "LEFT")
            sourceFieldArr.Add "CLng(" & LookupTable & "." & ReturnField & ") AS " & ModelField
            
        End If
        
        destFieldArr.Add ModelField
        
        rs.MoveNext
    Loop
    
    ''If the id will be put on the RecordImportID
    If DestinationConnector = "RecordImportID" Then
        sourceFieldArr.Add "CLng([" & sourceTableName & "]." & SourceConnector & ") AS RecordImportID"
        destFieldArr.Add "RecordImportID"
    End If
    
    sourceFields = sourceFieldArr.JoinArr
    destinationFields = destFieldArr.JoinArr
    typecastedFields = typecastedFieldArr.JoinArr
    
    GenerateImportDataFields = True
    
End Function

Public Function ImportDataModelIDAfterUpdate(frm As Form)
    
    Dim ModelID, ImportDataID
    ModelID = frm("ModelID")
    ImportDataID = frm("ImportDataID")
    
    If IsNull(ModelID) Then Exit Function

    frm("subImportDataFields").Form.ModelFieldID.RowSource = "SELECT ModelFieldID, ModelField FROM tblModelFields WHERE ModelID = " & ModelID & " ORDER BY ModelField"
    
    ''Insert all the ModelFields of the selected ModelID to the tblImportDataFields
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModelFields"
        .AddFilter "ModelID = " & ModelID
        .fields = "ModelFieldID," & ImportDataID & " AS ImportDataID"
        sqlStr = .sql
    End With
    
    ''Delete all data from tblImportDataFields
    RunSQL "DELETE FROM tblImportDataFields WHERE ImportDataID = " & ImportDataID
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblImportDataFields"
        .fields = "ModelFieldID, ImportDataID"
        .insertSQL = sqlStr
        .InsertFilterField = "ModelFieldID, ImportDataID"
        rowsAffected = .Run
    End With
    
    frm("subImportDataFields").Requery
    
End Function

Public Function ImportDataClick(frm As Object, Optional notify As Boolean = True)
    
    If Not areDataValid2(frm, "ImportData2") Then Exit Function
    
    ''Declare the Variables
    Dim ImportDataID, DestinationConnector, SourceConnector, SourceDatabaseFile, ExternalDBTableID, ModelID
    ImportDataID = frm("ImportDataID")
    DestinationConnector = frm("DestinationConnector")
    SourceConnector = frm("SourceConnector")
    SourceDatabaseFile = frm("SourceDatabaseFile")
    ExternalDBTableID = frm("ExternalDBTableID")
    ModelID = frm("ModelID")
       
    ''Get the TableName of the Destination Table
    Dim destTableName, sourceTableName
    destTableName = GetTableNameFromModelID(ModelID)
    sourceTableName = frm("ExternalDBTableID").Column(1)
    ''Get the Fields to be used on the subquery
    Dim sourceFields, destinationFields, joinObjs As New Collection, joinObj, typecastedFields
    If ExitIfTrue(Not GenerateImportDataFields(ImportDataID, sourceTableName, sourceFields, destinationFields, joinObjs, typecastedFields), _
        "Please put at least one field to be imported..") Then Exit Function
        
    ''End Goal
    ''Insert or Update statement depending if the source connector is present to the destination connector or not
    Dim sqlObj As clsSQL, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "[" & sourceTableName & "] IN " & EscapeString(SourceDatabaseFile)
        If joinObjs.count = 0 Then .fields = sourceFields
        If typecastedFields <> "" Then
            .fields.Add "*," & typecastedFields
        End If
        sqlStr = .sql
    End With
    
    If joinObjs.count > 0 Then
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = sqlStr
            .SourceAlias = "temp"
            .fields = sourceFields
            For Each joinObj In joinObjs
                .joins.Add joinObj
            Next joinObj
            sqlStr = .sql
        End With
    End If
    
    ''Create a temp table of the sqlStr above
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = sqlStr
        .MakeTable = "temp"
        .SourceAlias = "temp"
        .Run
    End With
    
    
    ''Update if the data is found
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = destTableName
        .SetStatement = BuildSetStatement(destTableName, "temp", destinationFields)
        .joins.Add GenerateJoinObj("temp", DestinationConnector)
        rowsAffected = .Run
    End With
    
    ''Insert if the connector isn't found
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "temp"
        .AddFilter destTableName & "." & DestinationConnector & " IS NULL"
        .fields = "temp.*"
        .joins.Add GenerateJoinObj(destTableName, DestinationConnector, , , "LEFT")
        sqlStr = .sql
    End With
           
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = destTableName
        .fields = destinationFields
        .insertSQL = sqlStr
        .InsertFilterField = destinationFields
        rowsAffected = .Run
    End With
    
    If notify Then MsgBox "Records imported successfully.."
      
End Function

Public Function ImportExternalModels()

    ''Get the shape of
End Function

