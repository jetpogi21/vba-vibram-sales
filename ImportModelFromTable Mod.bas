Attribute VB_Name = "ImportModelFromTable Mod"
Option Compare Database
Option Explicit

Public Function ImportModelFromTableCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
          frm("cmdDatabasePath").OnClick = "=SetFilePath([Form],""Access Database"",""DatabasePath"")"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function ImportSeqModelFromTable(frm As Form)
    
    ''This imports to the SeqModel
    Dim DatabasePath: DatabasePath = frm("DatabasePath"): If ExitIfTrue(isFalse(DatabasePath), "DatabasePath is empty..") Then Exit Function
    Dim TableName: TableName = frm("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Function
    Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim PrimaryKeyField: PrimaryKeyField = frm("PrimaryKeyField"): If ExitIfTrue(isFalse(PrimaryKeyField), "PrimaryKeyField is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM [" & TableName & "] IN " & Esc(DatabasePath)
    
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim fld As field
    
    Dim fldNames As New clsArray: fldNames.arr = "FieldName,SeqDataTypeID,PrimaryKey,SeqModelID," & _
        "DatabaseFieldName,FieldOrder,VerboseFieldName,ImportFieldName,WebControlTypeID,AllowNull,HideInTable"
    Dim fldValues As New clsArray
    Dim i As Integer: i = 1
    For Each fld In rs.fields
        
        Set fldValues = New clsArray
        If fld.Name = "CreatedBy" Or fld.Name = "RecordImportID" Or fld.Name = "Timestamp" Then
            GoTo NextField
        End If
    
        Dim fieldName
        If fld.Name = PrimaryKeyField Then
            fieldName = "id":
        Else
            fieldName = LCase(replace(ConvertToVerboseCaption(fld.Name), " ", "_"))
        End If
        
        fldValues.Add Esc(fieldName)
        
        Dim SeqDataTypeID: SeqDataTypeID = ELookup("tblFieldTypes", "FieldTypeID = " & fld.Type, "SeqDataTypeID")
        fldValues.Add SeqDataTypeID
        
        Dim PrimaryKey: PrimaryKey = fld.Name = PrimaryKeyField: fldValues.Add PrimaryKey
        fldValues.Add SeqModelID
        
        fldValues.Add Esc(fieldName)
        fldValues.Add i
        
        Dim VerboseFieldName: VerboseFieldName = ConvertToVerboseCaption(fld.Name)
        fldValues.Add Esc(VerboseFieldName)
        
        Dim ImportFieldName: ImportFieldName = fld.Name
        fldValues.Add Esc(ImportFieldName)
        
        Dim WebControlTypeID: WebControlTypeID = ELookup("tblFieldTypes", "FieldTypeID = " & fld.Type, "WebControlTypeID")
        If fld.Name = PrimaryKeyField Then
            WebControlTypeID = 13
        End If
        fldValues.Add Esc(WebControlTypeID)
        
        Dim AllowNull: AllowNull = -1
        If fld.Name = PrimaryKeyField Then
            AllowNull = 0
        ElseIf fld.Type = 1 Then
            AllowNull = 0
        End If
        fldValues.Add AllowNull
        
        Dim HideInTable: HideInTable = -1
        If fld.Name <> PrimaryKeyField Then
            HideInTable = 0
        End If
        fldValues.Add HideInTable
        
        RunSQL "INSERT INTO tblSeqModelFields (" & fldNames.JoinArr(",") & ") VALUES (" & fldValues.JoinArr(",") & ")"
        
        i = i + 1
        
NextField:
    Next fld
    
    MsgBox "Finished Importing"
    
    
End Function

Public Function ImportTableTo_tblModels(ByVal frm As Object, Optional ImportModelFromTableID = "")

    RunCommandSaveRecord

    If isFalse(ImportModelFromTableID) Then
        ImportModelFromTableID = frm("ImportModelFromTableID")
        If ExitIfTrue(isFalse(ImportModelFromTableID), "ImportModelFromTableID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblImportModelFromTables WHERE ImportModelFromTableID = " & ImportModelFromTableID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), """TableName"" is empty..") Then Exit Function
    ''Dim DatabasePath: DatabasePath = rs.fields("DatabasePath"): If ExitIfTrue(isFalse(DatabasePath), """DatabasePath"" is empty..") Then Exit Function
    
    DoCmd.OpenForm "frmModels", , , "ModelID = 0"
    Set frm = GetForm("frmModels")
    
    Dim Model: Model = TableName
    Model = RemoveMatchedPattern(Model, "^tbl")
    Model = ConvertToPascalCase(Model)
    Model = RemoveMatchedPattern(Model, "s$")
    
    frm("Model") = Model
    frmModelsModelAfterUpdate frm, TableName
    
    Dim ModelID: ModelID = frm("ModelID")
    ''Loop through the fields of the TableName
    Set rs = ReturnRecordset(TableName)
    Dim fld As field
    
    Dim exemptedFields As New clsArray: exemptedFields.arr = "created_at,updated_at"
    Dim i: i = 0
    
    Set frm = frm("subModelFields").Form
    
    Dim FieldType
    For Each fld In rs.fields
        Dim fldName: fldName = fld.Name
        If Not exemptedFields.InArray(fldName) Then
            
            frm("ModelField") = ConvertToPascalCase(fld.Name)
            frm("FieldTypeID") = fld.Type
            frm("FieldOrder") = i + 1
            
            Dim isFieldABoolean: isFieldABoolean = IsBoolean(TableName, fld)
            Dim isFieldANumber: isFieldANumber = IsNumber(TableName, fld)
            
            If isFieldABoolean Then
                frm("FieldTypeID") = 1
            ElseIf isFieldANumber Then
                
                If HasDecimal(TableName, fld) Then
                    frm("FieldTypeID") = 7
                Else
                    frm("FieldTypeID") = 4
                End If
                
            ElseIf fld.Name Like "*date*" Then
                frm("FieldTypeID") = 10
            End If
            
            Dim ValidationString
            If Not isFieldABoolean Then
                ValidationString = GetValidationString(TableName, fld)
            End If
            frm("ValidationString") = ValidationString
            
            
            
            Dim possibleValues
            If Not isFieldABoolean Then
                possibleValues = GetPossibleValues(TableName, fld)
            End If
            frm("PossibleValues") = possibleValues
            
            If Not isFalse(ValidationString) Then
                If ValidationString Like "*required*" And isFieldANumber Then
                    frm("DefaultValue") = "0"
                ElseIf ValidationString Like "*required*" And Not isFalse(possibleValues) Then
                    Dim values As New clsArray: values.arr = possibleValues
                    frm("DefaultValue") = values.arr(0)
                End If
            End If
            
            frm.Recordset.addNew
            
        End If
        i = i + 1
    Next fld
    
    rs.Close
    deleteTableIfExists TableName
 
End Function

Private Function GetPossibleValues(TableName, fld As field)
    
    GetPossibleValues = Null
    
    Dim fieldName: fieldName = fld.Name
    
    Dim values As New clsArray
    values.Add fieldName
    values.Add TableName
    
    Dim sqlStr: sqlStr = "SELECT [[FieldName]] FROM [[TableName]] WHERE NOT [[FieldName]] IS NULL GROUP BY [[FieldName]]"
    sqlStr = GetReplacedString(sqlStr, "FieldName,TableName", values)
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim recordCount: recordCount = CountRecordset(rs)
    
    If recordCount <= 10 Then
        Set values = New clsArray
        
        Do Until rs.EOF
            values.Add rs.fields(fieldName)
            rs.MoveNext
        Loop
        
        GetPossibleValues = values.JoinArr(",")
        Exit Function
    End If
    
    
End Function

Private Function IsNumber(TableName, fld As field)
    
    Dim values As New clsArray
    values.Add fld.Name

    Dim sqlStr: sqlStr = "Not [[FieldName]] Is NULL AND Not IsNumeric([[FieldName]])"
    sqlStr = GetReplacedString(sqlStr, "FieldName", values)
    
    IsNumber = Not isPresent(TableName, sqlStr)
    
End Function

Private Function HasDecimal(TableName, fld As field)
    
    Dim values As New clsArray
    values.Add fld.Name

    Dim sqlStr: sqlStr = "Not [[FieldName]] Is NULL AND [[FieldName]] - CInt([[FieldName]]) <> 0"
    sqlStr = GetReplacedString(sqlStr, "FieldName", values)
    
    HasDecimal = isPresent(TableName, sqlStr, True)
    
End Function

Private Function IsBoolean(TableName, fld As field)

    Dim values As New clsArray:  Set values = GetSampleFieldValue(TableName, fld.Name, , False)
    
    Dim validValues As New clsArray: validValues.arr = "true,false"
    
    Dim item
    Dim i As Integer: i = 0
    For Each item In values.arr
        If Not validValues.InArray(item) Then
            IsBoolean = False
            Exit Function
        End If
        i = i + 1
    Next item
    
    IsBoolean = True
    
End Function

Private Function GetValidationString(TableName, fld As field)
    
    GetValidationString = Null
    Dim items As New clsArray
    
    ''See if there's a null in the table field
    Dim fieldName: fieldName = fld.Name
    
    Dim HasNull: HasNull = isPresent(TableName, fld.Name & " Is Null")
    
    If Not HasNull Then
        items.Add "required"
    End If
    
    Dim values As New clsArray
    values.Add fieldName
    values.Add TableName
    
    Dim sqlStr: sqlStr = "SELECT [[FieldName]],Count(*) AS RecordCount FROM [[TableName]] GROUP BY [[FieldName]] HAVING COUNT(*) > 1"
    sqlStr = GetReplacedString(sqlStr, "FieldName,TableName", values)
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then
        items.Add "unique"
    End If
    
    If items.count > 0 Then
        GetValidationString = items.JoinArr(" ")
    End If
    
    
End Function
