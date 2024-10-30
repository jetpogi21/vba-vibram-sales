Attribute VB_Name = "FormDefaultOnCurrent Mod"
Option Compare Database
Option Explicit

Public Function FormDefaultOnCurrentCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

''This will run the =SetMiddleTableValues([Form],"tblSnippetCategories","tblSnippetCategoryID","CategoryID","Snippet")
Public Function frmDefaultAfterUpdate(frm As Object, ModelID)
    
    ''TABLE: tblControlCreationHelper Fields: ControlCreationHelperID|CustomControlTypeID|Model|PrimaryKey
    ''MainField|FieldToUse|Direction|Timestamp|CreatedBy|RecordImportID|Width|ParentModel|IsDynamicList|FieldCaption
    ''MiddleTable|PossibleValues|NoneValue|PairSize|AsInline
    Dim Model: Model = ELookup("tblModels", "ModelID = " & ModelID, "Model")
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblControlCreationHelper WHERE ParentModel = " & Esc(Model)
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim MiddleTable, ListFormRecordsource, FieldToUse
        ''MiddleTable this is the table to be updated once the form is updated/unloaded
        MiddleTable = rs.fields("MiddleTable")
        FieldToUse = rs.fields("FieldToUse")
        ListFormRecordsource = "tbl" & Model & FieldToUse
        
        SetMiddleTableValues frm, MiddleTable, ListFormRecordsource, FieldToUse, Model
        rs.MoveNext
    Loop
    
End Function

Public Function frmDefaultOnUnload(frm As Object, ModelID)
    
    ''TABLE: tblControlCreationHelper Fields: ControlCreationHelperID|CustomControlTypeID|Model|PrimaryKey
    ''MainField|FieldToUse|Direction|Timestamp|CreatedBy|RecordImportID|Width|ParentModel|IsDynamicList|FieldCaption
    ''MiddleTable|PossibleValues|NoneValue|PairSize|AsInline
    Dim Model: Model = ELookup("tblModels", "ModelID = " & ModelID, "Model")
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblControlCreationHelper WHERE ParentModel = " & Esc(Model)
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim MiddleTable, ListFormRecordsource, FieldToUse
        ''MiddleTable this is the table to be updated once the form is updated/unloaded
        MiddleTable = rs.fields("MiddleTable")
        FieldToUse = rs.fields("FieldToUse")
        ListFormRecordsource = "tbl" & Model & FieldToUse
        
        SetMiddleTableValues frm, MiddleTable, ListFormRecordsource, FieldToUse, Model
        rs.MoveNext
    Loop
    
End Function

Public Function frmDefaultOnCurrent(frm As Object, ModelID)

    ''Set which control will initially get focused on
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    ''TABLE: tblModels Fields: ModelID|Model|VerboseName|VerbosePlural|MainField|TableWideValidation|FormColumns
    ''SetFocus|IsKeyVisible|QueryName|OnFormCreate|SubformName|UserQueryFields|IsSystemTable|Timestamp|CreatedBy
    ''RecordImportID|PrimaryKey
    Dim SetFocus: SetFocus = rs.fields("SetFocus")
    Dim Model: Model = rs.fields("Model")
    
    If Not IsNull(SetFocus) Then SetFocusOnForm frm, SetFocus
    
    ''TABLE: tblControlCreationHelper Fields: ControlCreationHelperID|CustomControlTypeID|Model|PrimaryKey
    ''MainField|FieldToUse|Direction|Timestamp|CreatedBy|RecordImportID|Width|ParentModel|IsDynamicList|FieldCaption
    ''MiddleTable|PossibleValues
    Set rs = ReturnRecordset("SELECT * FROM tblControlCreationHelper WHERE ParentModel = " & Esc(Model))
    Do Until rs.EOF
    
        Dim IsDynamicList: IsDynamicList = rs.fields("IsDynamicList")
        Dim CustomControlTypeID: CustomControlTypeID = rs.fields("CustomControlTypeID")
        Dim CustomControlType: CustomControlType = ELookup("tblCustomControlTypes", "CustomControlTypeID = " & CustomControlTypeID, "CustomControlType")
        Dim FieldToUse: FieldToUse = rs.fields("FieldToUse")
        Dim MiddleTable: MiddleTable = rs.fields("MiddleTable")
        
        If IsDynamicList Or CustomControlType = "Checkbox" Then
            Dim SubformName: SubformName = "sub" & FieldToUse
            SetSubformCheck frm, Model, SubformName, MiddleTable, FieldToUse
        Else
            Dim ogName: ogName = "og" & FieldToUse
            SetOptionGroupValue frm, FieldToUse, ogName
        End If
        
        rs.MoveNext
    Loop
    
End Function
