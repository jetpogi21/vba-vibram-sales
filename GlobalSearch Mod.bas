Attribute VB_Name = "GlobalSearch Mod"
Option Compare Database
Option Explicit

Public Function GlobalSearchCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function DoGlobalSearch(frm As Object, Optional GlobalSearchID = "")

    RunCommandSaveRecord

    ''Get the value of the search box and the selected category
    Dim search: search = frm("fltrWildSearch"): If ExitIfTrue(isFalse(search), "search is empty..") Then Exit Function
    
    Dim categories As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblGlobalSearchfltrSearchCoverageID WHERE Selected"
    Dim rs: Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim value: value = rs.fields("Value"): If ExitIfTrue(isFalse(value), "Value is empty..") Then Exit Function
        categories.Add value
        rs.MoveNext
    Loop
    
    ''Start the insertion here
    sqlStr = "SELECT * FROM tblSearchCoverages"
    If categories.count > 0 Then
        sqlStr = sqlStr & " WHERE SearchCoverageID In(" & categories.JoinArr & ")"
    End If
    Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    Dim filterStr, fieldStr
    RunSQL "DELETE FROM tblGlobalSearchs"
    
    Do Until rs.EOF
        Dim SearchCoverageID: SearchCoverageID = rs.fields("SearchCoverageID")
        Dim SearchCoverageTable: SearchCoverageTable = rs.fields("SearchCoverageTable"): If ExitIfTrue(isFalse(SearchCoverageTable), "SearchCoverageTable is empty..") Then Exit Function
        Dim FieldsToUseAsDescription: FieldsToUseAsDescription = rs.fields("FieldsToUseAsDescription"): If ExitIfTrue(isFalse(FieldsToUseAsDescription), "FieldsToUseAsDescription is empty..") Then Exit Function
        Dim PrimaryKey: PrimaryKey = rs.fields("PrimaryKey")
        If isFalse(PrimaryKey) Then
            PrimaryKey = GetPrimaryKeyFieldFromTable(SearchCoverageTable)
        End If
        
        filterStr = GetGlobalSearchFieldFilters(SearchCoverageID, search)
        fieldStr = FieldsToUseAsDescription & " AS SearchDescription, " & SearchCoverageID & " AS SearchCoverageID, " & PrimaryKey & " AS RecordID"
        
        InsertToGlobalSearch fieldStr, filterStr, SearchCoverageTable
        rs.MoveNext
    Loop
    
    
    frm("subform").Form.Requery
    ''fltrWildSearch
    ''tblGlobalSearchfltrSearchCoverageID
    ''Selected, Value, FilterLabel -> Snippet, Buttons
    
    ''tblGlobalSearchs
    ''tblSnippets
    ''tblModelButtons
    ''tblSearchCoverages
    ''tblSearchCoverageFields
    
End Function

Private Function InsertToGlobalSearch(fieldStr, filtrStr, SearchCoverageTable)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = SearchCoverageTable
          If filtrStr <> "" Then .AddFilter filtrStr
          .fields = fieldStr
            sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblGlobalSearchs"
          .fields = "SearchDescription, SearchCoverageID, RecordID"
          .insertSQL = sqlStr
          .InsertFilterField = "SearchDescription, SearchCoverageID, RecordID"
          rowsAffected = .Run
    End With
        
End Function

''GetFieldsToUseAsDescription
Private Function GetGlobalSearchFieldFilters(SearchCoverageID, search) As String

    Dim sqlStr: sqlStr = "SELECT * FROM tblSearchCoverageFields WHERE SearchCoverageID = " & SearchCoverageID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim orFilters As New clsArray
    
    Do Until rs.EOF
        Dim SearchCoverageFieldName: SearchCoverageFieldName = rs.fields("SearchCoverageFieldName"): If ExitIfTrue(isFalse(SearchCoverageFieldName), "SearchCoverageFieldName is empty..") Then Exit Function
        orFilters.Add SearchCoverageFieldName & " LIKE '*" & search & "*'"
        rs.MoveNext
    Loop
    
    GetGlobalSearchFieldFilters = orFilters.JoinArr(" OR ")
    
End Function

Public Function SearchDescription_OnDblClick(frm As Form)
    
    Dim SearchCoverageID: SearchCoverageID = frm("SearchCoverageID"): If ExitIfTrue(isFalse(SearchCoverageID), "SearchCoverageID is empty..") Then Exit Function
    Dim RecordID: RecordID = frm("RecordID"): If ExitIfTrue(isFalse(RecordID), "RecordID is empty..") Then Exit Function
    Dim SearchCoverageName: SearchCoverageName = frm("SearchCoverageName"): If ExitIfTrue(isFalse(SearchCoverageName), "SearchCoverageName is empty..") Then Exit Function
    Dim FunctionName
    If SearchCoverageName Like "*Button*" Then
        
        FunctionName = ELookup("tblModelButtons", "ModelButtonID = " & RecordID, "FunctionName")
        DoCmd.OpenModule , FunctionName
        
    ElseIf SearchCoverageName Like "*Snippet*" Then
        
        DoCmd.OpenForm "frmSnippets", , , "SnippetID = " & RecordID
        
    ElseIf SearchCoverageName Like "*Models*" Then
        
        DoCmd.OpenForm "frmSeqModels", , , "SeqModelID = " & RecordID
        
    ElseIf SearchCoverageName Like "*Functions*" Then
        
        FunctionName = ELookup("tblCustomVBAFunctions", "CustomVBAFunctionID = " & RecordID, "CustomVBAFunction")
        DoCmd.OpenModule , FunctionName
        
    End If
    
End Function

Public Function GlobalSearchOpenMainForm(frm As Form)

    Dim SearchCoverageName: SearchCoverageName = frm("SearchCoverageName"): If ExitIfTrue(isFalse(SearchCoverageName), "SearchCoverageName is empty..") Then Exit Function
    
    Dim selectedText: selectedText = frm("SearchDescription").SelText
    
    Dim frmName
    If SearchCoverageName Like "*Button*" Then
        frmName = "mainModelButtons"
        DoCmd.OpenForm frmName
        If selectedText <> "" Then
            Set frm = Forms(frmName)
            frm("fltrWildSearch") = selectedText
            
            FilterSubform frm, 15
        End If
        
    ElseIf SearchCoverageName Like "*Snipp*" Then
        frmName = "mainSnippets"
        DoCmd.OpenForm frmName
        If selectedText <> "" Then
            Set frm = Forms(frmName)
            frm("fltrWildSearch") = selectedText
            
            SetSubformSQL frm, 627
        End If
        
    ElseIf SearchCoverageName Like "*Models*" Then
        frmName = "mainSeqModels"
        DoCmd.OpenForm frmName
        
    End If
    
End Function

