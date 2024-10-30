Attribute VB_Name = "SeqModelFilter Mod"
Option Compare Database
Option Explicit

Public Function SeqModelFilterCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm("SeqModelRelationshipID").RowSource = ""
            frm("SeqModelRelationshipID").TextAlign = 1
            frm("SeqModelFieldID").AfterUpdate = "=SeqModelFilterSeqModelFieldID_AfterUpdate([Form])"
            
            frm("VariableName").Name = "VariableName2"
            
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function GenerateModelFilterSnippet(frm As Object, Optional SeqModelFilterID = Null)
    
    RunCommandSaveRecord
    If IsNull(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
    End If
    
    If isFalse(SeqModelFilterID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName")
    Dim FilterOperator: FilterOperator = rs.fields("FilterOperator")
    Dim IsMultiple: IsMultiple = rs.fields("IsMultiple")
    Dim DataType: DataType = rs.fields("DataType")
    ''Get all the Likes filter type first
    If IsMultiple Then
        GenerateModelFilterSnippet = GetMultipleFilter(SeqModelFilterID)
        GoTo CopyClipboard:
    ElseIf DataType = "BOOLEAN" Then
        GenerateModelFilterSnippet = GetBooleanFilter(frm, SeqModelFilterID)
        GoTo CopyClipboard:
    ElseIf FilterOperator = "IsPresent" Then
        GenerateModelFilterSnippet = GetIsPresentFilter(frm, SeqModelFilterID)
        GoTo CopyClipboard:
    ElseIf FilterOperator = "Between" And DataType = "DATEONLY" Then
        GenerateModelFilterSnippet = GetDateBetweenFilter(frm, SeqModelFilterID)
        GoTo CopyClipboard:
    End If
    
    GenerateModelFilterSnippet = GetSingleFilter(SeqModelFilterID)
    GenerateModelFilterSnippet = GetGeneratedByFunctionSnippet(GenerateModelFilterSnippet, "GenerateModelFilterSnippet")
    
CopyClipboard:
    CopyToClipboard GenerateModelFilterSnippet
    
End Function

Private Function GetMultipleFilter(SeqModelFilterID) As String
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName")
    Dim FilterOperator: FilterOperator = rs.fields("FilterOperator")
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Multiple Filter")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[FilterQueryName]", FilterQueryName)
    replacedContent = replace(replacedContent, "[FilterOperator]", FilterOperator)
    replacedContent = replace(replacedContent, "[DatabaseFieldName]", DatabaseFieldName)
    
    GetMultipleFilter = replacedContent
    GetMultipleFilter = GetGeneratedByFunctionSnippet(GetMultipleFilter, "GetMultipleFilter", "Multiple Filter")
    
End Function

Private Function GetSingleFilter(SeqModelFilterID) As String
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName")
    Dim FilterOperator: FilterOperator = rs.fields("FilterOperator")
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If isFalse(DatabaseFieldName) Then DatabaseFieldName = ""
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Single Filter")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[FilterQueryName]", FilterQueryName)
    replacedContent = replace(replacedContent, "[FilterOperator]", FilterOperator)
    replacedContent = replace(replacedContent, "[DatabaseFieldName]", DatabaseFieldName)
    
    GetSingleFilter = replacedContent
    GetSingleFilter = GetGeneratedByFunctionSnippet(GetSingleFilter, "GetSingleFilter", "Single Filter")
    
End Function

Public Function GetLikeFilters(frm As Object, Optional SeqModelID = Null) As String
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    If isFalse(SeqModelID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterOperator = ""LIKE"""
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then Exit Function
    
    Dim fields As New clsArray
    
    Do Until rs.EOF
    
        Dim DatabaseFieldName: DatabaseFieldName = Esc(rs.fields("DatabaseFieldName"))
        fields.Add DatabaseFieldName
    
        rs.MoveNext
        
    Loop
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("LIKE Template")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[Fields]", "[" & fields.JoinArr(",") & "]")
    
    GetLikeFilters = replacedContent
    GetLikeFilters = GetGeneratedByFunctionSnippet(GetLikeFilters, "GetLikeFilters", "LIKE Template")
    
    CopyToClipboard GetLikeFilters
    
End Function

Public Function GetMatchFilters(frm As Object, Optional SeqModelID = Null) As String
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    If isFalse(SeqModelID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterOperator = ""Match"""
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then Exit Function
    
    Dim fields As New clsArray
    
    Do Until rs.EOF
    
        Dim DatabaseFieldName: DatabaseFieldName = Esc(rs.fields("DatabaseFieldName"))
        fields.Add DatabaseFieldName
    
        rs.MoveNext
        
    Loop
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Match Template")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[Fields]", "[" & fields.JoinArr(",") & "]")
    
    GetMatchFilters = replacedContent
    GetMatchFilters = GetGeneratedByFunctionSnippet(GetMatchFilters, "GetMatchFilters", "Match Template")
    
    CopyToClipboard GetMatchFilters
    
End Function

Public Function GenerateIndividualFilterControl(frm As Object, Optional SeqModelFilterID = Null)
    
    RunCommandSaveRecord
    If IsNull(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
    End If
    
    If isFalse(SeqModelFilterID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    

    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
    Dim DataType: DataType = rs.fields("DataType")
    Dim ListVariableName: ListVariableName = rs.fields("ListVariableName"): ''If ExitIfTrue(isFalse(ListVariableName), "ListVariableName is empty..") Then Exit Function
    Dim FilterCaption: FilterCaption = rs.fields("FilterCaption"): ''
    Dim PluralizedFieldName: PluralizedFieldName = rs.fields("PluralizedFieldName"): ''If ExitIfTrue(isFalse(PluralizedFieldName), "PluralizedFieldName is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface")
    Dim IsMultiple: IsMultiple = rs.fields("IsMultiple")
    
    Dim TemplateContent, replacedContent
    If ControlType = "Text" And FilterQueryName = "q" Then
        GenerateIndividualFilterControl = "<MUIText label=""Search"" name=""q"" id=""q"" sx={{ bgcolor: ""white"" }} size=""small"" />"
    ElseIf ControlType = "Text" And DataType = "DATEONLY" Then
        GenerateIndividualFilterControl = "<CustomDateRangePicker name=" & Esc(FilterQueryName) & " label=" & Esc(FilterCaption) & "/>"
    ElseIf ControlType = "Text" Then
        GenerateIndividualFilterControl = "<MUIText label=" & Esc(FilterCaption) & " name=" & Esc(FilterQueryName) & " id=" & Esc(FilterQueryName) & " sx={{ bgcolor: ""white"" }} size=""small"" />"
    ElseIf ControlType = "Option" And Not IsNull(ListVariableName) Then
        If ExitIfTrue(isFalse(FilterCaption), "FilterCaption is empty..") Then Exit Function
        TemplateContent = GetTemplateContent("Option Control Filter From Basic Model")
        replacedContent = replace(TemplateContent, "[ListVariableName]", ListVariableName)
        replacedContent = replace(replacedContent, "[FilterCaption]", FilterCaption)
        replacedContent = replace(replacedContent, "[FilterQueryName]", FilterQueryName)
        GenerateIndividualFilterControl = replacedContent
    ElseIf ControlType = "Option" And Not IsNull(PluralizedFieldName) And DataTypeInterface = "number" Then
        If ExitIfTrue(isFalse(FilterCaption), "FilterCaption is empty..") Then Exit Function
        TemplateContent = GetTemplateContent("Option Control Filter From Number Array")
        replacedContent = replace(TemplateContent, "[PluralizedFieldName]", PluralizedFieldName)
        replacedContent = replace(replacedContent, "[FilterCaption]", FilterCaption)
        replacedContent = replace(replacedContent, "[FilterQueryName]", FilterQueryName)
        GenerateIndividualFilterControl = replacedContent
    ElseIf ControlType = "Option" And Not IsNull(PluralizedFieldName) Then
        If ExitIfTrue(isFalse(FilterCaption), "FilterCaption is empty..") Then Exit Function
        TemplateContent = GetTemplateContent("Option Control Filter From Constant")
        replacedContent = replace(TemplateContent, "[PluralizedFieldName]", PluralizedFieldName)
        replacedContent = replace(replacedContent, "[FilterCaption]", FilterCaption)
        replacedContent = replace(replacedContent, "[FilterQueryName]", FilterQueryName)
        GenerateIndividualFilterControl = replacedContent
    ElseIf ControlType = "Autocomplete" Then
        TemplateContent = GetTemplateContent("Autocomplete Control")
        If ExitIfTrue(isFalse(ListVariableName), "ListVariableName is empty..") Then Exit Function
        If ExitIfTrue(isFalse(FilterCaption), "FilterCaption is empty..") Then Exit Function
        If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
        replacedContent = replace(TemplateContent, "[ListVariableName]", ListVariableName)
        replacedContent = replace(replacedContent, "[FilterCaption]", FilterCaption)
        replacedContent = replace(replacedContent, "[FilterQueryName]", FilterQueryName)
        replacedContent = replace(replacedContent, "[IsMultiple]", IIf(IsMultiple, "true", "false"))
        
        GenerateIndividualFilterControl = replacedContent
    ElseIf ControlType = "Switch" Then
        If ExitIfTrue(isFalse(FilterCaption), "FilterCaption is empty..") Then Exit Function
        GenerateIndividualFilterControl = "<MUISwitch label=" & Esc(FilterCaption) & " name=" & Esc(FilterQueryName) & " />"
    
    End If
    
    Dim lines As New clsArray
    lines.Add "{/* Generated by GenerateIndividualFilterControl */}"
    lines.Add GenerateIndividualFilterControl
    
    GenerateIndividualFilterControl = lines.JoinArr(vbNewLine)
    
    CopyToClipboard GenerateIndividualFilterControl
    
End Function

Public Function SeqModelFilterSeqModelFieldID_AfterUpdate(frm As Form)
    
    Dim SeqModelFieldID: SeqModelFieldID = frm("SeqModelFieldID"): If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    Dim DatabaseFieldName As String: DatabaseFieldName = frm("SeqModelFieldID").Column(1): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    
    Dim FilterCaption: FilterCaption = ConvertToVerboseCaption(DatabaseFieldName)
    
    frm("FilterQueryName") = DatabaseFieldName
    frm("FilterCaption") = FilterCaption
    
End Function

Public Function GenerateSimpleFilterFieldSnippet(frm As Object, Optional SeqModelFilterID = Null)
    
    RunCommandSaveRecord
    If IsNull(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
    Dim FilterOperator: FilterOperator = rs.fields("FilterOperator"): If ExitIfTrue(isFalse(FilterOperator), "FilterOperator is empty..") Then Exit Function
    Dim fieldName: fieldName = rs.fields("FieldName")
    If FilterQueryName = "q" Then
        sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterQueryName = ""q"" ORDER BY FilterOrder"
        Set rs = ReturnRecordset(sqlStr)
        Dim fields As New clsArray
        Do Until rs.EOF
            fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
            fields.Add "{" & fieldName & ": { [Op.like]: `%${query.q}%` } }"
            rs.MoveNext
        Loop
        GenerateSimpleFilterFieldSnippet = "if (query.q) { where[Op.or] = [ " & fields.JoinArr("," & vbNewLine) & " ]; }"
    ElseIf IsNull(fieldName) Then
        GenerateSimpleFilterFieldSnippet = ""
    ElseIf FilterOperator = "=" Then
        GenerateSimpleFilterFieldSnippet = "if (query." & FilterQueryName & ") {where." & fieldName & " = query." & FilterQueryName & ";}"
    End If
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateSimpleFilterFieldSnippet"
    lines.Add GenerateSimpleFilterFieldSnippet
    GenerateSimpleFilterFieldSnippet = lines.NewLineJoin
    
    CopyToClipboard GenerateSimpleFilterFieldSnippet
    
End Function

Public Function GetDateBetweenFilter(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GetDateBetweenFilter = GetReplacedTemplate(rs, "Date In Between Filter")
    GetDateBetweenFilter = GetGeneratedByFunctionSnippet(GetDateBetweenFilter, "GetDateBetweenFilter", "Date In Between Filter")
    CopyToClipboard GetDateBetweenFilter
    
End Function

Public Function GetThisFilterInterface(frm As Object, Optional SeqModelFilterID = "", Optional forceStringType As Boolean = False)

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    Dim filterType: filterType = "string"
    Dim MultipleControls As New clsArray: MultipleControls.arr = "FacetedControl,CheckboxGroup"
    
    Dim templateName As String
    
    If ControlType = "DateRangePicker" Then
        templateName = "GetThisBetweenFilterInterface"
        GetThisFilterInterface = GetReplacedTemplate(rs, templateName)
        GoTo ReallyEndThisFunction
    ElseIf ControlType = "Select" Then
        Dim fieldValue As New clsArray:  fieldValue.arr = Elookups("tblSeqModelFilterOptions", "SeqModelFilterID = " & SeqModelFilterID, "FieldValue", "SeqModelFilterOptionID")
        If fieldValue.count > 0 Then
            fieldValue.EscapeItems
            fieldValue.Add Esc("")
            filterType = fieldValue.JoinArr(" | ")
        End If
    ElseIf forceStringType Then
        GoTo EndThisFunction
    ElseIf ControlType = "Switch" Then
        filterType = "boolean"
        GoTo EndThisFunction
    ElseIf MultipleControls.InArray(ControlType) Then
        filterType = "string[]"
        GoTo EndThisFunction
    End If
    
EndThisFunction:
    GetThisFilterInterface = FilterQueryName & ": " & filterType & ";"
ReallyEndThisFunction:
    GetThisFilterInterface = GetGeneratedByFunctionSnippet(GetThisFilterInterface, "GetThisFilterInterface", templateName, , True)
    CopyToClipboard GetThisFilterInterface
    
End Function

Private Function GetModelFilterOptions(SeqModelFilterID)
    
    
End Function

Public Function GetModelFilterDefault(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim IsMultiple: IsMultiple = rs.fields("IsMultiple")
    Dim IsBoolean: IsBoolean = rs.fields("IsBoolean")
    
    Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    Dim DefaultValue: DefaultValue = Esc("")
    
    Dim templateName As String
    
    If ControlType = "DateRangePicker" Then
        templateName = "GetModelFilterDateBetweenDefault"
        GetModelFilterDefault = GetReplacedTemplate(rs, templateName)
        GoTo EndThisFunction
    ElseIf IsBoolean Then
        DefaultValue = "false"
    ElseIf IsMultiple Then
        DefaultValue = "[]"
    End If
    
    GetModelFilterDefault = FilterQueryName & ":" & DefaultValue & ","
EndThisFunction:
    GetModelFilterDefault = GetGeneratedByFunctionSnippet(GetModelFilterDefault, "GetModelFilterDefault", templateName, , True)
    
End Function

Public Function GetFilterQueryName(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
    
    Dim templateName As String
    If ControlType = "DateRangePicker" Then
        templateName = "GetDateBetweenFilterQueryName"
        GetFilterQueryName = GetReplacedTemplate(rs, templateName)
    Else
        GetFilterQueryName = FilterQueryName & ","
    End If
    
    GetFilterQueryName = GetGeneratedByFunctionSnippet(GetFilterQueryName, "GetFilterQueryName", templateName, , True)
    CopyToClipboard GetFilterQueryName
    
End Function

Public Function SeqModelFilter_OnCurrent(frm As Form)
    
    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    Dim sqlStr: sqlStr = "SELECT SeqModelRelationshipID,RightModelName FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " ORDER BY RightModelName"
    
    If Not isFalse(SeqModelID) Then
        frm("SeqModelRelationshipID").RowSource = sqlStr
    End If
    
End Function

Public Function GetBooleanFilter(frm As Object, Optional SeqModelFilterID = "")
    
    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetBooleanFilter = GetReplacedTemplate(rs, "Boolean Filter")
    GetBooleanFilter = GetGeneratedByFunctionSnippet(GetBooleanFilter, "GetBooleanFilter", "Boolean Filter")
    CopyToClipboard GetBooleanFilter

End Function

Public Function GetIsPresentFilter(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetIsPresentFilter = GetReplacedTemplate(rs, "isPresent Filter")
    GetIsPresentFilter = GetGeneratedByFunctionSnippet(GetIsPresentFilter, "GetIsPresentFilter", "isPresent Filter")
    CopyToClipboard GetIsPresentFilter
    
End Function

Public Function GetFormikFilterControl(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    
    If ControlType = "Switch" Then
        GetFormikFilterControl = GetFilterSwitchControl(frm, SeqModelFilterID)
    ElseIf ControlType = "FacetedControl" Then
        GetFormikFilterControl = GetFacetedControl(frm, SeqModelFilterID)
    ElseIf ControlType = "Combobox" Then
        GetFormikFilterControl = GetComboBoxControl(frm, SeqModelFilterID)
    ElseIf ControlType = "Select" Then
        GetFormikFilterControl = GetSelectFilter(frm, SeqModelFilterID)
    ElseIf ControlType = "DateRangePicker" Then
        GetFormikFilterControl = GetDateRangeFilter(frm, SeqModelFilterID)
    End If
   
EndThisFunction:
    CopyToClipboard GetFormikFilterControl
    
End Function

Public Function GetFilterSwitchControl(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetFilterSwitchControl = GetReplacedTemplate(rs, "Filter Switch Control")
    GetFilterSwitchControl = "{/* Generated by GetFilterSwitchControl */}" & GetFilterSwitchControl
    CopyToClipboard GetFilterSwitchControl
    
End Function

Public Function GetSearchParamVariable(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    
    Dim templateName: templateName = "Search Param Variable"
    
    If ControlType = "DateRangePicker" Then
        templateName = "GetDateBetweenSearchParams"
    End If
    
    GetSearchParamVariable = GetReplacedTemplate(rs, templateName)
    
    If isPresent("tblSeqModelFilterOptions", "SeqModelFilterID = " & SeqModelFilterID) Then
        GetSearchParamVariable = replace(GetSearchParamVariable, " || """"", "")
    End If
    
    GetSearchParamVariable = GetGeneratedByFunctionSnippet(GetSearchParamVariable, "GetSearchParamVariable", templateName, , True)
    CopyToClipboard GetSearchParamVariable
    
End Function

Public Function GetFacetedControl(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GetFacetedControl = GetReplacedTemplate(rs, "GetFacetedControl")
    GetFacetedControl = replace(GetFacetedControl, "[GetOptionOrModeList]", GetOptionOrModeList(frm, SeqModelFilterID))
    
    GetFacetedControl = "{/* Generated by GetFacetedControl */}" & GetFacetedControl
    CopyToClipboard GetFacetedControl
    
End Function

Public Function GetComboBoxControl(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim DataType: DataType = rs.fields("DataType")
    Dim ModelListID: ModelListID = rs.fields("ModelListID")
    Dim GetControlListName
    
    If DataType = "ENUM" Then
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        GetControlListName = "CONTROL_OPTIONS." & fieldName
    End If
    
    If Not IsNull(ModelListID) Then
        Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
        GetuseModelListts frm, ModelListID
        GetControlListName = VariableName & "List || []"
    End If
    
    GetComboBoxControl = GetReplacedTemplate(rs, "GetComboBoxControl")
    GetComboBoxControl = replace(GetComboBoxControl, "[GetControlListName]", GetControlListName)
    GetComboBoxControl = GetGeneratedByFunctionSnippet(GetComboBoxControl, "GetComboBoxControl", "GetComboBoxControl", True)
    CopyToClipboard GetComboBoxControl
    
End Function

Public Function GetRequiredQueryFromTanstack(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetRequiredQueryFromTanstack = GetReplacedTemplate(rs, "GetRequiredQueryFromTanstack")
    GetRequiredQueryFromTanstack = GetGeneratedByFunctionSnippet(GetRequiredQueryFromTanstack, "GetRequiredQueryFromTanstack")
    CopyToClipboard GetRequiredQueryFromTanstack
    
End Function

Public Function ImportUseModelListHook(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    ImportUseModelListHook = GetReplacedTemplate(rs, "ImportUseModelListHook")
    ImportUseModelListHook = GetGeneratedByFunctionSnippet(ImportUseModelListHook, "ImportUseModelListHook", "ImportUseModelListHook", , True)
    CopyToClipboard ImportUseModelListHook
    
End Function

Public Function GetBackendFilter(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    
    Dim IsMultiple: IsMultiple = rs.fields("IsMultiple")
    Dim templateName As String
    
    If IsMultiple Then
        templateName = "In Filter"
    Else
        templateName = "Equality Filter"
    End If
    
    GetBackendFilter = GetReplacedTemplate(rs, templateName)
    GetBackendFilter = GetGeneratedByFunctionSnippet(GetBackendFilter, "GetBackendFilter", "templateName")
    CopyToClipboard GetBackendFilter
    
End Function

Public Function GetOptionOrModeList(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName")
    Dim VariableName: VariableName = rs.fields("VariableName")
    
    Dim ModelListID: ModelListID = rs.fields("ModelListID")
    
    If IsNull(ModelListID) Then
        GetOptionOrModeList = "CONTROL_OPTIONS." & FilterQueryName
    Else
        GetOptionOrModeList = VariableName & "List || []"
        GetuseModelListts frm, ModelListID
    End If
    
    CopyToClipboard GetOptionOrModeList
    
End Function

Public Function GetFormOptionOrModeList(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim fieldName: fieldName = rs.fields("FieldName")
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName")
    Dim RelatedModelID: RelatedModelID = rs.fields("RelatedModelID")
    Dim RelatedVariableName: RelatedVariableName = rs.fields("RelatedVariableName")
    
    If Not IsNull(RelatedModelID) Then
        GetFormOptionOrModeList = RelatedVariableName & "List || []"
        CopyToClipboard GetFormOptionOrModeList
        Exit Function
    End If
    
    ''Find Related Model
    Dim ModelListID: ModelListID = ELookup("tblSeqModelRelationships", "LeftForeignKey = " & Esc(DatabaseFieldName) & " AND LeftModelID = " & SeqModelID, "RightModelID")
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftForeignKey = " & Esc(DatabaseFieldName) & " AND LeftModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then
        GetFormOptionOrModeList = "CONTROL_OPTIONS." & fieldName
    Else
        Dim RightModelID: RightModelID = rs.fields("RightModelID"): If ExitIfTrue(isFalse(RightModelID), "RightModelID is empty..") Then Exit Function
        Dim RightVariableName: RightVariableName = rs.fields("RightVariableName"): If ExitIfTrue(isFalse(RightVariableName), "RightVariableName is empty..") Then Exit Function
        GetFormOptionOrModeList = RightVariableName & "List || []"
        GetuseModelListts frm, RightModelID
    End If
    
    CopyToClipboard GetFormOptionOrModeList
    
End Function

Public Function GetSelectFilter(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSelectFilter = GetReplacedTemplate(rs, "Select Filter")
    GetSelectFilter = replace(GetSelectFilter, "[GetOptionOrModeList]", GetOptionOrModeList(frm, SeqModelFilterID))
    GetSelectFilter = "{/* Generated by GetSelectFilter */}" & GetSelectFilter
    CopyToClipboard GetSelectFilter
    
End Function

Public Function GetDatabaseFieldName(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetDatabaseFieldName = GetReplacedTemplate(rs, "GetDatabaseFieldName")
    CopyToClipboard GetDatabaseFieldName
    
End Function

Public Function GetQueryKeyValueOfGetPluralizedModelName(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    
    Dim templateName: templateName = "GetQueryKeyValueOfGetPluralizedModelName"
    
    If ControlType = "DateRangePicker" Then
        templateName = "GetQueryKVPairDateRange"
    End If
    
    GetQueryKeyValueOfGetPluralizedModelName = GetReplacedTemplate(rs, templateName)
    GetQueryKeyValueOfGetPluralizedModelName = GetGeneratedByFunctionSnippet(GetQueryKeyValueOfGetPluralizedModelName, "GetQueryKeyValueOfGetPluralizedModelName", templateName, , True)
    CopyToClipboard GetQueryKeyValueOfGetPluralizedModelName
    
End Function

Public Function GetAllFilterManualOption(frm As Object, Optional SeqModelFilterID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim template: template = GetReplacedTemplate(rs, "GetFilterWithManualOption")
    
    
    sqlStr = "SELECT SeqModelFilterOptionID FROM tblSeqModelFilterOptions WHERE SeqModelFilterID = " & SeqModelFilterID & _
        " ORDER BY SeqModelFilterOptionID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterOptionID: SeqModelFilterOptionID = rs.fields("SeqModelFilterOptionID")
        lines.Add GetFilterManualOption(frm, SeqModelFilterOptionID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllFilterManualOption = lines.JoinArr(vbNewLine)
        GetAllFilterManualOption = GetGeneratedByFunctionSnippet(GetAllFilterManualOption, "GetAllFilterManualOption")
        GetAllFilterManualOption = replace(template, "[GetAllFilterManualOption]", GetAllFilterManualOption)
    End If
    
    GetAllFilterManualOption = GetGeneratedByFunctionSnippet(GetAllFilterManualOption, "GetAllFilterManualOption", "GetFilterWithManualOption")
    CopyToClipboard GetAllFilterManualOption
    
End Function

Public Function GetDateRangeFilter(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetDateRangeFilter = GetReplacedTemplate(rs, "GetDateRangeFilter")
    GetDateRangeFilter = GetGeneratedByFunctionSnippet(GetDateRangeFilter, "GetDateRangeFilter", "GetDateRangeFilter", True)
    CopyToClipboard GetDateRangeFilter
End Function

Public Function GetSeqModelFilterKeys(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SeqModelFilterOptionKeys: SeqModelFilterOptionKeys = """options"": [" & GetAllSeqModelFilterOptionKeys(frm, SeqModelFilterID) & "],"
    
    lines.Add "{"
    lines.Add GetKVPairs("qrySeqModelFilters", rs)
    lines.Add SeqModelFilterOptionKeys
    lines.Add "}"

    GetSeqModelFilterKeys = lines.JoinArr(vbNewLine) & ","
    GetSeqModelFilterKeys = replace(GetSeqModelFilterKeys, "fieldValue: null", "fieldValue: ""null""")
    
End Function

Public Function GetAllSeqModelFilterOptionKeys(frm As Object, Optional SeqModelFilterID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFilterOptionID FROM tblSeqModelFilterOptions WHERE SeqModelFilterID = " & SeqModelFilterID & _
        " ORDER BY SeqModelFilterOptionID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterOptionID: SeqModelFilterOptionID = rs.fields("SeqModelFilterOptionID")
        lines.Add GetSeqModelFilterOptionKeys(frm, SeqModelFilterOptionID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSeqModelFilterOptionKeys = lines.JoinArr(vbNewLine)
        GetAllSeqModelFilterOptionKeys = GetGeneratedByFunctionSnippet(GetAllSeqModelFilterOptionKeys, "GetAllSeqModelFilterOptionKeys")
        CopyToClipboard GetAllSeqModelFilterOptionKeys
    End If

End Function

Public Function GetIndextStatementByFilter(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetIndextStatementByFilter = GetReplacedTemplate(rs, "Get Index Statement By Filter")
    ''GetIndextStatementByFilter = GetGeneratedByFunctionSnippet(GetIndextStatementByFilter, "GetIndextStatementByFilter", "Get Index Statement By Filter")
    CopyToClipboard GetIndextStatementByFilter
End Function
