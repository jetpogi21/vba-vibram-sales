Attribute VB_Name = "MainFormSQL Mod"
Option Compare Database
Option Explicit

Public Function MainFormSQLCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function MainFormSQLOnLoad(frm As Object, ModelID)

    DefaultMainFormLoad frm
    SetSubformSQL frm, ModelID
    
End Function


''This is a function that will get the sql string to be used in a mainform's table instead of the usual filtering
Public Function SetSubformSQL(frm As Object, ModelID)

    ''Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    Dim MainTable: MainTable = GetTableNameFromModelID(ModelID)
    
    ''Then the fields
    Dim fields: fields = GetFields(MainTable, , True)
    ''The filters => From the GetFilterArray also include the filter from the related table
    Dim filters As New clsArray: Set filters = GetFilterArray(frm, ModelID, True)
    ''Debug.Print filters.JoinArr(" AND ")
    ''The joins
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    ''tblRelatedFilterFields
    ''RelatedFilterFieldID,ModelID,TableName,MainConnectorField,SubConnectorField,FieldToUse,Timestamp,CreatedBy,RecordImportID,FilterOrder,IsOptionGroup,FilterOperation
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = MainTable
        If filters.count > 0 Then .AddFilter filters.JoinArr(" AND ")
        .fields = "DISTINCTROW " & fields
    End With
    
    ''Set the joins from the related field..
    GetSQLAndJoins frm, ModelID, sqlObj
    sqlStr = sqlObj.sql
    
    frm("subform").Form.recordSource = sqlStr
    If ControlExists("fltrWildSearch", frm) Then
        frm("fltrWildSearch").Requery
    End If
    
End Function

Private Function GetSQLAndJoins(frm As Object, ModelID, sqlObj)
    
    Dim sqlObj_1 As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblRelatedFilterFields WHERE ModelID = " & ModelID)
    
    Dim RelatedFilterFieldID, TableName, MainConnectorField, SubConnectorField, FieldToUse, FilterOrder, IsOptionGroup, FilterOperation
    Dim RightJoinKey, LeftJoinkey, IncludeInWildcardSearch, SatisfiesFilter

    Do Until rs.EOF
    
        RelatedFilterFieldID = rs.fields("RelatedFilterFieldID")
        TableName = rs.fields("TableName")
        MainConnectorField = rs.fields("MainConnectorField")
        SubConnectorField = rs.fields("SubConnectorField")
        FieldToUse = rs.fields("FieldToUse")
        FilterOrder = rs.fields("FilterOrder")
        IsOptionGroup = rs.fields("IsOptionGroup")
        FilterOperation = rs.fields("FilterOperation")
        LeftJoinkey = rs.fields("LeftJoinKey")
        RightJoinKey = rs.fields("RightJoinKey")
        IncludeInWildcardSearch = rs.fields("IncludeInWildcardSearch")
        SatisfiesFilter = rs.fields("SatisfiesFilter")
        
        ''Next field if the cb for the not FieldToUse is unchecked
        Dim cbValue
        If isFalse(FieldToUse) Then
            Dim cbName: cbName = "fltr" & TableName & SatisfiesFilter
            cbValue = frm(cbName)
            If Not cbValue Then GoTo NextField
        End If
        
        Dim filterStatement: filterStatement = ""
        Dim JoinType: JoinType = "INNER"
        ''Check if this is not a wild card search field, if it is not then
        ''the subform field is important. No need to get the value of the wildcard search here
        ''since it will be computed from the main table..
        ''Just need to show the field from the subquery so that the field can be used appropriately
        ''Add another check -> if FieldToUse is null then the subform field isn't necessary
        If Not IncludeInWildcardSearch And Not isFalse(FieldToUse) Then
            Dim SubformName: SubformName = "fltr" & FieldToUse
            Dim filterValues As New clsArray: Set filterValues = GetFilterValues(frm, SubformName)
            filterStatement = GetFilterStatement(filterValues, TableName, FieldToUse, FilterOperation)
        End If
        
        If isFalse(filterStatement) Then JoinType = "LEFT"
        
        ''another jointype for the satisfies filter
        ''no need to apply main table filter if this is a filter satisfaction only
        If SatisfiesFilter And isFalse(FieldToUse) Then JoinType = "INNER"
        ''apply main table filter if this is a filter unsatisfaction
        If Not SatisfiesFilter And isFalse(FieldToUse) Then JoinType = "LEFT"
        
        ''If This is included in Wildsearch the FieldToUse is important so that the main filter
        ''can compute the right SQL
        Dim fields: fields = "DISTINCT " & RightJoinKey
        If IncludeInWildcardSearch Then fields = "DISTINCTROW " & RightJoinKey & "," & FieldToUse
        
        Dim subQueryAlias: subQueryAlias = "temp" & TableName
        If SatisfiesFilter And isFalse(FieldToUse) Then subQueryAlias = "temp" & TableName & "Satisfaction"
        If Not SatisfiesFilter And isFalse(FieldToUse) Then subQueryAlias = "temp" & TableName & "UnSatisfaction"
        
        Set sqlObj_1 = New clsSQL
        With sqlObj_1
            .Source = TableName
            .fields = fields
            If Not isFalse(filterStatement) Then .AddFilter filterStatement
            sqlStr = .sql
        End With
        
        ''tempTableName is important since it will be refered to by the main filter
        sqlObj.joins.Add GenerateJoinObj(sqlStr, LeftJoinkey, subQueryAlias, RightJoinKey, JoinType)
NextField:
        rs.MoveNext
        
    Loop
    
End Function

Private Function GetFilterStatement(filterValues As clsArray, TableName, FieldToUse, FilterOperation) As String
    
    Dim filters As New clsArray
    Dim filterValue
    
    ''Get the fieldType of the FieldToUse from filterValue
    Dim Numbers As New clsArray: Numbers.arr = "3,4,5,6,7"
    
    If filterValues.count > 0 Then
        For Each filterValue In filterValues.arr
            If Not Numbers.InArray(CStr(GetFieldTypeFromRS(TableName, FieldToUse))) Then
                filterValue = EscapeString(filterValue)
            End If
            filters.Add TableName & "." & FieldToUse & " " & FilterOperation & " " & filterValue
        Next filterValue
    End If
    GetFilterStatement = filters.JoinArr(" OR ")
    
End Function

Private Function GetFilterValues(frm As Object, SubformName) As clsArray
    
    Dim subfrm As Form: Set subfrm = frm(SubformName).Form
    subfrm.Requery
    Dim rs As Recordset
    Set rs = subfrm.RecordsetClone
    
    If rs.recordCount > 0 Then rs.MoveFirst
    Dim filterValues As New clsArray
    Dim id, FilterLabel, Selected, value
    Do Until rs.EOF
        
        id = rs.fields("ID")
        FilterLabel = rs.fields("FilterLabel")
        Selected = rs.fields("Selected")
        value = rs.fields("Value")
        
        If Selected Then
            filterValues.Add value
        End If
        
        rs.MoveNext
    Loop
    
    Set GetFilterValues = filterValues
    
End Function
