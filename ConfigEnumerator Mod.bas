Attribute VB_Name = "ConfigEnumerator Mod"
Option Compare Database
Option Explicit

Public Function ConfigEnumeratorCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function GetConfigEnumeratorFields(frm As Form)
    
    RunCommandSaveRecord
    Dim ConfigEnumeratorID: ConfigEnumeratorID = frm("ConfigEnumeratorID"): If isFalse(ConfigEnumeratorID) Then Exit Function
    Dim QueryName: QueryName = frm("QueryName"): If ExitIfTrue(isFalse(QueryName), "Query Name is empty..") Then Exit Function
    Dim BaseTable: BaseTable = frm("BaseTable"): If ExitIfTrue(isFalse(BaseTable), "Base Table is empty..") Then Exit Function
    
    Dim fields As New clsArray, fieldVals As New clsArray
    fields.arr = "ConfigEnumeratorID,FieldName,VariableName,BaseField,Exclude"
    
    Dim excluded As New clsArray: excluded.arr = "CreatedBy,Timestamp,RecordImportID"
    
    Dim fld As field
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM " & QueryName)
    For Each fld In rs.fields
        Dim fieldName: fieldName = fld.Name
        Dim VariableName: VariableName = FirstCharLowercase(fieldName)
        Dim BaseField: BaseField = IsFieldInBaseTable(fieldName, BaseTable)
        Dim Exclude
        If excluded.InArray(fieldName) Then
            Exclude = -1
        Else
            Exclude = Not BaseField
        End If

        Set fieldVals = New clsArray
        fieldVals.Add ConfigEnumeratorID
        fieldVals.Add Esc(fieldName)
        fieldVals.Add Esc(VariableName)
        fieldVals.Add BaseField
        fieldVals.Add Exclude
        
        RunSQL "INSERT INTO tblConfigEnumeratorFields (" & fields.JoinArr & ") VALUES (" & fieldVals.JoinArr(",") & ")"
        
    Next fld
    
    If DoesPropertyExists(frm, "subConfigEnumeratorFields") Then
        frm("subConfigEnumeratorFields").Form.Requery
    End If
End Function

Public Function IsFieldInBaseTable(fieldName, BaseTable) As Boolean
    Dim rsBase As Recordset: Set rsBase = ReturnRecordset("SELECT * FROM " & BaseTable)
    Dim fld As field

    ' Check if the field is in BaseTable
    For Each fld In rsBase.fields
        If fld.Name = fieldName Then
            ' The field is in BaseTable
            IsFieldInBaseTable = True
            Exit Function
        End If
    Next fld

    ' If we've gotten this far, the field is not in BaseTable
    IsFieldInBaseTable = False
End Function



