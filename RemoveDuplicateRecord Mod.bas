Attribute VB_Name = "RemoveDuplicateRecord Mod"
Option Compare Database
Option Explicit

Public Function RemoveDuplicateRecordCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

''A sub procedure that will remove any duplicated records from a table
''At least 1 of the duplicates will remain
Public Function RemoveDuplicateRecords(frm As Form)
    
    Dim tblName, fields, pkName
    tblName = frm("TableName")
    fields = frm("Fields")
    pkName = frm("PrimaryKey")
    
    Dim fieldsArr As New clsArray: fieldsArr.arr = fields
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT " & fields & ",Count(" & pkName & ") AS RecordCount FROM " & tblName & " GROUP BY " & fields & _
                                    " HAVING Count(" & pkName & ") > 1")

    Do Until rs.EOF
        ''Get the first record matching the recordset
        Dim filterStatement: filterStatement = BuildFilter(fieldsArr, rs, tblName)
        Dim pkValue: pkValue = ELookup(tblName, filterStatement, pkName)
        RunSQL "DELETE FROM " & tblName & " WHERE " & filterStatement & " AND " & pkName & " <> " & pkValue
        rs.MoveNext
    Loop
End Function


Private Function BuildFilter(fieldsArr As clsArray, rs As Recordset, tblName) As String
    
    Dim field, filters As New clsArray
    For Each field In fieldsArr.arr
        filters.Add field & " = " & EscapeString(rs.fields(field), tblName, field)
    Next field
    
    BuildFilter = filters.JoinArr(" AND ")
    
End Function

