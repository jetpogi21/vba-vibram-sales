Attribute VB_Name = "Query Mod"
Option Compare Database
Option Explicit

Public Function QueryCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function QueryTemplate(QueryName)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblQueries WHERE QueryName = " & EscapeString(QueryName))
    
    If ExitIfTrue(rs.EOF, "There is no valid record...") Then Exit Function
    
    ''TABLE: tblQueries Fields: QueryID|QueryName|QueryFunction|QueryDescription|RelatedRecordsets
    Dim RelatedRecordsetArr As New clsArray, i, j As String
    RelatedRecordsetArr.arr = rs.fields("RelatedRecordsets")
    
    Dim clpbrdArr As New clsArray
    For Each i In RelatedRecordsetArr.arr
        j = i
        clpbrdArr.Add "''" & PrintFields(j)
    Next i
    
    Dim QueryFunction: QueryFunction = rs.fields("QueryFunction")
    clpbrdArr.Add "Public Function " & QueryFunction & "(Optional fltrStr = """")"
    clpbrdArr.Add ""
    clpbrdArr.Add GetSQLTemplate
    clpbrdArr.Add ""
    clpbrdArr.Add "End Function"
    
    CopyToClipboard clpbrdArr.JoinArr(vbCrLf)
    
End Function


Public Function GetFields(rsName, Optional except = "", Optional prefix = False, Optional isSQL = False)
    
    Dim sqlStr: sqlStr = rsName
    If Not isSQL Then sqlStr = "SELECT * FROM " & sqlStr
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fldArr As New clsArray, exceptArr As New clsArray, fld As field
    If except <> "" Then exceptArr.arr = except
    
    For Each fld In rs.fields
        If Not exceptArr.InArray(fld.Name) Then
            Dim fldName: fldName = fld.Name
            If prefix Then fldName = rsName & "." & fldName
            fldArr.Add fldName
        End If
    Next fld

    GetFields = fldArr.JoinArr(",")
    
End Function

Public Function array_As_sum(str)

    Dim item, newArr As New clsArray, arr As New clsArray
    arr.arr = str
    For Each item In arr.arr
        newArr.Add "CdblNz(Sum(" & item & ")) AS Sum" & item
    Next item
    
    array_As_sum = newArr.JoinArr(",")
    
End Function

Public Function array_as_nz(str, Optional useSum = False)

    Dim arr As New clsArray: arr.arr = str
    Dim i, newArr As New clsArray
    For Each i In arr.arr
        If useSum Then
            newArr.Add "CdblNz(" & i & ") AS z" & i
        Else
            newArr.Add "CdblNz(Sum" & i & ") AS " & i
        End If
        
    Next i
    
    array_as_nz = newArr.JoinArr(",")
    
End Function

Private Function IsTable(rsName) As Boolean

    Dim tblDef As TableDef
    On Error GoTo ErrHandler:
    Set tblDef = CurrentDb.TableDefs(rsName)
    IsTable = True
    Exit Function
ErrHandler:
    IsTable = False
    
End Function
