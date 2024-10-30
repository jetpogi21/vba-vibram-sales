Attribute VB_Name = "TableStat Mod"
Option Compare Database
Option Explicit

Public Function TableStatCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function EnumerateTableStats()

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim TableName As String
    Dim recordCount As Long
    
    ' Set the current database object
    Set db = CurrentDb
    
    Dim fields As New clsArray: fields.arr = "TableName,Records"
    Dim fieldValues As New clsArray
    
    RunSQL "DELETE FROM tblTableStats"
    
    ' Loop through all tables in the database
    For Each tdf In db.TableDefs
        TableName = tdf.Name
        If Not TableName Like "~TMP*" And Not TableName Like "MSyS*" Then ' Exclude temporary tables
            Dim rs As Recordset: Set rs = ReturnRecordset(tdf.Name)
            recordCount = CountRecordset(rs)
            
            Set fieldValues = New clsArray
            fieldValues.Add TableName
            fieldValues.Add recordCount
            UpsertRecord "tblTableStats", fields, fieldValues
            
        End If
    Next tdf
    
    ' Clean up
    Set tdf = Nothing
    Set db = Nothing
    
    Dim frm As Form: Set frm = GetForm("mainTableStats")
    If Not frm Is Nothing Then
        frm("subform").Form.Requery
    End If
    
End Function

