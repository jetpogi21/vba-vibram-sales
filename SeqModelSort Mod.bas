Attribute VB_Name = "SeqModelSort Mod"
Option Compare Database
Option Explicit

Public Function SeqModelSortCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SeqModelSortSeqModelFieldID_AfterUpdate(frm As Form)
    
    Dim SeqModelFieldID: SeqModelFieldID = frm("SeqModelFieldID"): If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    Dim DatabaseFieldName As String: DatabaseFieldName = frm("SeqModelFieldID").Column(1): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    
    Dim VerboseFieldName: VerboseFieldName = ConvertToVerboseCaption(DatabaseFieldName)
    
    frm("ModelFieldCaption") = VerboseFieldName
    
End Function
Public Function GetSeqModelSortKeys(frm As Object, Optional SeqModelSortID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelSortID) Then
        SeqModelSortID = frm("SeqModelSortID")
        If ExitIfTrue(isFalse(SeqModelSortID), "SeqModelSortID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelSorts WHERE SeqModelSortID = " & SeqModelSortID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSeqModelSortKeys = "{" & vbNewLine & GetKVPairs("qrySeqModelSorts", rs) & vbNewLine & "},"
    
End Function

Public Function GetSeqModelHookKeys(frm As Object, Optional SeqModelHookID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelHookID) Then
        SeqModelHookID = frm("SeqModelHookID")
        If ExitIfTrue(isFalse(SeqModelHookID), "SeqModelHookID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelHooks WHERE SeqModelHookID = " & SeqModelHookID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSeqModelHookKeys = "{" & vbNewLine & GetKVPairs("qrySeqModelHooks", rs) & vbNewLine & "},"
    
End Function
