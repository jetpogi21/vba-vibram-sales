Attribute VB_Name = "SeqModelFieldGroup Mod"
Option Compare Database
Option Explicit

Public Function SeqModelFieldGroupCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


Public Function GetSeqModelFieldGroups(frm As Object, Optional SeqModelFieldGroupID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldGroupID) Then
        SeqModelFieldGroupID = frm("SeqModelFieldGroupID")
        If ExitIfTrue(isFalse(SeqModelFieldGroupID), "SeqModelFieldGroupID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFieldGroups WHERE SeqModelFieldGroupID = " & SeqModelFieldGroupID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSeqModelFieldGroups = "{" & vbNewLine & GetKVPairs("qrySeqModelFieldGroups", rs) & vbNewLine & "},"
    
End Function

