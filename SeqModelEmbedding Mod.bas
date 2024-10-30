Attribute VB_Name = "SeqModelEmbedding Mod"
Option Compare Database
Option Explicit

Public Function SeqModelEmbeddingCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function GetSeqModelEmbeddings(frm As Object, Optional SeqModelEmbeddingID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelEmbeddingID) Then
        SeqModelEmbeddingID = frm("SeqModelEmbeddingID")
        If ExitIfTrue(isFalse(SeqModelEmbeddingID), "SeqModelEmbeddingID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelEmbeddings WHERE SeqModelEmbeddingID = " & SeqModelEmbeddingID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSeqModelEmbeddings = "{" & vbNewLine & GetKVPairs("qrySeqModelEmbeddings", rs) & vbNewLine & "},"
    
End Function

''Command Name: Get Embedding Condition
Public Function GetEmbeddingCondition(frm As Object, Optional SeqModelEmbeddingID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelEmbeddingID) Then
        SeqModelEmbeddingID = frm("SeqModelEmbeddingID")
        If ExitIfTrue(isFalse(SeqModelEmbeddingID), "SeqModelEmbeddingID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelEmbeddings WHERE SeqModelEmbeddingID = " & SeqModelEmbeddingID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetEmbeddingCondition = GetReplacedTemplate(rs, "embedding condition")
    CopyToClipboard GetEmbeddingCondition
    
End Function
