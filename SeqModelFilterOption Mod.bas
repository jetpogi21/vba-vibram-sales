Attribute VB_Name = "SeqModelFilterOption Mod"
Option Compare Database
Option Explicit

Public Function SeqModelFilterOptionCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm("FieldValue").AfterUpdate = "=SeqModelFilterOption_FieldValue_AfterUpdate([Form])"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SeqModelFilterOption_FieldValue_AfterUpdate(frm As Form)

    Dim fieldValue: fieldValue = frm("FieldValue")
    If isFalse(fieldValue) Then Exit Function
    
    frm("FieldCaption") = ConvertToVerboseCaption(fieldValue)
    
End Function


Public Function GetFilterManualOption(frm As Object, Optional SeqModelFilterOptionID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterOptionID) Then
        SeqModelFilterOptionID = frm("SeqModelFilterOptionID")
        If ExitIfTrue(isFalse(SeqModelFilterOptionID), "SeqModelFilterOptionID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilterOptions WHERE SeqModelFilterOptionID = " & SeqModelFilterOptionID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetFilterManualOption = GetReplacedTemplate(rs, "GetFilterManualOption")
    GetFilterManualOption = GetGeneratedByFunctionSnippet(GetFilterManualOption, "GetFilterManualOption", "GetFilterManualOption", , True)
    CopyToClipboard GetFilterManualOption
    
End Function

Public Function GetSeqModelFilterOptionKeys(frm As Object, Optional SeqModelFilterOptionID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterOptionID) Then
        SeqModelFilterOptionID = frm("SeqModelFilterOptionID")
        If ExitIfTrue(isFalse(SeqModelFilterOptionID), "SeqModelFilterOptionID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilterOptions WHERE SeqModelFilterOptionID = " & SeqModelFilterOptionID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSeqModelFilterOptionKeys = "{" & vbNewLine & GetKVPairs("qrySeqModelFilterOptions", rs) & vbNewLine & "},"

End Function
