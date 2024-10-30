Attribute VB_Name = "FunctionChainItem Mod"
Option Compare Database
Option Explicit

Public Function FunctionChainItemCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm("ModelButtonID").OnDblClick = "=FunctionChainItem_ModelButtonID_OnDblClick([Form])"
            frm("ModelButtonID").DisplayAsHyperlink = 2
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function FunctionChainItem_ModelButtonID_OnDblClick(frm As Form)
    
    Dim ModelButtonID: ModelButtonID = frm("ModelButtonID"): If ExitIfTrue(isFalse(ModelButtonID), "ModelButtonID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblModelButtons WHERE ModelButtonID = " & ModelButtonID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    DoCmd.OpenModule , FunctionName
    
End Function


Public Function RunBackendProjectFunction(frm As Object, Optional FunctionChainItemID = "")

    RunCommandSaveRecord

    If isFalse(FunctionChainItemID) Then
        FunctionChainItemID = frm("FunctionChainItemID")
        If ExitIfTrue(isFalse(FunctionChainItemID), "FunctionChainItemID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryFunctionChainItems WHERE FunctionChainItemID = " & FunctionChainItemID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    
    Run FunctionName, frm, BackendProjectID
    
End Function

Public Function FCItem_RunModelFunction(frm As Object, Optional FunctionChainItemID = "")

    RunCommandSaveRecord

    If isFalse(FunctionChainItemID) Then
        FunctionChainItemID = frm("FunctionChainItemID")
        If ExitIfTrue(isFalse(FunctionChainItemID), "FunctionChainItemID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryFunctionChainItems WHERE FunctionChainItemID = " & FunctionChainItemID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    
    Run FunctionName, frm, SeqModelID
    
End Function

Public Function FCItem_RunOnMultipleModelFunction(frm As Object, Optional FunctionChainItemID = "")

    RunCommandSaveRecord

    If isFalse(FunctionChainItemID) Then
        FunctionChainItemID = frm("FunctionChainItemID")
        If ExitIfTrue(isFalse(FunctionChainItemID), "FunctionChainItemID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryFunctionChainItems WHERE FunctionChainItemID = " & FunctionChainItemID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SeqModelID
    Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    
    Dim SubSeqModel
    
    Dim resp: resp = MsgBox("WARNING, this will replace all the files existing on the folders." & vbCrLf & "Do you want to proceed?", vbCritical + vbYesNo)
    
    If resp = vbNo Then Exit Function
    
    NoHasWriteToFilePrompt = True
    
    Set rs = frm.parent.subSeqModelID.Form.RecordsetClone
    rs.MoveFirst
    
    Do Until rs.EOF
        If rs.fields("Selected") Then
            SeqModelID = rs.fields("Value")
            Run FunctionName, frm, SeqModelID
        End If
        rs.MoveNext
    Loop
    
    NoHasWriteToFilePrompt = False
    
    MsgBox "Task complete!"
    
End Function
