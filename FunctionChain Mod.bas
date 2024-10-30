Attribute VB_Name = "FunctionChain Mod"
Option Compare Database
Option Explicit

Public Function FunctionChainCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        
            frm.OnCurrent = "=frmFunctionChains_OnCurrent([Form])"
            frm("BackendProjectID").AfterUpdate = "=frmFunctionChains_BackendProjectID_AfterUpdate([Form])"
            CreateButtonControl frm, "Open Project", "cmdOpenProject", "=OpenFormFromRecord([Form],""BackendProjectID"",""frmBackendProjects"")"
            With frm("cmdOpenProject")
                .Height = frm("BackendProjectID").Height
                .Width = frm("BackendProjectID").Width * 1 / 2
                .Top = frm("BackendProjectID").Top - .Height
                .Left = GetRight(frm("BackendProjectID")) - .Width
            End With
            
            CreateButtonControl frm, "Open Model", "cmdOpenModel", "=OpenFormFromRecord([Form],""SeqModelID"",""frmSeqModels"")"
            With frm("cmdOpenModel")
                .Height = frm("BackendProjectID").Height
                .Width = frm("BackendProjectID").Width * 1 / 2
                .Top = frm("SeqModelID").Top - .Height
                .Left = GetRight(frm("SeqModelID")) - .Width
            End With
            
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SetProjectToCurrent()

    If IsFormOpen("frmBackendProjects") Then
        Dim frm: Set frm = Forms("frmBackendProjects")
        Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
        
        RunSQL "UPDATE tblFunctionChains SET BackendProjectID = " & BackendProjectID
    End If
    
End Function

Private Function LoadSeqModels(frm As Form)

    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID"): If isFalse(BackendProjectID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT SeqModelID, ModelName FROM tblseqmodels WHERE BackendProjectID = " & BackendProjectID & " ORDER BY ModelName"
    
    frm("SeqModelID").RowSource = sqlStr
    
End Function

Public Function UpdateSubSeqModelID_Recordsource(frm As Form)

    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID"): If isFalse(BackendProjectID) Then Exit Function
    Dim sqlStr As String: sqlStr = "SELECT SeqModelID As [Value], ModelName AS Label From tblSeqModels WHERE BackendProjectID = " & BackendProjectID & " ORDER BY ModelName"
    
    Set frm = frm("SubSeqModelID").Form
    
    FilterContFormOnLoad frm, sqlStr, "tblFunctionChainSeqModelID"
    
End Function

Public Function frmFunctionChains_OnCurrent(frm As Form)

    LoadSeqModels frm
    UpdateSubSeqModelID_Recordsource frm
    
End Function

Public Function frmFunctionChains_BackendProjectID_AfterUpdate(frm As Form)
    
    LoadSeqModels frm
    UpdateSubSeqModelID_Recordsource frm
    
End Function

Public Function RunAllBackendProjectFunction(frm As Object, Optional FunctionChainID = "")

    RunCommandSaveRecord
    
    Dim resp: resp = MsgBox("WARNING, this will replace all the files existing on the folders." & vbCrLf & "Do you want to proceed?", vbCritical + vbYesNo)
    
    If resp = vbNo Then Exit Function

    If isFalse(FunctionChainID) Then
        FunctionChainID = frm("FunctionChainID")
        If ExitIfTrue(isFalse(FunctionChainID), "FunctionChainID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblFunctionChains WHERE FunctionChainID = " & FunctionChainID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        
    NoHasWriteToFilePrompt = True
    
    sqlStr = "SELECT FunctionChainItemID FROM tblFunctionChainItems WHERE FunctionChainID = " & FunctionChainID & " AND Not SuspendInBatch ORDER BY FunctionOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim FunctionChainItemID: FunctionChainItemID = rs.fields("FunctionChainItemID")
        RunBackendProjectFunction frm, FunctionChainItemID
        rs.MoveNext
    Loop
    
    NoHasWriteToFilePrompt = False
    
    MsgBox "Task complete!"

End Function

Public Function RunAll_FCItem_RunModelFunction(frm As Object, Optional FunctionChainID = "")

    RunCommandSaveRecord
    
    Dim resp: resp = MsgBox("WARNING, this will replace all the files existing on the folders." & vbCrLf & "Do you want to proceed?", vbCritical + vbYesNo)
    
    If resp = vbNo Then Exit Function

    If isFalse(FunctionChainID) Then
        FunctionChainID = frm("FunctionChainID")
        If ExitIfTrue(isFalse(FunctionChainID), "FunctionChainID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblFunctionChains WHERE FunctionChainID = " & FunctionChainID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    NoHasWriteToFilePrompt = True
    
    sqlStr = "SELECT FunctionChainItemID FROM tblFunctionChainItems WHERE FunctionChainID = " & FunctionChainID & " AND Not SuspendInBatch ORDER BY FunctionOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim FunctionChainItemID: FunctionChainItemID = rs.fields("FunctionChainItemID")
        FCItem_RunModelFunction frm, FunctionChainItemID
        rs.MoveNext
    Loop
    
    NoHasWriteToFilePrompt = False
    
    MsgBox "Task complete!"
    
End Function

Public Function RunAllFunctionsEachModels(frm As Object, Optional FunctionChainID = "")

    RunCommandSaveRecord
    
    Dim resp: resp = MsgBox("WARNING, this will replace all the files existing on the folders." & vbCrLf & "Do you want to proceed?", vbCritical + vbYesNo)
    
    If resp = vbNo Then Exit Function

    If isFalse(FunctionChainID) Then
        FunctionChainID = frm("FunctionChainID")
        If ExitIfTrue(isFalse(FunctionChainID), "FunctionChainID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblFunctionChains WHERE FunctionChainID = " & FunctionChainID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    NoHasWriteToFilePrompt = True
    
    Dim SeqModelID, seqModelIDs As New clsArray:  seqModelIDs.arr = Elookups("tblSeqModels", "BackendProjectID = " & BackendProjectID, "SeqModelID")
    
    sqlStr = "SELECT FunctionChainItemID, FunctionName FROM qryFunctionChainItems WHERE FunctionChainID = " & FunctionChainID & " AND Not SuspendInBatch ORDER BY FunctionOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
    
        Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
        For Each SeqModelID In seqModelIDs.arr
            Run FunctionName, frm, SeqModelID
        Next SeqModelID
        
        rs.MoveNext
    Loop
    
    NoHasWriteToFilePrompt = False
    
    MsgBox "Task complete!"
    
End Function
