Attribute VB_Name = "SeqModelHook Mod"
Option Compare Database
Option Explicit

Public Function SeqModelHookCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SeqModelHook_Caption_AfterUpdate(frm As Form)
    
    Dim Caption: Caption = frm("Caption"): If isFalse(Caption) Then Exit Function
    
    frm("Slug") = ConvertToModelPath(Caption, True)
End Function

Public Function WriteToHookPostRoute(frm As Object, Optional SeqModelHookID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelHookID) Then
        SeqModelHookID = frm("SeqModelHookID")
        If ExitIfTrue(isFalse(SeqModelHookID), "SeqModelHookID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelHooks WHERE SeqModelHookID = " & SeqModelHookID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim Slug: Slug = rs.fields("Slug"): If ExitIfTrue(isFalse(Slug), "Slug is empty..") Then Exit Function
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    
    WriteToHookPostRoute = GetReplacedTemplate(rs, "Create hook post route")
    WriteToHookPostRoute = GetGeneratedByFunctionSnippet(WriteToHookPostRoute, "WriteToHookPostRoute", "Create hook post route")
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim filePath: filePath = ProjectPath & "src\app\api\" & ModelPath & "\" & Slug & ".ts"
    WriteToFile filePath, WriteToHookPostRoute, SeqModelID

    CopyToClipboard WriteToHookPostRoute
    
End Function

