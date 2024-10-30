Attribute VB_Name = "LoopFunctionCreator Mod"
Option Compare Database
Option Explicit

Public Function LoopFunctionCreatorCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


Public Function LoopFunctionCreator_ModelID_OnChange(frm As Object, Optional parentMode As Boolean = True)
    
    Dim prefix: prefix = "Parent"
    If Not parentMode Then prefix = "Child"
    
    Dim ModelID: ModelID = frm(prefix & "ModelID")
    
    Dim PrimaryKey, TableName
    PrimaryKey = GetPrimaryKeyFromTable(ModelID)
    TableName = GetTableNameFromModelID(ModelID)
    
    frm(prefix & "PrimaryKey") = PrimaryKey
    frm(prefix & "Table") = TableName
    
    If parentMode Then
        LoopFunctionCreator_MainFunction_AfterUpdate frm
    End If
    
End Function

Public Function CreateALoopFunction(frm As Object, Optional LoopFunctionCreatorID = "")

    RunCommandSaveRecord

    If isFalse(LoopFunctionCreatorID) Then
        LoopFunctionCreatorID = frm("LoopFunctionCreatorID")
        If ExitIfTrue(isFalse(LoopFunctionCreatorID), "LoopFunctionCreatorID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryLoopFunctionCreators WHERE LoopFunctionCreatorID = " & LoopFunctionCreatorID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim Model: Model = rs.fields("Model"): If ExitIfTrue(isFalse(Model), "Model is empty..") Then Exit Function
    Dim ParentFunctionName: ParentFunctionName = rs.fields("ParentFunctionName"): If ExitIfTrue(isFalse(ParentFunctionName), "ParentFunctionName is empty..") Then Exit Function
    Dim Joiner: Joiner = rs.fields("Joiner"): If ExitIfTrue(isFalse(Joiner), "Joiner is empty..") Then Exit Function
    Dim ParentModelID: ParentModelID = rs.fields("ParentModelID"): If ExitIfTrue(isFalse(ParentModelID), "ParentModelID is empty..") Then Exit Function
    
    CreateALoopFunction = GetReplacedTemplate(rs, "LoopFunctionCreator")
    If Joiner = "NewLine" Then
        CreateALoopFunction = replace(CreateALoopFunction, Esc("NewLine"), "vbNewLine")
    ElseIf Joiner = "Semicolon" Then
        CreateALoopFunction = replace(CreateALoopFunction, Esc("Semicolon"), Esc(";"))
    ElseIf Joiner = "Period" Then
        CreateALoopFunction = replace(CreateALoopFunction, Esc("Period"), Esc("."))
    ElseIf Joiner = "Comma" Then
        CreateALoopFunction = replace(CreateALoopFunction, Esc("Comma"), Esc(","))
    ElseIf Joiner = "Space" Then
        CreateALoopFunction = replace(CreateALoopFunction, Esc("Space"), Esc(" "))
    End If
    
    RunSQL "DELETE FROM tblModelButtons WHERE FunctionName = " & Esc(ParentFunctionName)
    
    lines.Add ParentModelID
    lines.Add Esc(GetButtonCaptionFromFunctionName(ParentFunctionName))
    lines.Add Esc(ParentFunctionName)
    lines.Add Emax("tblModelButtons", "ModelID = " & ParentModelID, "ModelButtonOrder") + 1
    
    RunSQL "INSERT INTO tblModelButtons (ModelID, ModelButton, FunctionName, ModelButtonOrder) " & _
        " VALUES (" & lines.JoinArr & ")"
    
    Dim moduleName: moduleName = Model & " Mod"
    AddFunctionToModule moduleName, CreateALoopFunction
    
    Dim FunctionName: FunctionName = ParentFunctionName
    OpenModule moduleName, FunctionName
    
End Function

Public Function frmLoopFunctionCreators_TemplateName_OnDblClick(frm As Form)
    Dim templateName As String
    Dim Note
    Dim whereClause As String
    Dim rs As Recordset
    
    templateName = frm("TemplateName")
    Note = frm("Note")
    whereClause = "SnippetDescription = " & Esc("Template - " & templateName) & " OR SnippetDescription Like " & Esc("*" & templateName & "*")
    
    Set rs = ReturnRecordset("SELECT * FROM tblSnippets WHERE " & whereClause)
    
    If rs.EOF Or isFalse(templateName) Then
        DoCmd.OpenForm "frmSnippets", , , , acFormAdd
        Forms("frmSnippets")("SnippetDescription") = "Template - " & templateName
        Forms("frmSnippets")("Snippet") = Note
        RunSQL "UPDATE tblSnippetCategoryID SET Selected = -1 WHERE FilterLabel = " & Esc("Template")
    Else
        DoCmd.OpenForm "frmSnippets", , , whereClause
    End If
    
End Function


Public Function LoopFunctionCreator_CreateChildFunction(frm As Object, Optional LoopFunctionCreatorID = "", Optional notify = True)

    RunCommandSaveRecord

    If isFalse(LoopFunctionCreatorID) Then
        LoopFunctionCreatorID = frm("LoopFunctionCreatorID")
        If ExitIfTrue(isFalse(LoopFunctionCreatorID), "LoopFunctionCreatorID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryLoopFunctionCreators WHERE LoopFunctionCreatorID = " & LoopFunctionCreatorID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ChildFunctionName: ChildFunctionName = rs.fields("ChildFunctionName"): If ExitIfTrue(isFalse(ChildFunctionName), "ChildFunctionName is empty..") Then Exit Function
    Dim ChildModelID: ChildModelID = rs.fields("ChildModelID"): If ExitIfTrue(isFalse(ChildModelID), "ChildModelID is empty..") Then Exit Function
    Dim templateName: templateName = rs.fields("TemplateName"): If ExitIfTrue(isFalse(templateName), "TemplateName is empty..") Then Exit Function
    Dim Note: Note = rs.fields("Note")
    
    RunSQL "DELETE FROM tblModelButtons WHERE FunctionName = " & Esc(ChildFunctionName)
    
    lines.Add ChildModelID
    lines.Add Esc(GetButtonCaptionFromFunctionName(ChildFunctionName))
    lines.Add Esc(ChildFunctionName)
    lines.Add Emax("tblModelButtons", "ModelID = " & ChildModelID, "ModelButtonOrder") + 1
    lines.Add Esc(templateName)
    
    If Not isFalse(Note) Then
        lines.Add Esc(replace(Note, """", """"""))
    Else
        lines.Add "null"
    End If
    
    RunSQL "INSERT INTO tblModelButtons (ModelID,ModelButton, FunctionName, ModelButtonOrder, TemplateName, [Note]) " & _
        " VALUES (" & lines.JoinArr & ")"
        
    DoCmd.OpenForm "frmModelButtons", , , "FunctionName = " & Esc(ChildFunctionName)
    Set frm = Forms("frmModelButtons")
     
    ModelButtonCreateFunction frm
     
    DoCmd.Close acForm, frm.Name, acSaveNo
     
    If notify Then MsgBox "Function Created"
    
End Function

Public Function LoopFunctionCreator_GetMainTemplate(frm As Form)

    Dim MainFunction: MainFunction = frm("MainFunction"): If isFalse(MainFunction) Then Exit Function
    
    Dim MainTemplate: MainTemplate = ELookup("tblModelButtons", "FunctionName = " & Esc(MainFunction), "TemplateName")
    
    frm("MainTemplate") = MainTemplate
    
End Function

Public Function LoopFunctionCreator_MainFunction_AfterUpdate(frm As Form)

    Dim MainFunction: MainFunction = frm("MainFunction"): If isFalse(MainFunction) Then Exit Function
    Dim ParentFunctionName: ParentFunctionName = frm("ParentFunctionName"): If isFalse(ParentFunctionName) Then Exit Function
    Dim ParentPrimaryKey: ParentPrimaryKey = frm("ParentPrimaryKey"): If isFalse(ParentFunctionName) Then Exit Function
    
    frm("ReplaceSnippet") = MainFunction & " = Replace(" & MainFunction & ",""[" & ParentFunctionName & "]""," & ParentFunctionName & "(frm," & ParentPrimaryKey & "))"
   
    LoopFunctionCreator_GetMainTemplate frm
End Function
