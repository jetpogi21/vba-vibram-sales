Attribute VB_Name = "GetModelOptionWorkflow Mod"
Option Compare Database
Option Explicit

Public Function GetModelOptionWorkflowCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            ''[ClientPath]src\lib\[ModelPath]\customAddToRequiredList.tsx
            ''import { customAddToRequiredList as customAddToRequiredList[ModelName] } from "@/lib/[ModelPath]/customAddToRequiredList";
            ''getAllCustomAddToRequiredList
            ''customAddToRequiredList[ModelName]
            ''Record<string, useAddToRequiredListType>;
            frm("FilePathTemplate").Format = "@;" & Esc("[ClientPath]src\lib\[ModelPath]\customAddToRequiredList.tsx")
            frm("ModelOptionImportStatement").Format = "@;" & Esc("import { customAddToRequiredList as customAddToRequiredList[ModelName] } from ""@/lib/[ModelPath]/customAddToRequiredList""")
            frm("ModelOptionImportAs").Format = "@;" & Esc("getAllCustomAddToRequiredList")
            frm("ExportConst").Format = "@;" & Esc("customAddToRequiredList")
            frm("TypescriptInference").Format = "@;" & Esc("Record<string, useAddToRequiredListType>")
            
            frm("TemplateName").AfterUpdate = "=frmGetModelOptionWorkflows_TemplateName_AfterUpdate([Form])"
            
            frm("FunctionName").OnDblClick = "=OpenModule(Null,[FunctionName])"
            frm("FunctionName").DisplayAsHyperlink = 2
'            Dim item As control
'            For Each item In frm.controls
'                If item.ControlType = acCommandButton Then
'                    item.Height = item.Height * 2
'                End If
'            Next item
            
            
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function Open_frmSnippets_by_TemplateName(templateName As String)

    templateName = "Template - " & templateName
    DoCmd.OpenForm "frmSnippets", , , "SnippetDescription = " & Esc(templateName)
    
End Function

Public Function GetModelOptionWorkflow_ModeloptionClipboard(frm As Form)
    
    DoCmd.RunCommand acCmdSaveRecord
    
    Dim GetModelOptionWorkflowID: GetModelOptionWorkflowID = frm("GetModelOptionWorkflowID")
    Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim ImportLine: ImportLine = frm("ImportLine"): If ExitIfTrue(isFalse(ImportLine), "ImportLine is empty..") Then Exit Function
    Dim templateName: templateName = frm("TemplateName"): If ExitIfTrue(isFalse(templateName), "TemplateName is empty..") Then Exit Function
    
    Dim rs As Recordset: Set rs = frm.RecordsetClone
    
    Dim TemplateContent: TemplateContent = GetReplacedTemplate(rs, "GetModelOptionWorkflow")
    
    CopyToClipboard TemplateContent
    MsgBox "Paste the clipboard content within this function"
    
    DoCmd.OpenModule , "WriteToGetmodeloptions_tsx"
    
End Function

Public Function GetModelOptionWorkflow_RunActualReplacement(frm As Form)
    
    DoCmd.RunCommand acCmdSaveRecord
    
    Dim GetModelOptionWorkflowID: GetModelOptionWorkflowID = frm("GetModelOptionWorkflowID")
    Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim ImportLine: ImportLine = frm("ImportLine"): If ExitIfTrue(isFalse(ImportLine), "ImportLine is empty..") Then Exit Function
    Dim templateName: templateName = frm("TemplateName"): If ExitIfTrue(isFalse(templateName), "TemplateName is empty..") Then Exit Function
    
    Dim BackendProjectID: BackendProjectID = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "BackendProjectID")
    
    WriteToGetmodeloptions_tsx frm, BackendProjectID
    
End Function

Public Function GetModelOptionWorkflow_ConvertLineItem(frm As Form)
    
    Dim GetModelOptionWorkflowID: GetModelOptionWorkflowID = frm("GetModelOptionWorkflowID")
    Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim LineItem: LineItem = frm("LineItem"): If ExitIfTrue(isFalse(LineItem), "LineItem is empty..") Then Exit Function
    Dim templateName: templateName = frm("TemplateName"): If ExitIfTrue(isFalse(templateName), "TemplateName is empty..") Then Exit Function
    
    Dim ModelName: ModelName = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "ModelName")
    
    DoCmd.OpenForm "frmCodeReplacers", , , "CodeReplacerID = 61"
    
    Set frm = Forms("frmCodeReplacers")
    frm("SQLQuery") = "SELECT * FROM qrySeqModels WHERE ModelName = " & Esc(ModelName) & " ORDER BY SeqModelID DESC"
    frm("TemplateName") = templateName
    frm("Snippet") = LineItem
    
    TranslateCodeSnippet frm
    HookToSnippets frm
    
End Function

Public Function GetModelOptionWorkflow_ConvertTemplateContent(frm As Form)
    
    Dim GetModelOptionWorkflowID: GetModelOptionWorkflowID = frm("GetModelOptionWorkflowID")
    Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim templateName: templateName = frm("TemplateName"): If ExitIfTrue(isFalse(templateName), "TemplateName is empty..") Then Exit Function
    Dim TemplateContent: TemplateContent = frm("TemplateContent"): If ExitIfTrue(isFalse(TemplateContent), "TemplateContent is empty..") Then Exit Function
    
    Dim ModelName: ModelName = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "ModelName")
    
    DoCmd.OpenForm "frmCodeReplacers", , , "CodeReplacerID = 61"
    
    Set frm = Forms("frmCodeReplacers")
    frm("SQLQuery") = "SELECT * FROM qrySeqModels WHERE ModelName = " & Esc(ModelName) & " ORDER BY SeqModelID DESC"
    frm("TemplateName") = templateName
    frm("Snippet") = TemplateContent
    
    TranslateCodeSnippet frm
    HookToSnippets frm
    
End Function

Public Function GetModelOptionWorkflow_ConvertModelOptionTexts(frm As Form)
    
    Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    ''FilePathTemplate
    ''ModelOptionImportStatement
    ''ModelOptionImportAs
    ''ExportConst
    ''TypescriptInference
    
    Dim toBeConverteds As New clsArray: toBeConverteds.arr = "FilePathTemplate,ModelOptionImportStatement,ModelOptionImportAs," & _
        "ExportConst,TypescriptInference"
    
    DoCmd.OpenForm "frmCodeReplacers", , , "CodeReplacerID = 3"
    Dim frm2 As Form: Set frm2 = Forms("frmCodeReplacers")
    frm2("SQLQuery") = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelID DESC"
    
    Dim item, items As New clsArray
    For Each item In toBeConverteds.arr
        frm2("Snippet") = frm(item)
        TranslateCodeSnippet frm2
        frm(item) = frm2("TranslatedSnippet")
    Next item
    
    DoCmd.Close acForm, "frmCodeReplacers", acSaveNo
    
End Function

Public Function GetModelOptionWorkflow_ConvertImportLine(frm As Form)

    Dim GetModelOptionWorkflowID: GetModelOptionWorkflowID = frm("GetModelOptionWorkflowID")
    Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim ImportLine: ImportLine = frm("ImportLine"): If ExitIfTrue(isFalse(ImportLine), "ImportLine is empty..") Then Exit Function
    Dim templateName: templateName = frm("TemplateName"): If ExitIfTrue(isFalse(templateName), "TemplateName is empty..") Then Exit Function
    
    Dim ModelName: ModelName = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "ModelName")
    
    DoCmd.OpenForm "frmCodeReplacers", , , "CodeReplacerID = 61"
    
    Set frm = Forms("frmCodeReplacers")
    frm("SQLQuery") = "SELECT * FROM qrySeqModels WHERE ModelName = " & Esc(ModelName) & " ORDER BY SeqModelID DESC"
    frm("Snippet") = ImportLine
    
    TranslateCodeSnippet frm
    
    Open_frmSnippets_by_TemplateName "GetComponentImport"
    
    CopyFieldContent frm, "TranslatedSnippet"
    
    MsgBox "Paste the clipboard content at the end of this snippet"
    
End Function

Public Function frmGetModelOptionWorkflows_TemplateName_AfterUpdate(frm As Form)
    
    Dim templateName: templateName = frm("templateName")
    frm("FunctionCaption") = "Write to " & templateName
    frm("FunctionName") = CreateFunctionName(frm("FunctionCaption"))
    
End Function

''Command Name: Create WriteToFile Function
Public Function CreateWritetofileFunction(frm As Object, Optional GetModelOptionWorkflowID = "")

    RunCommandSaveRecord

    If isFalse(GetModelOptionWorkflowID) Then
        GetModelOptionWorkflowID = frm("GetModelOptionWorkflowID")
        If ExitIfTrue(isFalse(GetModelOptionWorkflowID), "GetModelOptionWorkflowID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblGetModelOptionWorkflows WHERE GetModelOptionWorkflowID = " & GetModelOptionWorkflowID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    Dim FilePathTemplate: FilePathTemplate = rs.fields("FilePathTemplate"): If ExitIfTrue(isFalse(FilePathTemplate), "FilePathTemplate is empty..") Then Exit Function
    FilePathTemplate = ConvertFilePathTemplate(FilePathTemplate)
    CreateWritetofileFunction = GetReplacedTemplate(rs, "WriteToFile Function", "FilePathTemplate")
    CopyToClipboard CreateWritetofileFunction
    
    Dim moduleName: moduleName = "SeqModel Mod"
    AddFunctionToModule moduleName, replace(CreateWritetofileFunction, "[FilePathTemplate]", FilePathTemplate)
    
    OpenModule moduleName, FunctionName
    
End Function


Public Function ConvertFilePathTemplate(FilePathTemplate) As String
    
    ''[ClientPath]src\lib\[ModelPath]\getModelProps.tsx from
    ''ClientPath & "src\lib\" & ModelPath & "\getModelProps.tsx"
    
    ''Remove the first character
    FilePathTemplate = Right(FilePathTemplate, Len(FilePathTemplate) - 1)
    
    ''Replace all the closing brackets with " & ""
    FilePathTemplate = replace(FilePathTemplate, "]", " & """)
    
    ''Replace all the remaining open brackets with " & ""
    FilePathTemplate = replace(FilePathTemplate, "[", """ & ")
    
    
    ConvertFilePathTemplate = FilePathTemplate
    
End Function
''Command Name: Hook To ModelButtons
Public Function HookToModelbuttons(frm As Object, Optional GetModelOptionWorkflowID = "")

    RunCommandSaveRecord

    If isFalse(GetModelOptionWorkflowID) Then
        GetModelOptionWorkflowID = frm("GetModelOptionWorkflowID")
        If ExitIfTrue(isFalse(GetModelOptionWorkflowID), "GetModelOptionWorkflowID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblGetModelOptionWorkflows WHERE GetModelOptionWorkflowID = " & GetModelOptionWorkflowID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName) is empty..") Then Exit Function
    Dim FunctionCaption: FunctionCaption = rs.fields("FunctionCaption"): If ExitIfTrue(isFalse(FunctionCaption), "FunctionCaption) is empty..") Then Exit Function
    Dim templateName: templateName = rs.fields("TemplateName"): If ExitIfTrue(isFalse(templateName), "TemplateName) is empty..") Then Exit Function
    Dim FilePathTemplate: FilePathTemplate = rs.fields("FilePathTemplate"): If ExitIfTrue(isFalse(FilePathTemplate), "FilePathTemplate) is empty..") Then Exit Function
    Dim ModelOptionImportStatement: ModelOptionImportStatement = rs.fields("ModelOptionImportStatement"): If ExitIfTrue(isFalse(ModelOptionImportStatement), "ModelOptionImportStatement) is empty..") Then Exit Function
    Dim ModelOptionImportAs: ModelOptionImportAs = rs.fields("ModelOptionImportAs"): If ExitIfTrue(isFalse(ModelOptionImportAs), "ModelOptionImportAs) is empty..") Then Exit Function
    Dim ExportConst: ExportConst = rs.fields("ExportConst"): If ExitIfTrue(isFalse(ExportConst), "ExportConst) is empty..") Then Exit Function
    Dim TypescriptInference: TypescriptInference = rs.fields("TypescriptInference"): If ExitIfTrue(isFalse(TypescriptInference), "TypescriptInference) is empty..") Then Exit Function
    
    Dim fields As New clsArray: fields.arr = "ModelID,ModelButton,FunctionName,ModelButtonOrder,TemplateName," & _
        "FilePathTemplate,ModelOptionImportStatement,ModelOptionImportAs,ExportConst,TypescriptInference"
    Dim fieldValues As New clsArray
    Set fieldValues = New clsArray
    
    Dim ModelID: ModelID = ELookup("tblModels", "Model = " & Esc("SeqModel"), "ModelID")
    
    fieldValues.Add ModelID
    fieldValues.Add FunctionCaption
    fieldValues.Add FunctionName
    
    Dim ModelButtonOrder: ModelButtonOrder = ELookup("tblModelButtons", "ModelButtonID > 0", "ModelButtonOrder", "ModelButtonOrder DESC")
    fieldValues.Add ModelButtonOrder
    
    fieldValues.Add templateName
    fieldValues.Add FilePathTemplate
    fieldValues.Add ModelOptionImportStatement
    fieldValues.Add ModelOptionImportAs
    fieldValues.Add ExportConst
    fieldValues.Add TypescriptInference
    
    UpsertRecord "tblModelButtons", fields, fieldValues, "FunctionName = " & Esc(FunctionName)

    ''ModelButton --> ModelButton
    ''FunctionName --> ModelButton
    ''ModelButtonOrder --> ModelButton
    ''TemplateName --> ModelButton
    ''FilePathTemplate --> ModelButton
    ''ModelOptionImportStatement --> ModelButton
    ''ModelOptionImportAs --> ModelButton
    ''ExportConst --> ModelButton
    ''TypescriptInference --> ModelButton
    
End Function

''Command Name: Hook To Function Chain Items
Public Function HookToFunctionChainItems(frm As Object, Optional GetModelOptionWorkflowID = "")

    RunCommandSaveRecord

    If isFalse(GetModelOptionWorkflowID) Then
        GetModelOptionWorkflowID = frm("GetModelOptionWorkflowID")
        If ExitIfTrue(isFalse(GetModelOptionWorkflowID), "GetModelOptionWorkflowID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblGetModelOptionWorkflows WHERE GetModelOptionWorkflowID = " & GetModelOptionWorkflowID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName) is empty..") Then Exit Function
    
    Dim ModelButtonID: ModelButtonID = ELookup("qryModelButtons", "FunctionName = " & Esc(FunctionName) & _
        " AND Model = ""SeqModel""", "ModelButtonID")
    
    ''Hardcode: FunctionChainID --> 28
    Dim FunctionChainID: FunctionChainID = ELookup("tblFunctionChains", "FunctionChainName = " & Esc("Model Related Functions"), "FunctionChainID")
    If isPresent("tblFunctionChainItems", "ModelButtonID = " & ModelButtonID & " AND FunctionChainID = " & FunctionChainID) Then Exit Function
    ''Get the ModelButtonID using the FunctionName
    ''FunctionChainID --> 28
    ''SuspendInBatch --> True
    ''Special File --> Note
    ''Get the max FunctionOrder where note is special file add + .01
    ''Insert only whenever there's no existing ModelButtonID
    
    Dim fields As New clsArray: fields.arr = "FunctionChainID,ModelButtonID,FunctionOrder,SuspendInBatch,Note"
    Dim fieldValues As New clsArray
    
    fieldValues.Add FunctionChainID
    fieldValues.Add ModelButtonID
    
    Dim FunctionOrder: FunctionOrder = ELookup("tblFunctionChainItems", "FunctionChainID = " & FunctionChainID & _
        " AND Note Like " & Esc("*Special File*"), "FunctionOrder", "FunctionOrder DESC")
    fieldValues.Add Coalesce(FunctionOrder, 0) + 0.01
    fieldValues.Add True
    fieldValues.Add "Special File"
    
    UpsertRecord "tblFunctionChainItems", fields, fieldValues
    
    
End Function
''Command Name: Run Complete Workflow
Public Function GetModelOptionWorkflow_RunCompleteWorkflow(frm As Object, Optional GetModelOptionWorkflowID = "")

    RunCommandSaveRecord

    If isFalse(GetModelOptionWorkflowID) Then
        GetModelOptionWorkflowID = frm("GetModelOptionWorkflowID")
        If ExitIfTrue(isFalse(GetModelOptionWorkflowID), "GetModelOptionWorkflowID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryGetModelOptionWorkflows WHERE GetModelOptionWorkflowID = " & GetModelOptionWorkflowID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelOptionWorkflow_ConvertModelOptionTexts frm
    RunCommandSaveRecord
    HookToModelbuttons frm
    HookToFunctionChainItems frm
    
End Function
