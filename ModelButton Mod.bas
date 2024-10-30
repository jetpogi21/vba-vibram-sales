Attribute VB_Name = "ModelButton Mod"
Option Compare Database
Option Explicit

Public Function ModelButtonCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function ModelButton_ModelButton_AfterUpdate(frm As Form)
    
    Dim ModelButton: ModelButton = frm("ModelButton"): If ExitIfTrue(isFalse(ModelButton), "ModelButton is empty..") Then Exit Function
    
    If Not isFalse(frm("FunctionName")) Then Exit Function
    frm("FunctionName") = CreateFunctionName(ModelButton)
    
End Function

Public Function CreateFunctionName(Text)
    
    Dim FunctionName
    FunctionName = replace(Text, ".", "_")
    FunctionName = StrConv(FunctionName, vbProperCase)
    FunctionName = replace(FunctionName, " ", "")
    
    CreateFunctionName = FunctionName
    
End Function

Public Function OpenButtonModule(frm As Form)

    Dim FunctionName: FunctionName = frm("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    DoCmd.OpenModule , FunctionName
    
End Function

Public Function ModelButton_FunctionName_AfterUpdate(frm As Form)
    
    Dim FunctionName: FunctionName = frm("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    
    If Not isFalse(frm("ModelButton")) Then Exit Function
    
    frm("ModelButton") = GetButtonCaptionFromFunctionName(FunctionName)
    
End Function

Public Function GetButtonCaptionFromFunctionName(FunctionName) As String
    
    Dim separatedWords As New clsArray
    separatedWords.arr = SeparateWords(FunctionName)
    GetButtonCaptionFromFunctionName = StrConv(separatedWords.JoinArr(" "), vbProperCase)
    
End Function

Public Function ModelButtonCreateFunction(frm As Form)
    
    Dim ModelButtonID: ModelButtonID = frm("ModelButtonID"): If ExitIfTrue(isFalse(ModelButtonID), "ModelButtonID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryModelButtons WHERE ModelButtonID = " & ModelButtonID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    Dim ModelID: ModelID = rs.fields("ModelID"): If ExitIfTrue(isFalse(ModelID), "ModelID is empty..") Then Exit Function
    Dim TableName: TableName = rs.fields("TableName")
    If isFalse(TableName) Then TableName = "TableName"
    Dim Model: Model = rs.fields("Model"): If ExitIfTrue(isFalse(Model), "Model is empty..") Then Exit Function
    Dim templateName: templateName = rs.fields("TemplateName")
    Dim ModelButton: ModelButton = rs.fields("ModelButton"): If ExitIfTrue(isFalse(ModelButton), "ModelButton is empty..") Then Exit Function
    
    Dim PrimaryKey: PrimaryKey = GetPrimaryKeyFromTable(ModelID)
    
    Dim strFunction As String
    strFunction = "''Command Name: " & ModelButton & vbCrLf & _
                    "Public Function " & FunctionName & "(frm As Object, Optional " & PrimaryKey & " = """")" & vbCrLf & _
                    vbCrLf & _
                    "    RunCommandSaveRecord" & vbCrLf & _
                    vbCrLf & _
                    "    If isFalse(" & PrimaryKey & ") Then" & vbCrLf & _
                    "        " & PrimaryKey & " = frm(""" & PrimaryKey & """)" & vbCrLf & _
                    "        If ExitIfTrue(isFalse(" & PrimaryKey & "), """ & PrimaryKey & " is empty.."") Then Exit Function" & vbCrLf & _
                    "    End If" & vbCrLf & _
                    vbCrLf & _
                    "    Dim lines As New clsArray" & vbCrLf & _
                    "    Dim sqlStr: sqlStr = ""SELECT * FROM " & TableName & " WHERE " & PrimaryKey & " = "" & " & PrimaryKey & vbCrLf & _
                    "    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)" & vbCrLf & _
                    vbCrLf & _
                    "    " & FunctionName & " = GetReplacedTemplate(rs, """ & templateName & """)" & vbCrLf & _
                    "    " & FunctionName & " = GetGeneratedByFunctionSnippet(" & FunctionName & "," & Esc(FunctionName) & "," & Esc(templateName) & ")" & vbCrLf & _
                    "    CopyToClipboard " & FunctionName & vbCrLf & _
                    "End Function"
    
    Dim moduleName: moduleName = Model & " Mod"
    AddFunctionToModule moduleName, replace(strFunction, "tbl", "qry")
    
    OpenModule moduleName, FunctionName
    
End Function

Public Sub AddFunctionToModule(moduleName, strFunction)
    
    Dim code As String
    Dim lineNum As Long
    Dim modObject As Module
    
    ' Set the code for the new function
    code = strFunction
    
'    Dim item
'    For Each item In Application.Modules
'        Debug.Print item.name
'    Next item
    
    Set modObject = Application.Modules(moduleName)
    ' Find the last line of code in the module
    lineNum = modObject.CountOfLines

    ' Insert the new function code into the module
    modObject.InsertLines lineNum + 1, code
    
End Sub

''Command Name: Process File Path Template
Public Function ProcessFilePathTemplate(frm As Object, Optional ModelButtonID = "")

    RunCommandSaveRecord

    If isFalse(ModelButtonID) Then
        ModelButtonID = frm("ModelButtonID")
        If ExitIfTrue(isFalse(ModelButtonID), "ModelButtonID is empty..") Then Exit Function
    End If

    Dim FilePathTemplate As String: FilePathTemplate = frm("FilePathTemplate")
    If isFalse(FilePathTemplate) Then
        MsgBox "File Path Template is empty..", vbOKOnly
        Exit Function
    End If
    
    frm("FilePathTemplate") = GetFilePathTemplate(FilePathTemplate)
    
    
End Function

''Command Name: Process Export Const
Public Function ProcessExportConst(frm As Object, Optional ModelButtonID = "")

    RunCommandSaveRecord

    If isFalse(ModelButtonID) Then
        ModelButtonID = frm("ModelButtonID")
        If ExitIfTrue(isFalse(ModelButtonID), "ModelButtonID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblModelButtons WHERE NOT ExportConst IS NULL"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim ExportConst: ExportConst = rs.fields("ExportConst")
        DoCmd.OpenForm "frmCodeReplacers", , , "TemplateName = " & Esc("Export As Const")
        
        Dim frm2 As Form: Set frm2 = Forms("frmCodeReplacers")
        frm2("Snippet") = ExportConst
        TranslateCodeSnippet frm2
        Dim TemplateContent: TemplateContent = frm2("TranslatedSnippet")
        rs.Edit
        rs.fields("ExportConst") = TemplateContent
        rs.Update
        rs.MoveNext
    Loop
    
    MsgBox "Successfully updated."
    
    
End Function
