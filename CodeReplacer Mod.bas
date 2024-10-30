Attribute VB_Name = "CodeReplacer Mod"
Option Compare Database
Option Explicit

Public Function CodeReplacerCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        
            Dim ctl As control
            For Each ctl In frm.controls
                If ctl.ControlType = acCommandButton And ctl.Name <> "cmdSourceSnippet" Then
                    ctl.Height = ctl.Height * 1.3
                End If
            Next ctl
            
            Dim item, buttons As New clsArray: buttons.arr = "cmdClearSnippet,cmdCopySnippet,cmdClearTranslatedSnippet,cmdCopyTranslatedSnippet"
            For Each item In buttons.arr
            
                ''CopyProperties frm, item, "cmdSmallButton"
                Dim clearFlag As Boolean: clearFlag = InStr(item, "Clear") > 0
                Dim ButtonCaption: ButtonCaption = IIf(clearFlag, "Clear", "Copy")
                ''frm(item).Caption = ButtonCaption
                Dim fieldName As String: fieldName = IIf(InStr(item, "Translated") > 0, "TranslatedSnippet", "Snippet")
                ''frm(item).OnClick = GetOnClick(item, fieldName, clearFlag)
                
                CreateButtonControl frm, ButtonCaption, item, GetOnClick(item, fieldName, clearFlag), "cmdSmallButton"
                
                frm(item).Width = frm("Snippet").Width / 10
                frm(item).Top = GetBottom(frm("Snippet"))
            Next item
            
            frm("Snippet").FontName = "Lucida Console"
            frm("Snippet").fontSize = 7
            frm("TranslatedSnippet").FontName = "Lucida Console"
            frm("TranslatedSnippet").fontSize = 7
            
            frm("SQLQuery").AfterUpdate = "=CodeReplaceSQLQueryAfterUpdate([Form])"
            frm("cmdSourceSnippet").OnClick = "=CodeReplacer_cmdSourceSnippetOnClick([Form])"
            
            frm("cmdCopySnippet").Left = GetRight(frm("Snippet")) - frm("cmdCopySnippet").Width
            frm("cmdClearSnippet").Left = frm("cmdCopySnippet").Left - frm("cmdClearSnippet").Width
            
            frm("cmdCopyTranslatedSnippet").Left = GetRight(frm("TranslatedSnippet")) - frm("cmdCopyTranslatedSnippet").Width
            frm("cmdClearTranslatedSnippet").Left = frm("cmdCopyTranslatedSnippet").Left - frm("cmdClearTranslatedSnippet").Width
            
    
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Private Function GetOnClick(ButtonName, fieldName, clearFlag As Boolean) As String
    If clearFlag Then
        GetOnClick = "=ClearField([Form]," & Esc(fieldName) & ")"
    Else
        GetOnClick = "=CopyFieldContent([Form]," & Esc(fieldName) & ")"
    End If
End Function

Public Function ClearField(frm As Object, fieldName)
    
    frm(fieldName) = Null
    
End Function

Public Function CopyFieldContent(frm As Object, fieldName)
    
    If Not isFalse(fieldName) Then
        CopyToClipboard frm(fieldName)
    End If
    
End Function

Public Function TranslateCodeSnippet(frm As Form)
    
    RunCommandSaveRecord
    ''TABLE: tblCodeReplacers Fields: CodeReplacerID|SQLQuery|ExcludedFieldNames|Snippet|TranslatedSnippet
    ''Timestamp|CreatedBy|RecordImportID|PriorityFields
    Dim sqlQuery: sqlQuery = frm("SQLQuery"): If ExitIfTrue(isFalse(sqlQuery), "SQLQuery is empty..") Then Exit Function
    Dim Snippet: Snippet = frm("Snippet"): If ExitIfTrue(isFalse(Snippet), "Snippet is empty..") Then Exit Function
    
    Dim ExcludedFieldNames: ExcludedFieldNames = frm("ExcludedFieldNames")
    Dim PriorityFields As New clsArray:
    If Not isFalse(frm("PriorityFields")) Then PriorityFields.arr = frm("PriorityFields")
    
    Dim ExcludedFields As New clsArray
    If Not isFalse(ExcludedFieldNames) Then ExcludedFields.arr = ExcludedFieldNames
    
    Dim fld As field, item
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlQuery)
    
    If PriorityFields.count > 0 Then
        For Each item In PriorityFields.arr
            If Not IsNull(rs.fields(item)) Then
                Snippet = replace(Snippet, rs.fields(item), "[" & item & "]", , , vbBinaryCompare)
            End If
        Next item
    End If
    
    For Each fld In rs.fields
        If Not ExcludedFields.InArray(fld.Name) Then
            If Not IsNull(rs.fields(fld.Name)) And fld.Type = 10 Then
                Snippet = replace(Snippet, rs.fields(fld.Name), "[" & fld.Name & "]", , , vbBinaryCompare)
            End If
        End If
    Next fld
    
    frm("TranslatedSnippet") = Snippet
    CopyToClipboard Snippet
    
End Function

Public Function CodeReplacer_cmdSourceSnippetOnClick(frm As Form)
    
    RunCommandSaveRecord
    Dim filePath: filePath = PromptFile: If ExitIfTrue(isFalse(filePath), "filePath is empty..") Then Exit Function
    Dim fileNumber, fileLine, fileContent
    fileNumber = FreeFile()
    Open filePath For Input As fileNumber
    
    Do Until EOF(fileNumber)
        Line Input #fileNumber, fileLine
        fileContent = fileContent & fileLine & vbCrLf
    Loop
    
    Close fileNumber
    
    frm("Snippet") = fileContent
    frm("SourceSnippet") = filePath
    
End Function

Public Function CodeReplaceSQLQueryAfterUpdate(frm As Form)
    
    Dim sqlQuery: sqlQuery = frm("SQLQuery"): If ExitIfTrue(isFalse(sqlQuery), "SQLQuery is empty..") Then Exit Function
    
    If isFalse(frm("PriorityFields")) Then
        If InStr(1, sqlQuery, "qrySeqModelFields") > 0 Then
            frm("PriorityFields") = "PluralizedModelName"
        ElseIf InStr(1, sqlQuery, "tblSeqModels") > 0 Then
            frm("PriorityFields") = "PluralizedModelName,VariablePluralName,PluralizedVerboseModelName"
        End If
    End If
    
End Function

Public Function HookToSnippets(frm As Form)
    
    ''Use 26 as the ID for CategoryID
    ''TABLE: tblSnippetCategories Fields: SnippetCategoryID|SnippetID|CategoryID|Timestamp|CreatedBy|RecordImportID
    ''TABLE: tblSnippets Fields: SnippetID|SnippetDescription|Snippet|SnippetNote|Timestamp|CreatedBy|RecordImportID
    Dim templateName: templateName = frm("TemplateName"): If ExitIfTrue(isFalse(templateName), "TemplateName is empty..") Then Exit Function
    Dim TranslatedSnippet: TranslatedSnippet = frm("TranslatedSnippet"): If ExitIfTrue(isFalse(TranslatedSnippet), "TranslatedSnippet is empty..") Then Exit Function
    
    Dim SnippetDescription: SnippetDescription = "Template - " & templateName
    Dim frm2 As Form
    
    If isPresent("tblSnippets", "SnippetDescription = " & Esc(SnippetDescription)) Then
        Dim resp: resp = MsgBox("Template already exists. Do you want to replace it?", vbYesNo)
        If resp = vbYes Then
            DoCmd.OpenForm "frmSnippets", , , "SnippetDescription = " & Esc(SnippetDescription)
            Set frm2 = Forms("frmSnippets")
            frm2("Snippet") = TranslatedSnippet
        End If
    Else
        
        DoCmd.OpenForm "frmSnippets", , , , acFormAdd
        Set frm2 = Forms("frmSnippets")
        frm2("SnippetDescription") = SnippetDescription
        frm2("Snippet") = TranslatedSnippet
        
    End If
    
End Function
