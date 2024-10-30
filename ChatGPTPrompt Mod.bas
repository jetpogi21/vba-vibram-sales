Attribute VB_Name = "ChatGPTPrompt Mod"
Option Compare Database
Option Explicit

Public Function ChatGPTPromptCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm("Content").Height = GetBottom(frm("Note")) - frm("Content").Top
            frm("lblNote").Width = frm("PrePrompt").Width
            frm("Note").Width = frm("PrePrompt").Width
            
            Dim ctl As control: Set ctl = CreateButtonControl(frm, "Open ChatGPT", "cmdChatGPT", "=GoToLink(" & Esc("https://chatgpt.com/") & ")")
            
            CopyControlProperty ctl, frm("cmdClearChatGPTPrompt"), "Width"
            AlignControl ctl, frm("cmdClearChatGPTPrompt"), "Top"
            MoveControl ctl, frm("cmdClearChatGPTPrompt")
            
            Set ctl = CreateButtonControl(frm, "Open Gemini", "cmdGemini", "=GoToLink(" & Esc("https://gemini.google.com/app") & ")")
            
            CopyControlProperty ctl, frm("cmdChatGPT"), "Width"
            AlignControl ctl, frm("cmdChatGPT"), "Top"
            MoveControl ctl, frm("cmdChatGPT")
            
            ResizeForm frm
            
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function CopyCombinedPrompt(frm As Object, Optional ChatGPTPromptID = "")

    RunCommandSaveRecord

    If isFalse(ChatGPTPromptID) Then
        ChatGPTPromptID = frm("ChatGPTPromptID")
        If ExitIfTrue(isFalse(ChatGPTPromptID), "ChatGPTPromptID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblChatGPTPrompts WHERE ChatGPTPromptID = " & ChatGPTPromptID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim content: content = rs.fields("Content"): If ExitIfTrue(isFalse(content), "Content is empty..") Then Exit Function
    
    Dim PrePrompt: PrePrompt = rs.fields("PrePrompt")
    If Not IsNull(PrePrompt) Then
        lines.Add PrePrompt
        content = "My prompt is: " & content
    End If
    
    lines.Add content
    
    CopyToClipboard lines.JoinArr(vbNewLine & vbNewLine)
    
End Function

Public Function ClearChatGPTPrompt(frm As Object, Optional ChatGPTPromptID = "")

    frm("PrePrompt") = ""
    
End Function
