Attribute VB_Name = "ChatGPTMessage Mod"
Option Compare Database
Option Explicit

Public Function ChatGPTMessageCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SendGPTMessagePrompt(frm As Form)
    
    RunCommandSaveRecord
    Dim API As String: API = "https://api.openai.com/v1/chat/completions"
    Dim GPTModel As String: GPTModel = "gpt-3.5-turbo"
    
    ''TABLE: tblChatGPTMessages Fields: ChatGPTMessageID|ChatGPTThreadID|UserPrompt|AssistantPrompt|Timestamp
    ''CreatedBy|RecordImportID
    
    ''TABLE: tblChatGPTThreads Fields: ChatGPTThreadID|Timestamp|CreatedBy|RecordImportID|SystemMessage
    Dim ChatGPTThreadID: ChatGPTThreadID = frm("ChatGPTThreadID")
    
    Dim SystemMessage: SystemMessage = ELookup("tblChatGPTThreads", "ChatGPTThreadID = " & ChatGPTThreadID, "SystemMessage")
    
    Dim ChatGPTMessageID: ChatGPTMessageID = frm("ChatGPTMessageID")
    
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblChatGPTMessages WHERE ChatGPTThreadID = " & ChatGPTThreadID & " AND ChatGPTMessageID < " & ChatGPTMessageID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    Dim UserPrompt, AssistantPrompt, msgCollection As New Collection
    
    Dim isFirst: isFirst = True
    Do Until rs.EOF
        UserPrompt = rs.fields("UserPrompt")
        AssistantPrompt = rs.fields("AssistantPrompt")
        If isFirst And Not isFalse(SystemMessage) Then
            isFirst = False
            UserPrompt = SystemMessage & "." & UserPrompt
        End If
        
        If Not isFalse(UserPrompt) Then msgCollection.Add CreateGPTMessageItem("user", UserPrompt)
        If Not isFalse(AssistantPrompt) Then msgCollection.Add CreateGPTMessageItem("assistant", JsonEscape(AssistantPrompt))
        
        rs.MoveNext
    Loop
    
    UserPrompt = frm("UserPrompt")
    If isFirst And Not isFalse(SystemMessage) Then
        isFirst = False
        UserPrompt = SystemMessage & ". " & UserPrompt
    End If
    If Not isFalse(UserPrompt) Then msgCollection.Add CreateGPTMessageItem("user", UserPrompt)
    
    Dim Messages: Messages = BuildMessages(msgCollection)
    
    ''Build the body object
    Dim BodyObject As New clsDictionary: Set BodyObject = GetBodyObject(GPTModel, Messages)
    
    ''Debug.Print BodyObject.ToFormatString
    Debug.Print BodyObject.ToFormatString
    Debug.Print ""
    
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")

    xmlhttp.Open "POST", API, False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "Authorization", "Bearer [SomeAPIKey]"

    xmlhttp.send BodyObject.ToFormatString

    Dim response: response = xmlhttp.responseText
    
    Dim prompt_tokens, completion_tokens, total_tokens, content
    Dim dict As Object: Set dict = JsonConverter.ParseJson(response)
    
    prompt_tokens = dict("usage")("prompt_tokens")
    completion_tokens = dict("usage")("completion_tokens")
    total_tokens = dict("usage")("total_tokens")
    content = dict("choices")(1)("message")("content")
    ''Remove the first 2 new lines
    content = Right(content, Len(content) - 4)
    
    frm("PromptTokens") = prompt_tokens
    frm("CompletionTokens") = completion_tokens
    frm("TotalTokens") = total_tokens
    frm("AssistantPrompt") = content
    
End Function

Public Function JsonEscape(ByVal sText As String) As String
    sText = replace(sText, "\", "\\") ' Escape backslashes
    sText = replace(sText, """", "\""") ' Escape quotes
    sText = replace(sText, vbNewLine, "\n") ' Escape new lines
    sText = replace(sText, vbCr, "\r") ' Escape carriage returns
    sText = replace(sText, vbTab, "\t") ' Escape tabs
    sText = replace(sText, "/", "\/") ' Escape forward slashes
    ' Add any additional replacements for other special characters as needed
    JsonEscape = sText
End Function

Private Function CreateGPTMessageItem(role, content) As clsDictionary
    
    ''role can be: user, assistant or system
    ''{"role": "user", "content": "What is the OpenAI mission?"}
    Dim dict As New clsDictionary
    dict.Add Esc("role"), Esc(role)
    dict.Add Esc("content"), Esc(content)
    
    Set CreateGPTMessageItem = dict
    
End Function

Private Function BuildMessages(msgCollection As Collection)

    Dim msg, msgs As New clsArray
    For Each msg In msgCollection
        msgs.Add msg.ToFormatString
    Next msg
    
    BuildMessages = "[" & msgs.JoinArr & "]"
    
End Function

Private Function GetBodyObject(GPTModel, Messages) As clsDictionary

    Dim dict As New clsDictionary
    dict.Add Esc("model"), Esc(GPTModel)
    dict.Add Esc("messages"), Messages
    
    Set GetBodyObject = dict
    
End Function
