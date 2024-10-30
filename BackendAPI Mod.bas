Attribute VB_Name = "BackendAPI Mod"
Option Compare Database
Option Explicit

Public Function BackendAPICreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function TestApi(frm As Form) As String
    
    Dim endpointUrl As String: endpointUrl = frm("Endpoint")
    If isFalse(endpointUrl) Then Exit Function
    
    endpointUrl = endpointUrl & "?rand=" & Rnd()
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", endpointUrl, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.setRequestHeader "Pragma", "no-cache"
    http.send
    
    Dim StatusCode: StatusCode = http.status
    
    Dim json As Object
    Set json = JsonConverter.ParseJson(http.responseText)
    
    Dim prettyJson As String
    prettyJson = JsonConverter.ConvertToJson(json)
    
    frm("ReturnStringUnformatted") = prettyJson
    prettyJson = UnescapeJson(prettyJson)
    
    TestApi = PrettifyJson(prettyJson)
    TestApi = Left(TestApi, 500)
    frm("StatusCode") = StatusCode
    frm("ReturnString") = TestApi
    
    
End Function

Public Function UnescapeJson(json As String) As String
    json = replace(json, "\/", "/")
    json = replace(json, "\""", """")
    json = replace(json, "\n", vbNewLine)
    json = replace(json, "\r", vbCr)
    json = replace(json, "\t", vbTab)
    UnescapeJson = json
End Function

Public Function PrettifyJson(json As String) As String
    Dim indentLevel As Integer
    Dim inString As Boolean
    Dim result As String
    Dim currentChar As String
    Dim newLine As String
    Dim indentString As String
    
    indentString = "    "
    newLine = vbNewLine
    inString = False
    indentLevel = 0
    
    Dim i
    For i = 1 To Len(json)
        currentChar = Mid(json, i, 1)
        Select Case currentChar
            Case """"
                If Not inString Then
                    inString = True
                    result = result & currentChar
                Else
                    inString = False
                    result = result & currentChar
                End If
            Case "{", "["
                If Not inString Then
                    indentLevel = indentLevel + 1
                    result = result & currentChar & newLine & String(indentLevel, indentString)
                Else
                    result = result & currentChar
                End If
            Case "}", "]"
                If Not inString Then
                    indentLevel = indentLevel - 1
                    result = result & newLine & String(indentLevel, indentString) & currentChar
                Else
                    result = result & currentChar
                End If
            Case ","
                If Not inString Then
                    result = result & currentChar & newLine & String(indentLevel, indentString)
                Else
                    result = result & currentChar
                End If
            Case Else
                result = result & currentChar
        End Select
    Next i
    
    PrettifyJson = result
End Function
