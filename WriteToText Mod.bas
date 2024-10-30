Attribute VB_Name = "WriteToText Mod"
Option Compare Database
Option Explicit

Public Function WriteToTextCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Sub AppendTextToFile(ByVal strText As String, ByVal strPath As String, Optional ByVal overwriteContent As Boolean = False)
    Dim objFSO As Object
    Dim objFile As Object
    Dim strContents As String
    
    ' Check if the file exists
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.fileExists(strPath) Then
        ' If not, create the file
        Set objFile = objFSO.CreateTextFile(strPath, True)
    Else
        ' If yes, check if we should overwrite the content
        If overwriteContent Then
            ' Overwrite the content of the file
            Set objFile = objFSO.OpenTextFile(strPath, 2)   '   2 means overwrite mode
        Else
            ' Read the contents of the file
            Set objFile = objFSO.OpenTextFile(strPath, 1)   '   1 means read mode
            strContents = objFile.ReadAll
            objFile.Close
            
            ' Append a newline character if the file is not empty
            If Len(strContents) > 0 Then
                strText = vbCrLf & strText
            End If
            
            ' Open the file again in append mode
            Set objFile = objFSO.OpenTextFile(strPath, 8)   '   8 means append mode
        End If
    End If
    
    ' Write the text to the file
    objFile.Write strText
    ' Close the file
    objFile.Close
    
    ' Release objects
    Set objFile = Nothing
    Set objFSO = Nothing
End Sub

Public Function WriteToFile(filePath, Text, Optional SeqModelID = "", Optional FunctionName = "")

    DoCmd.OpenForm "frmClipboardForms"
    Forms("frmClipboardForms")("Snippet") = Text
    CopyFieldContent Forms("frmClipboardForms"), "Snippet"
    
    Dim IsProtected: IsProtected = False
    If Not isFalse(FunctionName) Then
        Dim Note: Note = ELookup("qryFunctionChainItems", "FunctionName = " & Esc(FunctionName), "Note")
        IsProtected = Note Like "*Special File*"
    End If
    
    ''Check first if this file is protected
    If isPresent("tblSeqModelFiles", "filePath = " & Esc(filePath) & " AND IsProtected") And Not isFalse(SeqModelID) Then
        If NoHasWriteToFilePrompt = False Then
            MsgBox "The file at " & Esc(filePath) & " is protected.", vbCritical + vbOKOnly
        End If
        RunSQL "UPDATE tblSeqModelFiles SET FunctionName = " & Esc(FunctionName) & ", IsProtected = " & IsProtected & " WHERE SeqModelID = " & SeqModelID & " AND filePath = " & Esc(filePath)
        Exit Function
    End If
    
    If isPresent("tblBackendProjectFiles", "filePath = " & Esc(filePath) & " AND IsProtected") Then
        If NoHasWriteToFilePrompt = False Then
            MsgBox "The file at " & Esc(filePath) & " is protected.", vbCritical + vbOKOnly
        End If
        Exit Function
    End If
    
    DoCmd.Close acForm, "frmClipboardForms", acSaveNo

    If NoHasWriteToFilePrompt = False Then
        Dim resp: resp = MsgBox("This will replace the file currently on " & Esc(filePath) & "." & vbCrLf & "Do you want to proceed?", vbYesNo)
        If resp = vbNo Then Exit Function
    End If
    
    Dim folderPath As String
    
    ' Extract the folder path from the file path
    folderPath = Left(filePath, InStrRev(filePath, "\"))
    
    ' Create the folder if it doesn't exist
    If Len(folderPath) > 0 And Dir(folderPath, vbDirectory) = "" Then
        CreateDirectories folderPath
    End If
    
    ' Open a text file for writing
    Open filePath For Output As #1
    
    ' Write some text to the file
    Print #1, Text
    
    ' Close the file
    Close #1
    
    If Not isFalse(SeqModelID) Then
        If Not isPresent("tblSeqModelFiles", "SeqModelID = " & SeqModelID & " AND filePath = " & Esc(filePath)) Then
            RunSQL "INSERT INTO tblSeqModelFiles (SeqModelID,filePath,FunctionName,IsProtected) VALUES (" & _
                SeqModelID & "," & Esc(filePath) & "," & Esc(FunctionName) & "," & IsProtected & ")"
        Else
            RunSQL "UPDATE tblSeqModelFiles SET FunctionName = " & Esc(FunctionName) & ",IsProtected = " & IsProtected & " WHERE SeqModelID = " & SeqModelID & " AND filePath = " & Esc(filePath)
        End If
        ''Generate the getModelOptions for the BackendFile in reference to the SeqModelID
        If IsProtected Then
            WriteToGetmodeloptions_tsx Forms(0), ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "BackendProjectID")
        End If
    End If
    
    ' Return True to indicate success
    WriteToFile = True

End Function

Public Sub CreateDirectories(ByVal folderPath As String)
    Dim folderArray() As String
    folderArray = Split(folderPath, "\")
    Dim currentPath As String
    currentPath = ""
    Dim i
    For i = LBound(folderArray) To UBound(folderArray)
        If i = UBound(folderArray) Then
            currentPath = currentPath & folderArray(i)
        Else
            currentPath = currentPath & folderArray(i) & "\"
        End If
        
        If Len(Dir(currentPath, vbDirectory)) = 0 Then
            MkDir currentPath
        End If
    Next i
End Sub

