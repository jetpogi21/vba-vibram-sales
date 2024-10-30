Attribute VB_Name = "GitDiff Mod"
Option Compare Database
Option Explicit

Public Function CopyGitDiffToClipboardWithProjectPath(ctl As IRibbonControl)
    
    Dim templateName: templateName = "git diff detailed"
    Dim wsh As Object
    Dim waitOnReturn As Boolean
    Dim windowStyle As Integer

    ' Set waitOnReturn to True if you want to wait for the command to complete
    waitOnReturn = True
    ' Set windowStyle to 1 to show the command window, or 0 to hide it
    windowStyle = 1
    
    Dim ProjectPath: ProjectPath = InputBox("Paste the project directory of a local git repository..")
    
    If isFalse(ProjectPath) Then Exit Function
    
    If Right(ProjectPath, 1) = "\" Then
        ProjectPath = Left(ProjectPath, Len(ProjectPath) - 1)
    End If
    
    ' Check if the path contains "\src" and extract the substring before it
    If InStr(ProjectPath, "\src") > 0 Then
        ProjectPath = Left(ProjectPath, InStrRev(ProjectPath, "\src") - 1)
    End If
    
    ''src\app\resume\_components\MyDocument.tsx
    
    Dim diffFile: diffFile = "my.diff"
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /c cd /d " & Esc("C:\Users\User\Desktop\Programming Tools\Programming Guides")
    shellArr.Add "py ""git diff.py"" " & Esc(ProjectPath)

    Set wsh = CreateObject("WScript.Shell")
    
    wsh.Run shellArr.JoinArr(" & "), windowStyle, waitOnReturn

    ''Read the content of the text file
    Dim filePath As String: filePath = ProjectPath & "\" & diffFile
    Dim fileContent: fileContent = ReadTextFile(filePath)
    
    If isFalse(fileContent) Then
        MsgBox "File content is empty.", vbCritical + vbOKOnly
        Exit Function
    End If

    CopyGitDiffToClipboardWithProjectPath = GetTemplateContent(templateName)
    CopyGitDiffToClipboardWithProjectPath = replace(CopyGitDiffToClipboardWithProjectPath, "[gitdiff]", fileContent)
    CopyToClipboard CopyGitDiffToClipboardWithProjectPath
    
    MsgBox "Prompt copied to clipboard.."

End Function

Public Function CopyThisGitDiffToClipboard(ctl As IRibbonControl)

    Dim wsh As Object
    Dim waitOnReturn As Boolean
    Dim windowStyle As Integer

    ' Set waitOnReturn to True if you want to wait for the command to complete
    waitOnReturn = True
    ' Set windowStyle to 1 to show the command window, or 0 to hide it
    windowStyle = 1
    
    Dim ProjectPath: ProjectPath = CurrentProject.path & "\VBA Code\"
    ProjectPath = Left(ProjectPath, Len(ProjectPath) - 1)
    Dim diffFile: diffFile = "my.diff"
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /c cd /d " & Esc(CurrentProject.path)
    shellArr.Add "py ""git diff.py"" " & Esc(ProjectPath)

    Set wsh = CreateObject("WScript.Shell")
    
    wsh.Run shellArr.JoinArr(" & "), windowStyle, waitOnReturn

    ''Read the content of the text file
    Dim filePath As String: filePath = ProjectPath & "\" & diffFile
    Dim fileContent: fileContent = ReadTextFile(filePath)
    
    If isFalse(fileContent) Then
        MsgBox "File content is empty.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    CopyThisGitDiffToClipboard = GetTemplateContent("git diff detailed")
    CopyThisGitDiffToClipboard = replace(CopyThisGitDiffToClipboard, "[gitdiff]", fileContent)
    CopyToClipboard CopyThisGitDiffToClipboard
    
    MsgBox "Prompt copied to clipboard.."
    
    GoToLink "https://chatgpt.com/"
    
End Function
