Attribute VB_Name = "Backup Mod"
Option Compare Database
Option Explicit

Public Function BackupCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

'Public Function BackupBackendFile()
'
'    ''Get the current backend path
'    Dim currentBEDirectory, currentBEPath, backedupPath, BEfileName, completeBackupPath
'    currentBEDirectory = GetcurrentBEDirectory()
'    currentBEPath = currentBEDirectory & "PTS Backend.accdb"
'    ''Check if the currentBEPath file exists, if not throw an error
'    If ExitIfTrue(Not DoesFileExist(currentBEPath), EscapeString(currentBEPath) & " does not exist.") Then Exit Function
'    ''Get the path where the backend path will be sent
'
'    backedupPath = GetbackedupPath: If DirectoryExists(backedupPath) Then MakeActualBackup backedupPath, currentBEPath
'    backedupPath = GetbackedupPath("Z:\MY PANDA APP\Backup\"): If DirectoryExists(backedupPath) Then MakeActualBackup backedupPath, currentBEPath
'    backedupPath = GetbackedupPath("C:\Users\appli\OneDrive\MY PANDA REALTY\PANDA APP\Backup\"): If DirectoryExists(backedupPath) Then MakeActualBackup backedupPath, currentBEPath
'
'End Function

Private Function MakeActualBackup(backedupPath, currentBEPath)

    Dim BEfileName, completeBackupPath
    BEfileName = GetBEfileName
    completeBackupPath = backedupPath & BEfileName
    Debug.Print completeBackupPath
    ''Copy the current be to the backed up BE path
    ''Use this snippet to copy
    BackUpCopy currentBEPath, completeBackupPath
        
End Function

Private Function BackUpCopy(currentBEPath, completeBackupPath)
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CopyFile currentBEPath, completeBackupPath, True
    
End Function

Private Function GetBEfileName()
    
    GetBEfileName = Format$(Now(), "YYYYMMDDHHNNSS") & ".accdb"

End Function

Private Function GetbackedupPath(Optional folderPath = "C:\Users\Owner\OneDrive\MY PANDA REALTY\PANDA APP\Backup\")
    
    ''C:\Users\Owner\OneDrive\MY PANDA REALTY\PANDA APP\Backup
    GetbackedupPath = CurrentProject.path & "\Backup\"
    If Environ("computername") <> "LAPTOP-4EL19IO4" Then
        GetbackedupPath = folderPath
    End If

End Function

Private Function GetcurrentBEDirectory()
    
    GetcurrentBEDirectory = CurrentProject.path & "\"
    
    If Environ("computername") <> "LAPTOP-4EL19IO4" Then
        Dim ProjectPath, filePath
        ProjectPath = "Z:\MY PANDA APP"
        If Not DirectoryExists(ProjectPath) Then
            ProjectPath = "\\TRUENAS\database\MY PANDA APP"
            If Not DirectoryExists(ProjectPath) Then
                MsgBox "The database tables can't be linked to the backend file. The app will exit.", vbCritical
                DoCmd.Quit
                Exit Function
            End If
        End If
        GetcurrentBEDirectory = ProjectPath & "\"
    End If
    
End Function
