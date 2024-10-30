Attribute VB_Name = "File Handling"
Option Compare Database
Option Explicit

Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As LongPtr
    
Public Sub ImportCSVToTable(strPath, Optional tblName = "tblCSVData")
    
    deleteTableIfExists tblName
    DoCmd.TransferText acImportDelim, , tblName, strPath, True
    
End Sub
    
Public Function RenameFile(ByVal filePath As String, ByVal updatedName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Get the current directory path
    Dim directoryPath As String
    directoryPath = Left(filePath, InStrRev(filePath, "\"))
    
    ' Build the new file path with the updated name
    Dim newFilePath As String
    newFilePath = directoryPath & updatedName
    
    ' Rename the file
    Name filePath As newFilePath
    
    ' Check if the renaming was successful
    If Dir(newFilePath) <> "" Then
        RenameFile = True
    Else
        RenameFile = False
    End If
    
    Exit Function
    
ErrorHandler:
    RenameFile = False
End Function

Public Function ReadTextFile(filePath) As String
    Dim fileContent As String
    Dim fileNumber As Integer
    
    ' Open the file for reading
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    
    ' Read the content of the file
    fileContent = Input$(LOF(fileNumber), fileNumber)
    
    ' Close the file
    Close fileNumber
    
    ' Return the file content
    ReadTextFile = fileContent
End Function


Public Function SanitizeFileName(fileName)
    
    Dim bannedArr As New clsArray, i
    bannedArr.arr = "#,%,&,{,},\,<,>,*,?,/,$,!,',"",:,@,+,`,|,="
    
    For Each i In bannedArr.arr
        fileName = replace(fileName, i, "-")
    Next i
    
    SanitizeFileName = fileName
    
End Function

Public Function IsFileOpen(fileName)

    Dim fileNum As Integer
    Dim errNum As Integer
    
    'Allow all errors to happen
    On Error Resume Next
    fileNum = FreeFile()
    
    'Try to open and close the file for input.
    'Errors mean the file is already open
    Open fileName For Input Lock Read As #fileNum
    Close fileNum
    
    'Get the error number
    errNum = Err
    
    'Do not allow errors to happen
    On Error GoTo 0
    
    'Check the Error Number
    Select Case errNum
    
        'errNum = 0 means no errors, therefore file closed
        Case 0
        IsFileOpen = False
     
        'errNum = 70 means the file is already open
        Case 70
        IsFileOpen = True
    
        'Something else went wrong
        Case Else
        IsFileOpen = errNum
    
    End Select

End Function


Function fileExists(filePath) As Boolean
 
    '--------------------------------------------------
    'Checks if a file exists (using the Dir function).
    '--------------------------------------------------
 
    On Error Resume Next
    If Len(filePath) > 0 Then
        If Not Dir(filePath, vbDirectory) = vbNullString Then fileExists = True
    End If
    On Error GoTo 0
 
End Function

Public Function DirectoryExists(path) As Boolean
On Error GoTo DirError:
    DirectoryExists = Dir(path, vbDirectory) <> ""
    Exit Function
DirError:
    If Err.Number = 52 Then
        DirectoryExists = False
    End If
End Function

Public Function SetDefaultDirectory(ctl As control) As String

    ''Open the fileDialog selecting a directory
    Dim fd As FileDialog
    Dim strPath As String
    ' Set up the File Dialog.
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        ' Allow user to make multiple selections in dialog box
        .AllowMultiSelect = False
             
        ' Set the title of the dialog box.
        .Title = "Please select a directory"
        
        If .Show = True Then
            If .SelectedItems.count > 0 Then
                'get the file path selected by the user
                ctl = .SelectedItems(1)
                Exit Function
            End If
        End If
         
    End With
    
    
    ctl = Environ$("USERPROFILE") & "\Downloads"
   
End Function

Public Function PromptDirectory(defaultFileName As String) As String

    ''Open the fileDialog selecting a directory
    Dim fd As FileDialog
    Dim strPath As String
    ' Set up the File Dialog.
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    Dim filePath As String
    
    With fd
        ' Allow user to make multiple selections in dialog box
        .AllowMultiSelect = False
             
        ' Set the title of the dialog box.
        .Title = "Please select a save directory"
        .InitialFileName = defaultFileName
        
        If .Show = True Then
            If .SelectedItems.count > 0 Then
                'get the file path selected by the user
                PromptDirectory = .SelectedItems(1)
                Exit Function
            End If
        End If
         
    End With
    
    
    PromptDirectory = ""
   
End Function

Public Function SetFilePath(frm As Object, FileType As String, fieldName)

    Dim filePath: filePath = PromptFile(FileType)
    frm(fieldName) = filePath
    
End Function

Public Function PromptFile(Optional FileType As String = "") As String
    
    ''Open the fileDialog selecting a directory
    Dim fd As FileDialog
    Dim strPath As String
    ' Set up the File Dialog.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim filePath As String
    With fd
        Dim filterStr, Title
        
        If FileType <> "" Then
            Dim rs As Recordset
            Set rs = ReturnRecordset("SELECT * FROM tblFileDialogs WHERE FileType = '" & FileType & "'")
            fd.filters.Add rs.fields("FileType"), rs.fields("FileFilters")
            fd.Title = rs.fields("Title")
        End If

        ' Allow user to make multiple selections in dialog box
        .AllowMultiSelect = False
        
        If .Show = True Then
            If .SelectedItems.count > 0 Then
                'get the file path selected by the user
                PromptFile = .SelectedItems(1)
                Exit Function
            End If
        End If
         
    End With
    
    
    PromptFile = ""
   
End Function


Public Function directoryPath(fieldName As String, Optional GlobalSettingName As String) As String
    
    Dim strPath As String
    If fieldName <> "" Then
        strPath = ELookup("tblUsers", "UserID = " & g_userID, fieldName)
    End If
    
    ''If strPath is "", set the filepath to Default Download Path
    If strPath = "" Then
    
        If Not isFalse(GlobalSettingName) Then
            strPath = ELookup("tblGlobalSettings", "GlobalSetting = '" & GlobalSettingName & "'", "GlobalSettingValue")
            If strPath = "" Then
                ShowError "A fallback path was not defined under the global setting"
                directoryPath = ""
                Exit Function
            End If
        Else
            ShowError "A fallback path was not defined under the global setting"
            directoryPath = ""
            Exit Function
        End If

    End If
    
    ''Validate filePath if existing
    If Not DirectoryExists(strPath) Then
        MsgBox "The directory path: """ & strPath & """ is not a valid directory...", vbCritical + vbOKOnly
        directoryPath = ""
        Exit Function
    End If
    
    directoryPath = strPath
    
End Function

Public Function SelectDEDirectory(frm As Object, fldName, Optional AdditionalTrigger = Null)
    
    ''Open the fileDialog selecting a directory
    Dim fd As FileDialog
    ' Set up the File Dialog.
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    Dim directory
    
    With fd
        ' Allow user to make multiple selections in dialog box
        .AllowMultiSelect = False
             
        ' Set the title of the dialog box.
        .Title = "Please select a save directory"
        ''.InitialFileName = CurrentProject.path
        
        If .Show Then
            
            'get the file path selected by the user
            directory = .SelectedItems(1) & "\"
           
        End If
         
    End With
    
    frm(fldName) = directory
    
    If Not IsNull(AdditionalTrigger) Then Run AdditionalTrigger, frm
    
End Function

Public Function DeleteFileIfExists(filePath As String) As Boolean
    ' Check if the file exists
    If Dir(filePath) <> "" Then
        ' File exists, attempt to delete it
        On Error Resume Next ' Ignore errors if the file cannot be deleted
        Kill filePath
        On Error GoTo 0 ' Reset error handling
        
        ' Check if the file was successfully deleted
        If Dir(filePath) = "" Then
            ' File was successfully deleted
            DeleteFileIfExists = True
        Else
            ' File could not be deleted
            DeleteFileIfExists = False
        End If
    Else
        ' File does not exist, no action needed
        DeleteFileIfExists = True
    End If
End Function

Public Function CreateFolder(folderPath) As Boolean
    Dim pathParts As Variant
    Dim currentPath As String
    Dim i As Integer
    
    ' Split the folder path into parts
    pathParts = Split(folderPath, "\")
    
    ' Start with an empty string for the current path
    currentPath = ""
    
    ' Loop through each part of the path
    For i = LBound(pathParts) To UBound(pathParts)
        ' Add the current part to the current path
        currentPath = currentPath & pathParts(i) & "\"
        
        ' Check if the current path exists
        If Dir(currentPath, vbDirectory) = "" Then
            ' If not, create the directory
            On Error Resume Next ' In case of error (e.g., permission issues)
            MkDir currentPath
            On Error GoTo 0 ' Reset error handling
            
            ' Check if the directory was successfully created
            If Dir(currentPath, vbDirectory) = "" Then
                CreateFolder = False
                Exit Function
            End If
        End If
    Next i
    
    ' If all directories were successfully created, return True
    CreateFolder = True
End Function

Function CopyFolderContents(SourceFolder As String, DestinationFolder As String, Optional CopyTheSourceFolderItself As Boolean = False)

    Dim fso As Object
    Dim Source As Object
    Dim Destination As Object
    Dim SubFolder As Object
    Dim File As Object
    Dim NewFolder As String

    'Create the File System Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Check if the source folder exists
    If fso.FolderExists(SourceFolder) = False Then
        MsgBox "Source folder does not exist"
        Exit Function
    End If

    'Check if the destination folder exists, if not create it
    If fso.FolderExists(DestinationFolder) = False Then
        NewFolder = fso.BuildPath(fso.GetParentFolderName(DestinationFolder), fso.GetFileName(DestinationFolder))
        fso.CreateFolder (NewFolder)
    End If

    'Set the source folder object
    Set Source = fso.GetFolder(SourceFolder)

    'If CopyTheSourceFolderItself is True, create a new folder inside the destination folder with the name of the source folder
    If CopyTheSourceFolderItself Then
        NewFolder = fso.BuildPath(DestinationFolder, fso.GetFileName(SourceFolder))
        If fso.FolderExists(NewFolder) = False Then
            fso.CreateFolder NewFolder
        End If
        Set Destination = fso.GetFolder(NewFolder)
        DestinationFolder = NewFolder
    Else
        Set Destination = fso.GetFolder(DestinationFolder)
    End If

    'Copy all files and subfolders to the destination folder
    For Each SubFolder In Source.SubFolders
        NewFolder = fso.BuildPath(DestinationFolder, Split(SubFolder.path, "\")(UBound(Split(SubFolder.path, "\"))))
        CopyFolderContents SubFolder.path, NewFolder
    Next SubFolder

    For Each File In Source.files
        Dim DestinationFileName: DestinationFileName = fso.BuildPath(DestinationFolder, File.Name)
        If isPresent("tblBackendProjectFiles", "filePath = " & Esc(DestinationFileName) & " AND IsProtected") Then
            ''MsgBox "The file at " & Esc(filePath) & " is protected.", vbCritical + vbOKOnly
            GoTo NextFile
        End If
        fso.CopyFile File.path, fso.BuildPath(DestinationFolder, File.Name), True
NextFile:
    Next File

    'Clean up the objects
    Set File = Nothing
    Set SubFolder = Nothing
    Set Destination = Nothing
    Set Source = Nothing
    Set fso = Nothing

End Function

'Public Function OpenFolderLocation(filePath As String)
'    Shell "explorer.exe /select," & filePath
'End Function

Public Function OpenFolderLocation(filePath)
    
    If ExitIfTrue(isFalse(filePath), "filePath is empty..") Then Exit Function
    
    If Not isFalse(filePath) Then
        ShellExecute 0, "open", filePath, vbNullString, vbNullString, vbNormalFocus 'open the folder in a non-minimized state
    End If
    
End Function

Sub CopyFile(sourceFilePath As String, destinationFilePath As String)
    On Error Resume Next
    ' Disable error handling temporarily to check if the destination file already exists
    
    ' Check if the source file exists
    If Dir(sourceFilePath) = "" Then
        MsgBox "Source file does not exist."
        Exit Sub
    End If
    
    ' Check if the destination file already exists
    If Dir(destinationFilePath) <> "" Then
        ' File already exists, prompt to overwrite
        If MsgBox("Destination file already exists. Do you want to overwrite it?", vbYesNo + vbQuestion) <> vbYes Then
            Exit Sub
        End If
        ' Delete the existing file
        Kill destinationFilePath
    End If
    
    ' Copy the file to the destination path
    FileCopy sourceFilePath, destinationFilePath
    
    ' Check if the copy operation was successful
    If Err.Number <> 0 Then
        MsgBox Err.Number & " " & Err.description & " An error occurred while copying the file."
    Else
        MsgBox "File copied successfully."
    End If
    
    On Error GoTo 0
    ' Re-enable error handling
End Sub

Public Function CopyFileToTemplateFolder(sourcePath As String, Optional IsSupabase As Boolean = False)

    ''A function that will copy sourcePath to the destination path
    Dim TemplatePath, ProjectPath
    TemplatePath = IIf(IsSupabase, "\Files\Next 13 templates - supabase\", "\Files\Next 13 templates\")
    ProjectPath = CurrentProject.path & TemplatePath

    ''Check if the projectPath exists if not then msgbox the error
    If Not DirectoryExists(ProjectPath) Then
        MsgBox "Project path does not exist: " & ProjectPath
        Exit Function
    End If
    
    ''Check if the sourcePath exists if not then msgbox the error
    If Not fileExists(sourcePath) Then
        MsgBox "Source file does not exist: " & sourcePath
        Exit Function
    End If
    ''example sourcePath "C:\Users\User\Desktop\Web Development\task-manager-next-13\src\components\form\FormikControl.tsx"
    ''Get the destinationPath by replacing the string before the src of the sourcePath with the projectPath.
    ''destinationPath will now be projectPath & "src\components\form\FormikControl.tsx"
    Dim fileName As String
    fileName = Mid(sourcePath, InStrRev(sourcePath, "\") + 1)
    Dim destinationPath As String
    destinationPath = ProjectPath & Mid(sourcePath, InStr(sourcePath, "src"))
    
    ''if the destinationPath doesn't exist then create the directory
    If Not DirectoryExists(Left(destinationPath, InStrRev(destinationPath, "\"))) Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        CreateFolder Left(destinationPath, InStrRev(destinationPath, "\"))
        ''fso.CreateFolder Left(destinationPath, InStrRev(destinationPath, "\"))
    End If

    ''use fso.buildPath and fso.createFolder if necessary
    ' Copy the file from sourcePath to destinationPath
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFile sourcePath, destinationPath
    
    MsgBox Esc(sourcePath) & " copied to " & Esc(destinationPath)
    
End Function

Public Function OpenFileWithDefaultProgram(filePath As String) As Boolean
    On Error Resume Next
    Shell filePath, vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "An error occurred while trying to open the file. Error number: " & Err.Number, vbCritical, "Error"
        OpenFileWithDefaultProgram = False
    Else
        OpenFileWithDefaultProgram = True
    End If
    On Error GoTo 0
End Function

Public Function GetFilePaths(directoryPath, Optional fileExtension = "") As clsArray

    Dim filesArr As New clsArray
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FolderExists(directoryPath) Then Exit Function

    Set objFolder = objFSO.GetFolder(directoryPath)
    
    For Each objFile In objFolder.files
        ' Check if fileExtension is provided and matches the file extension
        If fileExtension = "" Or objFSO.GetExtensionName(objFile.path) = fileExtension Then
            filesArr.Add objFile.path
        End If
    Next objFile
    
    ' Clean up
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
    
    Set GetFilePaths = filesArr
    
End Function







