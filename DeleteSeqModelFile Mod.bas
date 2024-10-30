Attribute VB_Name = "DeleteSeqModelFile Mod"
Option Compare Database
Option Explicit

Public Function DeleteSeqModelFileCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function DeleteSeqModelFiles(frm As Object, Optional disableNotif As Boolean = False)

    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim FileNamePattern: FileNamePattern = frm("FileNamePattern"): If ExitIfTrue(isFalse(FileNamePattern), "FileNamePattern is empty..") Then Exit Function
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\components\sales\SaleFilterForm.tsx
    ''[ClientPath]src\components\[ModelPath]\[ModelName]FilterForm.tsx
    ''C:\Users\User\Desktop\Web Development\vibram-sales\
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID)
    Dim deletedFiles As New clsArray
    Do Until rs.EOF
        Dim ReplacedPattern As String: ReplacedPattern = GetReplacedTemplate(rs, "", , FileNamePattern)
        If DeleteFileIfExists(ReplacedPattern) Then
            deletedFiles.Add ReplacedPattern
        End If
        rs.MoveNext
    Loop
    
    If deletedFiles.count > 0 And Not disableNotif Then
        MsgBox "Files Deleted: " & vbCrLf & deletedFiles.JoinArr(vbCrLf), vbOKOnly
    End If
    
End Function
