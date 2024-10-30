Attribute VB_Name = "CopySourceFile Mod"
Option Compare Database
Option Explicit

Public Function CopySourceFileCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm("cmdCopyFile").OnClick = "=CopyFileToTemplateFolderFromForm([Form], True)"
            frm("SourceFile").OnClick = "=SelectAllContent([Form], ""SourceFile"")"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function CopyFileToTemplateFolderFromForm(frm As Object, Optional IsSupabase As Boolean = False)
    
    Dim sourceFile As String: sourceFile = frm("SourceFile")
    If ExitIfTrue(isFalse(sourceFile), "Source file is empty..") Then Exit Function
    
    CopyFileToTemplateFolder sourceFile, IsSupabase
    
End Function
