Attribute VB_Name = "Snippet Mod"
Option Compare Database
Option Explicit

Public Function SnippetCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm("Snippet").FontName = "Lucida Console"
            frm("Snippet").fontSize = 7
            
        Case 5: ''Datasheet Form
            
        Case 6: ''Main Form
            frm("lblfltrCategoryID").Width = frm("lblfltrCategoryID").Width * 2 / 3
            frm("txtSearchfltrCategoryID").Left = GetRight(frm("lblfltrCategoryID"))
            frm("txtSearchfltrCategoryID").Width = frm("cleartxtSearchfltrCategoryID").Left - frm("txtSearchfltrCategoryID").Left
            frm.OnLoad = "=mainSnippets_OnLoad([Form])"
        Case 7: ''Tabular Report
    End Select

End Function

Public Function mainSnippets_OnLoad(frm As Object)

    MainFormSQLOnLoad frm, 788
    Set frm = frm("subform").Form
    
    frm.OrderBy = "SnippetID DESC"
    frm.OrderByOn = True
    
End Function


Public Function GenerateSubHeader()

    Dim strInput As String, processedInput As String
    strInput = InputBox("Please enter the subheader title..")
    
    If strInput <> "" Then
        
        processedInput = "**************************************"
        processedInput = processedInput & strInput
        processedInput = processedInput & "**************************************"
        
        CopyToClipboard processedInput
        
    End If

End Function

Public Function ClearSnippetContent(frm As Form)
    
    frm("Snippet") = ""
    
End Function

Public Function CopySnippetContent(frm As Form)

    CopyFieldContent frm, "Snippet"
    
End Function

