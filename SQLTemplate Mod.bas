Attribute VB_Name = "SQLTemplate Mod"
Option Compare Database
Option Explicit

Public Function SQLTemplateCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function GetSQLTemplateFromForm(frm As Form)
    
    Dim SQLTemplateID, SQLTemplateOperation, SQLTemplate
    SQLTemplateID = frm("SQLTemplateID")
    SQLTemplateOperation = frm("SQLTemplateOperation")
    SQLTemplate = frm("SQLTemplate")
    
    If ExitIfTrue(IsNull(SQLTemplateID), "Record is empty...") Then Exit Function
    
    CopyToClipboard SQLTemplate
    
End Function

Public Function GetSQLTemplate(Optional SQLTemplateOperation = "SELECT")

    CopyToClipboard ELookup("tblSQLTemplates", "SQLTemplateOperation = " & EscapeString(SQLTemplateOperation), "SQLTemplate")
    
End Function


