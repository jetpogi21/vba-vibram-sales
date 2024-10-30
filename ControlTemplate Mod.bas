Attribute VB_Name = "ControlTemplate Mod"
Option Compare Database
Option Explicit

Public Function ControlTemplateCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Sub SyncControlTemplateOfForm(frmName, Optional closeForm As Boolean = False)

    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    CopyControlTemplateProperties frm
    
    Dim ctl As control
    For Each ctl In frm.controls
        CopyControlTemplateProperties frm, ctl.Name
    Next ctl
    
    If closeForm Then
        DoCmd.Close acForm, frmName, acSaveYes
    End If
    
End Sub


