Attribute VB_Name = "VBA IDE Module"
Option Compare Database
Option Explicit

Public Sub CreateOnListEvent()
    
    DoCmd.OpenForm "dshtBookDetails", acDesign
    Forms!dshtBookDetails.HasModule = True

    Dim VBProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim lineNum As Long
    
    Set VBProj = Application.VBE.VBProjects("Database")
    If DoesPropertyExists(VBProj.VBComponents, "Form_dshtBookDetails") Then
        Set vbComp = VBProj.VBComponents("Form_dshtBookDetails")
    Else
        Set vbComp = VBProj.VBComponents.Add(vbext_ct_MSForm)
        vbComp.Name = "Form_dshtBookDetails"
    End If
    
    Set CodeMod = vbComp.CodeModule
    
    With CodeMod
        lineNum = .CreateEventProc("NotInList", "StudentID")
        lineNum = lineNum + 1
        .InsertLines lineNum, "    MsgBox " & EscapeString("Hello World")
    End With
    
End Sub

