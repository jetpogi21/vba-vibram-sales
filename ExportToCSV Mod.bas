Attribute VB_Name = "ExportToCSV Mod"
Option Compare Database
Option Explicit

Public Function ExportToCSVCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function ExportToCSV()
    
    Dim tblName: tblName = "taskmanager_taskinterval"
    DoCmd.TransferText acExportDelim, "Export Specification", tblName, CurrentProject.path & "\" & tblName & ".csv"
    
End Function
