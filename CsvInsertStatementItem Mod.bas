Attribute VB_Name = "CsvInsertStatementItem Mod"
Option Compare Database
Option Explicit

Public Function CsvInsertStatementItemCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


