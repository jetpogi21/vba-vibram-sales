Attribute VB_Name = "SeqModelFile Mod"
Option Compare Database
Option Explicit

Public Function SeqModelFileCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm("FunctionName").OnDblClick = "=RunSeqModelFunctionFromFile([Form])"
            frm("FunctionName").DisplayAsHyperlink = 2
            frm("FilePath").OnDblClick = "=OpenFolderLocation([FilePath])"
            frm("FilePath").DisplayAsHyperlink = 2
            
            frm.AllowAdditions = False
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function RunSeqModelFunctionFromFile(frm As Form)
    
    Dim FunctionName: FunctionName = frm("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
    Dim SeqModelID: SeqModelID = frm("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Run FunctionName, frm, SeqModelID
    
End Function
