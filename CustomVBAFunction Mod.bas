Attribute VB_Name = "CustomVBAFunction Mod"
Option Compare Database

Public Function CustomVBAFunctionCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

''Command Name: Hook to Model Buttons
Public Function HookTo_ModelButtonsFromCustomVBAFunction(frm As Object, Optional CustomVBAFunctionID = "")

    Dim CustomVBAFunction: CustomVBAFunction = frm("CustomVBAFunction")
    Dim moduleName: moduleName = frm("ModuleName")
    Dim ModelID: ModelID = ELookup("tblModels", "Model = " & Esc(RemoveMatchedPattern(moduleName, " Mod")), "ModelID")
    
    Dim fields As New clsArray: fields.arr = "ModelID,ModelButton,FunctionName,ModelButtonOrder"
    Dim fieldValues As New clsArray
    Set fieldValues = New clsArray
    
    fieldValues.Add ModelID
    fieldValues.Add GetButtonCaptionFromFunctionName(CustomVBAFunction)
    fieldValues.Add CustomVBAFunction
    
    Dim ModelButtonOrder: ModelButtonOrder = ELookup("tblModelButtons", "ModelID = " & ModelID, "ModelButtonOrder", "ModelButtonOrder DESC")
    fieldValues.Add CDbl(ModelButtonOrder) + 1
    
    UpsertRecord "tblModelButtons", fields, fieldValues, "FunctionName = " & Esc(CustomVBAFunction)
    
End Function
