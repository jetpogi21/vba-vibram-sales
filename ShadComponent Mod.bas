Attribute VB_Name = "ShadComponent Mod"
Option Compare Database
Option Explicit

Public Function ShadComponentCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function
''ShadComponent_ShadComponent_AfterUpdate

Public Function ShadComponent_ShadComponent_AfterUpdate(frm As Form)

    Dim ShadComponent: ShadComponent = frm("ShadComponent"): If ExitIfTrue(isFalse(ShadComponent), "ShadComponent is empty..") Then Exit Function
    ShadComponent = replace(ShadComponent, "-", " ")
    ShadComponent = StrConv(ShadComponent, vbProperCase)
    ShadComponent = replace(ShadComponent, " ", "")
    
    frm("ShadComponentFile") = ShadComponent
    
End Function
