Attribute VB_Name = "SeqModelSetting Mod"
Option Compare Database
Option Explicit

Public Function SeqModelSettingCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function lblContainerWidth_OnClick(frm As Form)

    frm("ContainerWidth") = "max-w-[750px]"
    
End Function

Public Function GetMaxWidthSize(size)
    
    If isFalse(size) Then
        size = "xl"
    End If
    
    GetMaxWidthSize = "max-w-screen-" & size
    
End Function
