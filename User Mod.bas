Attribute VB_Name = "User Mod"
Option Compare Database
Option Explicit

Public Function UserCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
            ''Create_mainForm_CloseButton frm
            
            frm("cmdAdd").OnClick = "=OpenFormFromMain(""frmUsers_2"")"
            frm("cmdView").OnClick = "=OpenFormFromMain(""frmUsers_2"",""subform"",""UserID"",[Form])"
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function
