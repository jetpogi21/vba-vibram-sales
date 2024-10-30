Attribute VB_Name = "Model2 Mod"
Option Compare Database
Option Explicit

Public Function GenerateAdditionalOptionButton(frm As Object, ModelFieldID, SubformName, pgName)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset, y As Long
    Dim ModelID
    
    y = frm(pgName).Top + 100
    
    ModelID = ELookup("tblModelFields", "ModelFieldID = " & ModelFieldID, "ModelID")
    
    ''Collapsed button so this is a combo box
    ''Create a combo box
    ''Left position should account for the label "Action:"
    ''55 is the space between controls,
    Dim lblWidth, maxX As Long: lblWidth = 1000
    maxX = frm(SubformName).Width - 3100
    
    Dim ctl As control
    Set ctl = CreateControl(frm.Name, acComboBox, , pgName, , maxX, y, 3000, 400)
    ''Set the Default Control Properties Here
    SetControlPropertiesFromTemplate ctl, frm
    ''Additional Property make the RowSource to be the SQLStr, ColumnCount to 2, ColumnWidths to 0;1
    ''Set the Height to be the same height as the buttons
    sqlStr = "SELECT ModelButtonID,ModelButton FROM tblModelButtons WHERE ModelID = " & ModelID & _
             " AND HideOnMain <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
    
    Dim cboName, lblName, btnName As String
    cboName = "cbo" & SubformName & "FormActions": lblName = "lbl" & SubformName & "FormActions"
    btnName = "cmdRun" & SubformName & "FormActions"
    ctl.Name = cboName
    ctl.RowSource = sqlStr
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;1"
    ctl.Height = 400
    ctl.TopMargin = 75
    ctl.LeftMargin = 75
    ctl.FontBold = True
    ctl.HorizontalAnchor = acHorizontalAnchorRight
    
    ''Render the label here
    Set ctl = CreateControl(frm.Name, acLabel, , cboName, , maxX - 55 - lblWidth, y, lblWidth, 400)
    ''Set the Default Control Properties Here
    SetControlPropertiesFromTemplate ctl, frm
    ctl.Name = lblName
    ctl.Caption = "Actions: "
    ctl.TextAlign = 3
    ctl.Height = 400
    ctl.TopMargin = 75
    ctl.LeftMargin = 75
    ctl.FontBold = True
    ctl.HorizontalAnchor = acHorizontalAnchorRight
    
    maxX = frm(cboName).Left + frm(cboName).Width + 55
    RenderButton maxX, y, "Run", 23, frm, btnName, pgName
    
    btnName = "cmd" & btnName
    frm(btnName).Width = frm(btnName).Width / 2
    frm(btnName).HorizontalAnchor = acHorizontalAnchorRight
    
    ''Resize the page width to be that of the subform but with a little but of margin
    frm(pgName).Width = frm(SubformName).Width + 200
    
    frm(btnName).OnClick = "=RunFormActions([Form],[" & cboName & "], " & EscapeString(SubformName) & ")"
    
    
End Function

''Insert the default constructed tblName to the tblSupplementalModels

Public Function frmModelsModelAfterUpdate(frm As Form, Optional TableName = "")
    RunCommandSaveRecord
    
    Dim ModelID As Long
    Dim Model As Variant
    Dim QueryName As Variant
    Dim PrimaryKey As Variant
    Dim VerboseCaption As Variant
    
    ModelID = frm("ModelID")
    Model = frm("Model")
    QueryName = frm("QueryName")
    
    If isFalse(TableName) Then
        TableName = GetTableName(Model)
    End If
    
    PrimaryKey = GetPrimaryKeyFromTable(ModelID)
    frm("PrimaryKey") = PrimaryKey
    
    VerboseCaption = ConvertToVerboseCaption(Model)
    
    Dim VerbosePluralCaption: VerbosePluralCaption = ReplaceMatchedPattern(VerboseCaption, "y$", "ies")
    
    If VerbosePluralCaption = VerboseCaption Then
        VerbosePluralCaption = VerbosePluralCaption & "s"
    End If
    
    frm("VerbosePluralCaption") = VerbosePluralCaption
    frm("VerbosePlural") = RegExReplace(VerbosePluralCaption, "\s+", "")
    
    Dim MainField: MainField = frm("MainField")
    If isFalse(MainField) Then
        frm("MainField") = Model & "ID"
    End If
    
    UpdateSupplementalModelsTable ModelID, TableName, VerboseCaption
    
End Function

Private Sub UpdateSupplementalModelsTable(ModelID As Long, TableName As Variant, VerboseCaption As Variant)
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblSupplementalModels WHERE ModelID = " & ModelID)
    
    If rs.EOF Then
        rs.addNew
        rs.fields("ModelID") = ModelID
    Else
        rs.Edit
    End If
    
    rs.fields("TableName") = TableName
    rs.fields("VerboseCaption") = VerboseCaption
    rs.Update
    
    Dim frm As Form: Set frm = GetForm("frmModels")
    If Not frm Is Nothing Then
        frm("subSupplementalModels").Form.Requery
    End If
    
End Sub
