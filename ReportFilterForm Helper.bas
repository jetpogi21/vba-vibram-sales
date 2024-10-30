Attribute VB_Name = "ReportFilterForm Helper"
Option Compare Database
Option Explicit
Private Color1, Color2, Color3, Color4

Private Sub SetPrivateVariables()
    Color1 = ELookup("tblApplicationSettings", "ApplicationSettingName = 'Color1'", "ApplicationSettingValue") ''Yellow1
    Color2 = ELookup("tblApplicationSettings", "ApplicationSettingName = 'Color2'", "ApplicationSettingValue") ''Yellow2
    Color3 = ELookup("tblApplicationSettings", "ApplicationSettingName = 'Color3'", "ApplicationSettingValue") ''Blue1
    Color4 = ELookup("tblApplicationSettings", "ApplicationSettingName = 'Color4'", "ApplicationSettingValue") ''Blue2
End Sub


Public Function CustomReportGenerateFilterForm(frm As Form)
    
    Dim CustomReportID, ReportName, ReportObjectName, FilterFormName, recordsetName, PreAppliedFilter, OrderBy, ReportOrientation, PaperSize
    
    CustomReportID = frm("CustomReportID")
    ReportName = frm("ReportName")
    ReportObjectName = frm("ReportObjectName")
    FilterFormName = frm("FilterFormName")
    recordsetName = frm("RecordsetName")
    PreAppliedFilter = frm("PreAppliedFilter")
    OrderBy = frm("OrderBy")
    ReportOrientation = frm("ReportOrientation")
    PaperSize = frm("PaperSize")
    
    SetPrivateVariables
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReportFilterFields WHERE CustomReportID = " & CustomReportID & _
                             " ORDER BY CustomReportFilterFieldID")
    
    
    If rs.EOF Then Exit Function
    
    ''Create The Form
    Dim frm2 As Form, frmName
    Set frm2 = CreateForm
    frmName = frm2.Name
    SetFormProperties 4, frm2
    frm2.Section(acDetail).BackColor = Color4
    frm2.PopUp = -1
    
    Dim x, y, ctl As control
    
    Dim CustomReportFieldID, IsComboBox, rs2 As Recordset
    Dim CustomReportField, FieldTypeID, VerboseName, FieldOrder, FieldProportion

    x = 400: y = 400
    
    Do Until rs.EOF
        
        CustomReportFieldID = rs.fields("CustomReportFieldID")
        IsComboBox = rs.fields("IsComboBox")
        
        Set rs2 = ReturnRecordset("SELECT * FROM tblCustomReportFields WHERE CustomReportFieldID = " & CustomReportFieldID)
        
        CustomReportField = rs2.fields("CustomReportField")
        FieldTypeID = rs2.fields("FieldTypeID")
        VerboseName = rs2.fields("VerboseName")
        FieldOrder = rs2.fields("FieldOrder")
        FieldProportion = rs2.fields("FieldProportion")
        If IsNull(VerboseName) Then VerboseName = AddSpaces(CustomReportField)
        
        RenderFormControl frm2, CustomReportField, VerboseName, x, y, recordsetName
        
        rs.MoveNext
    Loop
    
    RenderFormButton frm2, x, y, CustomReportID
    
    Dim maxX, maxY
    maxX = GetMaxX(frm2)
    maxY = GetMaxY(frm2)
    
    frm2.Width = maxX + 400
    frm2.Section(acDetail).Height = maxY + 400
    frm2.Caption = ReportName
    
    DoCmd.Close acForm, frm2.Name, acSaveYes
    DoCmd.Rename FilterFormName, acForm, frmName
    
End Function

Public Function PreviewCustomReport(frm As Object, CustomReportID, Optional AdditionalFunctionName = "")

    Dim rs As Recordset
    Set rs = ReturnRecordset("SElECT * FROM tblCustomReports WHERE CustomReportID = " & CustomReportID)
    
    Dim ReportName, ReportObjectName, FilterFormName, recordsetName, PreAppliedFilter, OrderBy, ReportOrientation, PaperSize
    
    ReportName = rs.fields("ReportName")
    ReportObjectName = rs.fields("ReportObjectName")
    FilterFormName = rs.fields("FilterFormName")
    recordsetName = rs.fields("RecordsetName")
    PreAppliedFilter = rs.fields("PreAppliedFilter")
    OrderBy = rs.fields("OrderBy")
    ReportOrientation = rs.fields("ReportOrientation")
    PaperSize = rs.fields("PaperSize")
    
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReportFilterFields WHERE CustomReportID = " & CustomReportID & _
                             " ORDER BY CustomReportFilterFieldID")
                             
    Dim CustomReportFieldID, IsComboBox, filterStr, filterArr As New clsArray
    
    Do Until rs.EOF
    
        CustomReportFieldID = rs.fields("CustomReportFieldID")
        IsComboBox = rs.fields("IsComboBox")
        
        filterStr = GetFilterStatement(frm, CustomReportFieldID, recordsetName)
        
        If filterStr = "" Then Exit Function

        filterArr.Add filterStr
        
        rs.MoveNext
    Loop
    
    filterStr = filterArr.JoinArr(" AND ")
    
    If ExitIfTrue(Not isPresent(recordsetName, filterStr), "There is no record to show..") Then Exit Function
    
    DoCmd.OpenReport ReportObjectName, acViewPreview, , filterStr
    
End Function

Private Function GetFilterStatement(frm As Object, CustomReportFieldID, recordsetName) As String
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReportFields WHERE CustomReportFieldID = " & CustomReportFieldID)
    
    Dim CustomReportID, CustomReportField, FieldTypeID, VerboseName, FieldOrder, FieldProportion
    
    CustomReportID = rs.fields("CustomReportID")
    CustomReportField = rs.fields("CustomReportField")
    FieldTypeID = rs.fields("FieldTypeID")
    VerboseName = rs.fields("VerboseName")
    FieldOrder = rs.fields("FieldOrder")
    FieldProportion = rs.fields("FieldProportion")
    If IsNull(VerboseName) Then VerboseName = AddSpaces(CustomReportField)
    
    Dim filterValue
    filterValue = frm(CustomReportField)
    
    If ExitIfTrue(IsNull(filterValue), VerboseName & " is a required field..") Then Exit Function
    
    filterValue = EscapeString(filterValue, recordsetName, CustomReportField)
    
    GetFilterStatement = CustomReportField & " = " & filterValue

End Function


Private Sub RenderFormButton(frm As Object, ByVal x, y, CustomReportID)
    
    y = y + 100
    
    Dim ctl As control
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , x, y)
    ctl.Caption = "Cancel"
    ctl.OnClick = "=CancelEdit([Form])"
    SetControlPropertiesFromTemplate ctl, frm
    
    x = x + ctl.Width + 100
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , x, y)
    ctl.Caption = "Preview Report"
    ctl.OnClick = "=PreviewCustomReport([Form]," & CustomReportID & ")"
    SetControlPropertiesFromTemplate ctl, frm
    
End Sub

Private Sub RenderFormControl(frm As Object, CustomReportField, VerboseName, ByRef x, ByRef y, recordsetName)

    Dim ctl As control
    Set ctl = CreateControl(frm.Name, acLabel, , , VerboseName, x, y)
    ctl.Name = "lbl" & CustomReportField
    ctl.Width = 1440 * 3
    SetControlPropertiesFromTemplate ctl, frm
    y = y + ctl.Height
    
    Set ctl = CreateControl(frm.Name, acComboBox, , , , x, y)
    ctl.Name = CustomReportField
    ctl.Width = 1440 * 3
    SetControlPropertiesFromTemplate ctl, frm
    y = y + ctl.Height + 200
    
    Dim sqlStr
    sqlStr = "SELECT " & CustomReportField & " FROM " & recordsetName & " GROUP BY " & CustomReportField & " ORDER BY " & CustomReportField
    ctl.RowSource = sqlStr
    
End Sub
