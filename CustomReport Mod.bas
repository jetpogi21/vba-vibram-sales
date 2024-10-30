Attribute VB_Name = "CustomReport Mod"
Option Compare Database
Option Explicit
Private Color1, Color2, Color3, Color4

Public Function OpenCustomReportForm(frmName)
    
    DoCmd.OpenForm frmName

End Function

Public Function CustomReportGenerateFields(frm As Form)
    
    Dim CustomReportID, recordsetName, PreAppliedFilter
    CustomReportID = frm("CustomReportID")
    recordsetName = frm("RecordsetName")
    PreAppliedFilter = frm("PreAppliedFilter")
    
    If ExitIfTrue(IsNull(CustomReportID), "Record is empty...") Then Exit Function
    If ExitIfTrue(IsNull(recordsetName), "Recordset is empty...") Then Exit Function
    
    Dim rs As Recordset
    If Not IsNull(PreAppliedFilter) Then recordsetName = concat("SELECT * FROM ", recordsetName, " WHERE ", PreAppliedFilter)
        
    Set rs = ReturnRecordset(recordsetName)
    
    Dim fld As DAO.field
    For Each fld In rs.fields
        Select Case fld.Name
            Case "Timestamp", "CreatedBy", "RecordImportID":
            
            Case Else:
                
                If Not isPresent("tblCustomReportFields", "CustomReportField = " & EscapeString(fld.Name) & _
                             " AND CustomReportID = " & CustomReportID) Then
                    RunSQL "INSERT INTO tblCustomReportFields (CustomReportID, CustomReportField, FieldTypeID, VerboseName) VALUES (" & _
                            CustomReportID & "," & EscapeString(fld.Name) & "," & fld.Type & "," & EscapeString(GetFieldCaption(fld)) & ")"
                End If
                
        End Select
        
    Next fld
    
    frm("subCustomReportFields").Requery
    frm("subCustomReportFilterFields").Requery
    
End Function

Private Function GetFieldCaption(fld As DAO.field) As String
    
    If DoesPropertyExists(fld.Properties, "Caption") Then
        GetFieldCaption = fld.Properties("Caption")
        Exit Function
    End If
    
    GetFieldCaption = AddSpaces(fld.Name)

End Function

Private Function GetReportWidthLessMargin(rpt As Report, ReportOrientation, PaperSize) As Double
    
    ''acPRPSLetter 8.5 x 11
    ''acPRPSLegal 8.5 x 14
    Dim grossWidth
    grossWidth = 8.5
    If ReportOrientation = acPRORLandscape Then
        grossWidth = 11
        If PaperSize = acPRPSLegal Then grossWidth = 14
    End If
    
    rpt.Printer.LeftMargin = 700
    rpt.Printer.RightMargin = 700
    rpt.Printer.TopMargin = 700
    rpt.Printer.BottomMargin = 700
    
    Dim pageMargin
    pageMargin = rpt.Printer.LeftMargin * 2
    GetReportWidthLessMargin = (grossWidth * 1440) - pageMargin - 50
    
End Function

Private Sub SetReportRecordSource(rpt As Report, recordsetName, PreAppliedFilter, OrderBy)
    
    Dim sqlStr
    sqlStr = "SELECT * FROM " & recordsetName
    If Not IsNull(PreAppliedFilter) Then sqlStr = sqlStr & " WHERE " & PreAppliedFilter
    If Not IsNull(OrderBy) Then sqlStr = sqlStr & " ORDER BY " & OrderBy
    
    rpt.recordSource = sqlStr
    
End Sub

Private Sub SetPaperSizeAndOrientation(rpt As Report, ReportOrientation, PaperSize)
    
    rpt.Printer.Orientation = ReportOrientation ''acPROR
    rpt.Printer.PaperSize = PaperSize ''acPRPS
    
End Sub

Public Function CustomReportCreateTabularReport(frm As Form)
    
    Dim CustomReportID, ReportName, ReportObjectName, FilterFormName, recordsetName, PreAppliedFilter, OrderBy
    Dim ReportOrientation, PaperSize
    
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
    
    Dim rpt As Report, rptName
    Set rpt = CreateReport
    rptName = rpt.Name
    rpt.Caption = ReportName
    
    '''Set the PaperSize and Orientation
    SetPaperSizeAndOrientation rpt, ReportOrientation, PaperSize
    
    Dim totalReportWidth
    totalReportWidth = GetReportWidthLessMargin(rpt, ReportOrientation, PaperSize)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReportFields WHERE CustomReportID = " & CustomReportID & _
                             " AND FieldOrder <> 0 ORDER BY FieldOrder ASC, CustomReportFieldID ASC")
    
    ''Close If End of File
    If rs.EOF Then
        DoCmd.Close acReport, rpt.Name, acSaveNo
        rs.Close
        ShowError "No custom report field found.."
        Exit Function
    End If
    
    
    SetReportRecordSource rpt, recordsetName, PreAppliedFilter, OrderBy
    
    Dim CustomReportField, FieldTypeID, VerboseName, FieldOrder, FieldProportion
    Dim totalProportion As Double
    
    Do Until rs.EOF
    
        CustomReportField = rs.fields("CustomReportField")
        FieldTypeID = rs.fields("FieldTypeID")
        VerboseName = rs.fields("VerboseName")
        FieldOrder = rs.fields("FieldOrder")
        FieldProportion = rs.fields("FieldProportion")

        CreateCustomReportControl rpt, CustomReportField, FieldTypeID, VerboseName
        totalProportion = totalProportion + FieldProportion
        
        rs.MoveNext
        
    Loop
    
    DoCmd.RunCommand acCmdTabularLayout
    DoCmd.RunCommand acCmdControlPaddingNone
    
    AdjustReportSizes rpt, rs, totalProportion, totalReportWidth
    
    GroupCustomReport rpt, CustomReportID
    
    RenderReportHeading rpt, CustomReportID
    
    RenderCurrentDateAndPages rpt, totalReportWidth
    
    DoCmd.Close acReport, rpt.Name, acSaveYes
    DoCmd.Rename ReportObjectName, acReport, rptName
    DoCmd.OpenReport ReportObjectName, acViewPreview
     
End Function

Private Sub RenderCurrentDateAndPages(rpt As Report, totalReportWidth)

    Dim ctl As control
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageHeader)
    SetControlPropertiesFromTemplate ctl, rpt
    ctl.Width = 1440 * 2
    ctl.Left = totalReportWidth - ctl.Width
    ctl.BorderStyle = 0
    ctl.ControlSource = "=Now()"
    
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageFooter, , , 0, 0)
    SetControlPropertiesFromTemplate ctl, rpt
    ctl.Width = 1440 * 2
    ctl.ControlSource = "=""Page "" & [Page] & "" of "" & [Pages]"
    ctl.BorderStyle = 0
    
    
End Sub

Private Sub RenderReportHeading(rpt As Report, CustomReportID)

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReportHeadings WHERE CustomReportID = " & CustomReportID & _
                             " ORDER BY CustomReportHeadingOrder ASC")
                             
    Dim CustomReportHeading, fontSize, ctl As control, y
    fontSize = 12
    Do Until rs.EOF
        
        y = GetMaxY(rpt, acPageHeader)
        
        CustomReportHeading = rs.fields("CustomReportHeading")
        Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageHeader, , , , y)
        SetControlPropertiesFromTemplate ctl, rpt
        ctl.Width = 1440 * 5
        ctl.fontSize = fontSize
        ctl.Height = ctl.Height + 100
        ctl.FontBold = True
        ctl.ControlSource = "=" & CustomReportHeading
        ctl.BorderStyle = 0
        fontSize = fontSize * 0.8
        rs.MoveNext
    Loop

End Sub

Private Sub GroupCustomReport(rpt As Report, CustomReportID)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReportGroups WHERE CustomReportID = " & CustomReportID & " ORDER BY GroupOrder")
    
    If rs.EOF Then Exit Sub
    
    Dim GroupOrder, CustomReportFieldID, i As Integer, rs2 As Recordset, groupHeaderIndex
    
    Do Until rs.EOF
        
        GroupOrder = rs.fields("GroupOrder")
        CustomReportFieldID = rs.fields("CustomReportFieldID")
        groupHeaderIndex = (i * 2) + 5
        
        Set rs2 = ReturnRecordset("SELECT * FROM tblCustomReportFields WHERE CustomReportFieldID = " & CustomReportFieldID)
        
        Dim CustomReportField, FieldTypeID, VerboseName, FieldOrder, FieldProportion
        
        CustomReportField = rs2.fields("CustomReportField")
        FieldTypeID = rs2.fields("FieldTypeID")
        VerboseName = rs2.fields("VerboseName")
        FieldOrder = rs2.fields("FieldOrder")
        FieldProportion = rs2.fields("FieldProportion")
        If IsNull(VerboseName) Then VerboseName = AddSpaces(CustomReportField)
        
        CreateGroupLevel rpt.Name, CustomReportField, -1, -1
        rpt.GroupLevel(i).KeepTogether = 1
        
        ''Move all the labels down
        MoveAllLabelsDown rpt
        
         ''Adjust the Height of the Group
        AdjustGroupHeight groupHeaderIndex, rpt
        
        i = i + 1
        rs.MoveNext
    Loop
    
End Sub

Private Sub AdjustGroupHeight(groupHeaderIndex, rpt As Report)
    
    With rpt.Section(groupHeaderIndex)
        .Height = 0
        .KeepTogether = True
        .AlternateBackColor = RGB(254, 254, 254)
    End With
    
    With rpt.Section(groupHeaderIndex + 1)
        .Height = 0
        .AlternateBackColor = RGB(254, 254, 254)
    End With
    
End Sub

Private Sub MoveAllLabelsDown(rpt As Report)

    ''Select all the controls with lbl as beginning
    Dim ctl As control
    For Each ctl In rpt.controls
        ctl.InSelection = False
        If ctl.Name Like "lbl*" Then ctl.InSelection = True
    Next ctl
    
    DoCmd.RunCommand acCmdMoveColumnCellDown
    
    For Each ctl In rpt.controls
        ctl.InSelection = True
    Next ctl
    
    DoCmd.RunCommand acCmdAlignToTallest
    DoCmd.RunCommand acCmdControlPaddingNone

    
End Sub

Private Sub SetPrivateVariables()
    Color1 = ELookup("tblApplicationSettings", "ApplicationSettingName = 'Color1'", "ApplicationSettingValue") ''Yellow1
    Color2 = ELookup("tblApplicationSettings", "ApplicationSettingName = 'Color2'", "ApplicationSettingValue") ''Yellow2
    Color3 = ELookup("tblApplicationSettings", "ApplicationSettingName = 'Color3'", "ApplicationSettingValue") ''Blue1
    Color4 = ELookup("tblApplicationSettings", "ApplicationSettingName = 'Color4'", "ApplicationSettingValue") ''Blue2
End Sub

Private Sub AdjustReportSizes(rpt As Report, rs As Recordset, totalProportion, totalReportWidth)
    
    rpt.Section(acPageHeader).Height = 0
    rpt.Section(acDetail).Height = 0
    
    rs.MoveFirst
    
    Dim CustomReportField, FieldTypeID, VerboseName, FieldOrder, FieldProportion
    
    Do Until rs.EOF
    
        CustomReportField = rs.fields("CustomReportField")
        FieldTypeID = rs.fields("FieldTypeID")
        VerboseName = rs.fields("VerboseName")
        FieldOrder = rs.fields("FieldOrder")
        FieldProportion = rs.fields("FieldProportion")
        
        rpt.controls(CustomReportField).Width = FieldProportion / totalProportion * totalReportWidth
        rs.MoveNext
    Loop
    
    rpt.Width = totalReportWidth

End Sub


Private Sub CustomReportControlFormat(ctl As control)
    
    SetControlProperties ctl
    ctl.InSelection = True
    
    If ctl.ControlType = acLabel Then
        ctl.BackStyle = 1
        ctl.BackColor = CLng(Color1)
        ctl.ForeColor = CLng(Color3)
        ctl.BorderStyle = 1
        ctl.TextAlign = 2
    Else
        
        ctl.CanGrow = True
        ctl.BorderStyle = 0
        
    End If
    
End Sub


Private Sub CreateCustomReportControl(rpt As Report, CustomReportField, FieldTypeID, VerboseName)
    
    Dim ctl As control, maxX
    maxX = GetMaxX(rpt)
    Set ctl = CreateReportControl(rpt.Name, acTextBox, , , CustomReportField, maxX, 0)
    ctl.Name = CustomReportField
    CustomReportControlFormat ctl
    
    If IsNull(VerboseName) Then VerboseName = AddSpaces(CustomReportField)
    
    Set ctl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , VerboseName, maxX, 0)
    ctl.Name = "lbl" & CustomReportField
    ctl.TextAlign = 2
    CustomReportControlFormat ctl
    
    
End Sub


