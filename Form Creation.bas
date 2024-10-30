Attribute VB_Name = "Form Creation"
Option Compare Database
Option Explicit

Public Sub CopyControlProperty(ctl As control, targetCtl As control, propertyName)

    ctl.Properties(propertyName) = targetCtl.Properties(propertyName)
    
End Sub

Public Sub MoveControl(ctl As control, targetCtl As control, Optional toPosition = "Right", Optional Margin = 100)

    Select Case toPosition
        Case "Left":
            ctl.Left = targetCtl.Left - Margin - ctl.Width
        Case "Right":
            ctl.Left = GetRight(targetCtl) + Margin
        Case "Top":
            ctl.Top = targetCtl.Top - Margin - ctl.Height
        Case "Bottom":
            ctl.Top = GetBottom(targetCtl) + Margin
    End Select
    
End Sub

Public Sub AlignControl(ctl As control, targetCtl As control, Optional toPosition = "Left")
    
    Select Case toPosition
        Case "Left":
            ctl.Left = targetCtl.Left
        Case "Right":
            ctl.Left = GetRight(targetCtl) - ctl.Width
        Case "Top":
            ctl.Top = targetCtl.Top
        Case "Bottom":
            ctl.Top = GetBottom(targetCtl) - ctl.Height
    End Select
    
End Sub

Public Sub ResizeForm(frm As Object)
    
    frm.Width = 0
    frm.Width = frm.Width + 400
    
    frm.Section(acDetail).Height = 0
    frm.Section(acDetail).Height = 400
    
End Sub

Public Sub IncreaseButtonHeight(frm As Form)

    Dim ctl As control
    
    Dim i As Integer: i = 0
    For Each ctl In frm.controls
        If ctl.ControlType = acCommandButton Then ctl.Height = ctl.Height * 1.3
        i = i + 1
    Next ctl
            
End Sub

Public Function CreateButtonControl(obj As Object, Caption, ControlName, OnClick, Optional TemplateControlName = "ButtonControl", Optional Section As AcSection = acDetail) As CommandButton

    Dim ctl As control: Set ctl = CreateFlexControl(obj.Name, acCommandButton, Section, , , 0, GetMaxY(obj, Section), obj.Width)
    ctl.Name = ControlName
    ctl.Caption = Caption
    ctl.OnClick = OnClick
    CopyProperties obj, ctl.Name, TemplateControlName
    
    Set CreateButtonControl = ctl
    
End Function


Public Sub RenameFormOrReport(objName, newName)
    
    Dim obj As Object: Set obj = GetFormOrReport(objName, True, True)
    Dim isAReport As Boolean: isAReport = IsObjectAReport(obj)
    If Not obj Is Nothing Then
        isAReport = True
        DoCmd.Close IIf(isAReport, acReport, acForm), obj.Name, acSaveYes
    Else
        Exit Sub
    End If
    
    Set obj = GetFormOrReport(newName)
    If Not obj Is Nothing Then
        DoCmd.Close IIf(isAReport, acReport, acForm), newName, acSaveYes
    End If
    DoCmd.Rename newName, IIf(isAReport, acReport, acForm), objName
    
End Sub

Public Sub SetCommonReportProperties(rpt As Report)
    
    rpt.Printer.PaperSize = acPRPSLetter
    rpt.Printer.TopMargin = InchToTwip(0.25)
    rpt.Printer.BottomMargin = InchToTwip(0.25)
    rpt.Printer.LeftMargin = InchToTwip(0.25)
    rpt.Printer.RightMargin = InchToTwip(0.25)
    
    Dim rptWidth: rptWidth = InchToTwip(8 - 0.5)
    rpt.Width = rptWidth
    
    rpt.Section(acDetail).AlternateBackColor = vbWhite
    rpt.Section(acDetail).BackColor = vbWhite
    
    rpt.Section(acPageHeader).BackColor = vbWhite
    
    rpt.Section(acPageFooter).BackColor = vbWhite
    
    DoCmd.RunCommand acCmdReportHdrFtr
    
    rpt.Section(acHeader).BackColor = vbWhite
    rpt.Section(acFooter).BackColor = vbWhite
    
    rpt.PopUp = True
    rpt.Modal = True
    rpt.AutoCenter = True
    
    rpt.Section(acDetail).OnFormat = "=FormReports_ToggleSubformVisiblity([Report])"
    
End Sub

Public Sub CleanUpReportProperties(rpt As Report)

    rpt.Section(acPageHeader).Height = 0
    rpt.Section(acPageFooter).Height = 0
    rpt.Section(acHeader).Height = 0
    rpt.Section(acFooter).Height = 0
    rpt.Section(acDetail).Height = 0
    rpt.Width = 0
    
End Sub

Public Function CreateLabelControl(obj As Object, Caption, ControlName, Optional TemplateControlName = "LabelControl", Optional Section As AcSection = acDetail) As Label

    Dim ctl As control: Set ctl = CreateFlexControl(obj.Name, acLabel, Section, , , 0, GetMaxY(obj, Section), obj.Width)
    ctl.Name = ControlName
    ctl.Caption = Caption
    CopyProperties obj, ctl.Name, TemplateControlName
    
    Set CreateLabelControl = ctl
    
End Function

Public Function CreateTextboxControl(obj As Object, ControlSource, ControlName, Optional labelCaption = "", Optional TemplateControlName = "TextControl", _
    Optional Section As AcSection = acDetail, Optional Hidden As Boolean = False) As TextBox

    Dim ctl As control
    If Not isFalse(labelCaption) Then
        Set ctl = CreateLabelControl(obj, labelCaption, "lbl" & ControlName, , Section)
    End If
    
    Set ctl = CreateFlexControl(obj.Name, acTextBox, Section, , , 0, GetMaxY(obj, Section), obj.Width)
    ctl.Name = ControlName
    ctl.ControlSource = "=" & ControlSource
    CopyProperties obj, ctl.Name, TemplateControlName
    
    If Hidden Then
        ctl.Top = 0
        ctl.Left = 0
        ctl.Width = 0
        ctl.Visible = False
    End If
    
    Set CreateTextboxControl = ctl
    
End Function

Public Function CreateSubformControl(obj As Object, SourceObject, ControlName, Optional Section As AcSection = acDetail, Optional LinkMasterFields, Optional LinkChildFields) As subform
    
    Dim ctl As control
    Set ctl = CreateFlexControl(obj.Name, acSubform, Section, , , , GetMaxY(obj, Section), obj.Width, InchToTwip(0.5))
    ctl.Name = ControlName
    ctl.SourceObject = SourceObject
    ctl.BorderStyle = 0
    
    If Not isFalse(LinkMasterFields) Then
        ctl.LinkMasterFields = LinkMasterFields
        If isFalse(LinkChildFields) Then
            ctl.LinkChildFields = LinkMasterFields
        End If
    End If
    
    If Not isFalse(LinkChildFields) Then
        ctl.LinkChildFields = LinkChildFields
    End If
    
    ctl.CanGrow = True
    Set CreateSubformControl = ctl
    
End Function

Public Function GetFormOrReport(objName, Optional OpenFormWhenClosed As Boolean = False, Optional AsDesignView As Boolean = False) As Object

    Dim obj As Object: Set obj = GetForm(objName, OpenFormWhenClosed, AsDesignView)
    If obj Is Nothing Then
        Set obj = GetReport(objName, OpenFormWhenClosed, AsDesignView)
    End If
    
    If obj Is Nothing Then
        Exit Function
    End If
    
    Set GetFormOrReport = obj
    
End Function

Public Function CreateFlexControl(objectName, ControlType, Optional Section As AcSection = acDetail, Optional parent, _
    Optional ColumnName, Optional Left, Optional Top, Optional Width, Optional Height) As control
    
    Dim isForm As Boolean: isForm = True
    Dim obj As Object: Set obj = GetForm(objectName)
    If obj Is Nothing Then
        isForm = False
        Set obj = GetReport(objectName)
    End If
    
    If obj Is Nothing Then
        Exit Function
    End If
    
    Dim ctl As control
    If isForm Then
        Set ctl = CreateControl(objectName, ControlType, Section, parent, ColumnName, Left, Top, Width, Height)
    Else
        Set ctl = CreateReportControl(objectName, ControlType, Section, parent, ColumnName, Left, Top, Width, Height)
    End If
    
    Set CreateFlexControl = ctl
    
End Function

Public Sub CreateTotalLabel(frm As Object, ToTheLeftOfThisControlName)
    
    Dim ctl As control
    Set ctl = CreateFlexControl(frm.Name, acLabel, acFooter)
    SetControlPropertiesFromTemplate ctl, frm
    ctl.Name = "lblTotal"
    ctl.Caption = "TOTAL"
    ctl.Width = frm("Sum" & ToTheLeftOfThisControlName).Width
    ctl.Left = frm("Sum" & ToTheLeftOfThisControlName).Left - ctl.Width
    ctl.Top = 0
            
End Sub

Public Sub CreateTotalControl(frm As Object, fieldName)

    Dim ctl As control
            
    Set ctl = CreateFlexControl(frm.Name, acTextBox, acFooter)
    SetControlPropertiesFromTemplate ctl, frm
    ctl.Name = "Sum" & fieldName
    ctl.ControlSource = "=CdblNz(Sum([" & fieldName & "]))"
    ctl.Format = "Standard"
    ctl.Left = frm(fieldName).Left
    ctl.Width = frm(fieldName).Width
    ctl.Top = 0
            
End Sub

Public Sub CreateBannerControls(frm As Object)
    
    Dim ctl As control, bannerCtl As control
    For Each ctl In frm.controls
        If ctl.ControlType = acSubform Then
            Dim SubformName: SubformName = ctl.Name
            Dim BannerCtlName: BannerCtlName = "banner_" & SubformName
            If DoesPropertyExists(frm, BannerCtlName) Then
                Set bannerCtl = frm(BannerCtlName)
            Else
                Set bannerCtl = CreateFlexControl(frm.Name, acTextBox, acDetail)
                bannerCtl.Name = BannerCtlName
            End If
            
            SetControlPropertiesFromTemplate bannerCtl, frm
            
            bannerCtl.Top = ctl.Top
            bannerCtl.Left = ctl.Left
            bannerCtl.Width = ctl.Width
            bannerCtl.Height = ctl.Height
            
            bannerCtl.ControlSource = "=" & Esc("No record to show")
            CopyProperties frm, bannerCtl.Name, "Heading3"
            bannerCtl.ForeColor = vbRed
            
        End If
    Next ctl
    
End Sub

Private Function RemoveDEFormButtons(frm As Form)
    
    Dim buttonArr As New clsArray: buttonArr.arr = "Cancel,New,SaveClose,Delete"
    Dim btn
    For Each btn In buttonArr.arr
        DeleteControl frm.Name, "cmd" & btn
    Next btn
    
End Function

Public Function FormatFormAsReport(frm As Object, FormTypeID, Optional ReferencedControlName = "")
    
    If Not IsObjectAReport(frm) Then
        frm.AllowAdditions = False
        frm.AllowEdits = False
        frm.AllowDeletions = False
    End If

    If FormTypeID = 4 Then
        Dim ctl As control: Set ctl = frm(ReferencedControlName)
        OffsetControlPositions frm, (ctl.Left * -1) + 25, (frm("lbl" & ReferencedControlName).Top * -1) + 25
        frm.ScrollBars = 0
        If Not IsObjectAReport(frm) Then RemoveDEFormButtons frm
    End If
    
    If FormTypeID = 8 Then
        If Not IsObjectAReport(frm) Then
            frm.NavigationButtons = False
            frm.RecordSelectors = False
        End If
        
        OffsetControlPositions frm, 50
    End If
    
    
End Function

Public Function GetFormSuffix(ModelID)
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblModels WHERE ModelID = " & ModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim Model: Model = rs.fields("Model")
    Dim VerbosePlural: VerbosePlural = rs.fields("VerbosePlural")
    Dim SubformName: SubformName = rs.fields("SubformName")
    
    Dim newName: newName = IIf(IsNull(VerbosePlural), Model & "s", VerbosePlural)
    If Not isFalse(SubformName) Then
        newName = SubformName
    End If
    
    GetFormSuffix = newName
    
End Function

Public Sub Offset_ctlPositions(frm As Object, ctl As control, Optional leftOffset As Double = 0, Optional topOffset As Double = 0)
    
    Dim newLeftPosition As Double
    Dim newTopPosition As Double
    
    ' Calculate the new Left position
    newLeftPosition = ctl.Left + leftOffset
    If newLeftPosition > 0 Then
        ctl.Left = newLeftPosition
    End If
    
    ' Calculate the new Top position
    newTopPosition = ctl.Top + topOffset
    If newTopPosition > 0 Then
        ctl.Top = newTopPosition
    End If
        
End Sub

Public Sub OffsetControlPositions(frm As Object, Optional leftOffset As Double = 0, Optional topOffset As Double = 0)
    
    Dim item As control
    Dim newLeftPosition As Double
    Dim newTopPosition As Double
    
    For Each item In frm.controls
        ' Calculate the new Left position
        newLeftPosition = item.Left + leftOffset
        If newLeftPosition > 0 Then
            item.Left = newLeftPosition
        End If
        
        ' Calculate the new Top position
        newTopPosition = item.Top + topOffset
        If newTopPosition > 0 Then
            item.Top = newTopPosition
        End If
        
    Next item
    
End Sub

Public Function GetStandardControlHeight(frm As Form)

    Dim ControlHeight As Double
    Dim ctl As control, items As New clsArray
    For Each ctl In frm.controls
        If ctl.ControlType = acLabel Then
            ControlHeight = ctl.Height
            Exit For
        End If
    Next ctl
    
    GetStandardControlHeight = ControlHeight
    
End Function

Public Sub CreateContinuousFormDeleteButton(frm As Object, ModelID)
    
    Dim maxX: maxX = GetMaxX(frm)
    Dim standardHeight: standardHeight = GetStandardControlHeight(frm)
    
    Dim ctl As control
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , maxX, 0, InchToTwip(0.5), standardHeight)
    ctl.ControlSource = "=""Delete"""
    ctl.Name = "txtDelete"
    ctl.TabStop = False
    CopyProperties frm, ctl.Name, "ReverseTextControlDanger", False
    
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , maxX, 0, InchToTwip(0.5), standardHeight)
    ctl.Name = "cmdDelete"
    CopyProperties frm, ctl.Name, "TransparentButton", False
    ctl.Height = standardHeight
    
    Dim TableName: TableName = GetTableNameFromModelID(ModelID, True)
    Dim PrimaryKey:  PrimaryKey = GetPrimaryKeyFieldFromTable(TableName)
    
    ctl.OnClick = "=DeleteRecord([Form]," & Esc(PrimaryKey) & "," & Esc(TableName) & ")"

End Sub

Public Sub CreateContinuousFormButton(frm As Object, standardHeight, ControlSource, TextBoxName, ButtonName)
    
    Dim maxX: maxX = GetMaxX(frm)
    Dim ctl As control
    
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , maxX, 0, InchToTwip(1), standardHeight)
    ctl.ControlSource = ControlSource
    ctl.Name = TextBoxName
    ctl.TabStop = False
    CopyProperties frm, ctl.Name, "ReverseTextControl", False
    
    Dim cond As FormatCondition
    
    Set cond = ctl.FormatConditions.Add(acExpression, acEqual, "[" & TextBoxName & "] = """"")
    cond.BackColor = vbWhite
    
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , maxX, 0, InchToTwip(1), standardHeight)
    ctl.Name = ButtonName
    'CopyProperties frm, ctl.Name, "TransparentButton", False
    CopyProperties frm, ctl.Name, "TransparentButton", False
    ctl.Height = standardHeight
    
    
End Sub

Public Sub CenterVertically(frmName, ReferenceControl, ControlToCenter)
    
    Dim frm As Form: Set frm = Forms(frmName)
    Dim Top: Top = (frm(ReferenceControl).Height - frm(ControlToCenter).Height) / 2
    frm(ControlToCenter).Top = frm(ReferenceControl).Top + Top
    
End Sub

Function GetLeftPosition(totalWidth, controlWidth)
    ' Calculate the left position
    GetLeftPosition = (totalWidth - controlWidth) / 2
End Function

Public Sub ResizeAssociatedLabel(frmName)
    
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)

    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
            frm("lbl" & ctl.Name).Width = ctl.Width
            frm("lbl" & ctl.Name).Left = ctl.Left
        End If
    Next ctl
    
End Sub

Public Sub GetControlsWithZeroWidth(frmName)
    
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)

    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.Width = 0 And ctl.Visible = True Then
            Debug.Print ctl.Name
            ctl.Visible = False
        End If
    Next ctl

End Sub

Public Function RunFunctionFromSubform(frm As Object, SubformName, FunctionName)
    
    Run FunctionName, frm(SubformName).Form

End Function

'Public Function CreateSetFromMain(frm As Form)
'
'    Dim TableName As String, ViewName As String, DataEntryCaption As String, RsPrefix As String, VerbosePlural As String
'    Dim MainFormCaption As String, SetFocus As String, RequeryOnClose As String
'
'    TableName = "tbl" & frm("TableRecord") & "s"
'    If Not IsNull(frm("PluralForm")) Then
'        TableName = "tbl" & frm("PluralForm")
'    End If
'
'    ''Creation of View Name
'    If frm("isQuery") Then
'        RsPrefix = "qry"
'    Else
'        RsPrefix = "tbl"
'    End If
'
'    ViewName = RsPrefix & frm("TableRecord") & "s"
'
'    If Not IsNull(frm("PluralForm")) Then
'        ViewName = RsPrefix & frm("PluralForm")
'    End If
'
'    ''DataEntryCaption
'    If Not IsNull(frm("ReadableCaption")) Then
'        DataEntryCaption = frm("ReadableCaption") & " Form"
'        MainFormCaption = frm("ReadableCaption") & " List"
'    Else
'        DataEntryCaption = frm("TableRecord") & " Form"
'        MainFormCaption = frm("TableRecord") & " List"
'    End If
'
'    ''Set Focus
'    If IsNull(frm("SetFocus")) Then
'        SetFocus = frm("TableRecord")
'    Else
'        SetFocus = frm("SetFocus")
'    End If
'
'    RequeryOnClose = "main" & frm("TableRecord") & "s.subform"
'
'    If Not IsNull(frm("PluralForm")) Then
'        RequeryOnClose = "main" & frm("PluralForm") & ".subform"
'    End If
'
'    ''Primary Key
'    Dim PrimaryKey As String
'    PrimaryKey = frm("TableRecord") & "ID"
'
'    ''FormName
'    Dim FormName As String
'    FormName = frm("TableRecord") & "s"
'
'    If Not IsNull(frm("PluralForm")) Then
'        FormName = frm("PluralForm")
'    End If
'
'    Dim Fields(8) As String, fieldValues(8) As String
'
'    If IsNull(frm("VerbosePlural")) Then
'        If Not IsNull(frm("ReadableCaption")) Then
'            VerbosePlural = frm("ReadableCaption") & "s"
'        Else
'            VerbosePlural = frm("TableRecord") & "s"
'        End If
'    Else
'        VerbosePlural = frm("VerbosePlural")
'    End If
'
'    Fields(0) = "TableName"
'    Fields(1) = "ViewName"
'    Fields(2) = "DataEntryCaption"
'    Fields(3) = "MainFormCaption"
'    Fields(4) = "SetFocus"
'    Fields(5) = "RequeryOnClose"
'    Fields(6) = "PrimaryKey"
'    Fields(7) = "FormName"
'    Fields(8) = "VerbosePlural"
'
'    fieldValues(0) = "'" & TableName & "'"
'    fieldValues(1) = "'" & ViewName & "'"
'    fieldValues(2) = "'" & DataEntryCaption & "'"
'    fieldValues(3) = "'" & MainFormCaption & "'"
'    fieldValues(4) = "'" & SetFocus & "'"
'    fieldValues(5) = "'" & RequeryOnClose & "'"
'    fieldValues(6) = "'" & PrimaryKey & "'"
'    fieldValues(7) = "'" & FormName & "'"
'    fieldValues(8) = EscapeString(VerbosePlural)
'
'    InsertAndLog "tblTables", Fields, fieldValues
'
'    CreateSet TableName
'
'End Function

'Public Function CreateSet(RecordSource As String)
'    CreateDataEntryForm RecordSource
'    CreateDataSheetForm RecordSource
'    CreateMainForm RecordSource
'End Function

'Public Function CreateNewForm(RecordSource As String, Optional FormType As Integer)
'
'    ''RecordSource => Must be the recordsource in which the form will be based
'    ''Form Type => Can either be 0 = "DataEntry", 1 = "DataSheet", 2 = "MainForm"
'
'    Select Case FormType
'        Case 0:
'            CreateDataEntryForm RecordSource
'        Case 1:
'            CreateDataSheetForm RecordSource
'        Case 2:
'            CreateMainForm RecordSource
'    End Select
'
'End Function

Private Function SetFormProperties(frmType As String, frm As Form)
    ''Set the Form Properties
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblFormProps WHERE FrmType = '" & frmType & "'")
    Do Until rs.EOF
        frm.Properties(rs.fields("FrmPropName")) = rs.fields("FrmPropVal")
        rs.MoveNext
    Loop
End Function

Private Function SetCaption(fld As field)

    Dim ctl As control
    ''Add a filed Caption when it does not exist
    Dim fldCaption
    If DoesPropertyExists(fld.Properties, "Caption") Then
        fldCaption = fld.Properties("Caption")
    Else
        fldCaption = AddSpaces(fld.Name)
        Dim prop As property
        Set prop = fld.CreateProperty("Caption", dbText, fldCaption)
        fld.Properties.Append prop
    End If
    
    SetCaption = fldCaption
    
End Function

Private Function CreateDataSheetForm(recordSource As String)

    ''Get the tblTables properties of the Matched TableName RecordSource
    Dim props As Recordset
    Set props = CurrentDb.OpenRecordset("SELECT * FROM tblTables WHERE TableName = '" & recordSource & "'")
    Dim ViewName As String, PrimaryKey As String
    ViewName = props.fields("viewName")
    PrimaryKey = props.fields("PrimaryKey")
    
    Dim frm As Form
    Set frm = CreateForm
    frm.recordSource = ViewName
    
    frm.Caption = props.fields("MainFormCaption")
    frm.BeforeUpdate = "=SaveFormData([Form],""" & recordSource & """,""" & PrimaryKey & """)"
    frm.OnLoad = "=SetDefaultUserID([Form])"
    
    ''Set the Form Properties
    SetFormProperties "Datasheet", frm
    
    ''Establish the Recordset
    Dim rsDef As Object
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    If DoesPropertyExists(CurrentDb.TableDefs, ViewName) Then
        Set rsDef = dbs.TableDefs(ViewName)
    Else
        Set rsDef = dbs.QueryDefs(ViewName)
    End If
    
    Dim x, y
    'x is the starting left, y is the starting top
    x = 800: y = 600
    
    Dim fld As field
    
    For Each fld In rsDef.fields
    
        If fld.Name = props.fields("PrimaryKey") Then
            GoTo NextField
        End If
        
        ''Create the label first
        Dim ctl As control
        ''Add a filed Caption when it does not exist
        Dim fldCaption
        fldCaption = SetCaption(fld)
        
        Dim ControlTypeID
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeID = acTextBox
        Else
            ControlTypeID = fld.Properties("DisplayControl")
        End If
        
        Set ctl = CreateControl(frm.Name, ControlTypeID, , "", fld.Name, x, y, 3000)
        ctl.Name = fld.Name: ctl.Properties("DatasheetCaption") = fldCaption
        
        ''Set Control Caption
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCtlProps WHERE isNull(CtlPropType)")
        rs.MoveFirst
        
        Do Until rs.EOF
            If DoesPropertyExists(ctl.Properties, rs.fields("CtlPropName")) Then
                ctl.Properties(rs.fields("CtlPropName")) = rs.fields("CtlPropVal")
            End If
            rs.MoveNext
        Loop
        
        y = y + 400
        
        If y > 15000 Then
            y = 600
            x = 2500
        End If
        
        InsertToFields fld, recordSource, fldCaption
        
NextField:
        
    Next fld
    
    frm("Timestamp").ColumnHidden = True
    frm("CreatedBy").ColumnHidden = True
    
    Dim frmName As String, customFrmName As String, i As Integer
    frmName = frm.Name: customFrmName = "dsht" & props.fields("FormName")
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    Do Until Not FrmExist(customFrmName)
        customFrmName = customFrmName & "_1"
    Loop
    
    DoCmd.Rename customFrmName, acForm, frmName
    
End Function

Public Function RenderButton(x As Long, y As Long, Caption As String, QuickStyle As Integer, frm As Form, cmdName As String, Optional parentName = "")
    
    Dim ctl As control
    Set ctl = CreateControl(frm.Name, acCommandButton, , parentName, , x, y, 1250)
    With ctl
        .Name = "cmd" & cmdName
        .Properties("Caption") = Caption
        
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCtlProps WHERE CtlPropType = '104'")
        rs.MoveFirst
        
        Do Until rs.EOF
            If DoesPropertyExists(.Properties, rs.fields("CtlPropName")) Then
                .Properties(rs.fields("CtlPropName")) = rs.fields("CtlPropVal")
            End If
            rs.MoveNext
        Loop
        
        .Properties("QuickStyle") = QuickStyle
        .Properties("UseTheme") = False
        .Properties("CursorOnHover") = 1
    End With
    
End Function


Private Function CreateDataEntryForm(recordSource As String)

    ''Get the tblTables properties of the Matched TableName RecordSource
    Dim props As Recordset
    Set props = CurrentDb.OpenRecordset("SELECT * FROM tblTables WHERE TableName = '" & recordSource & "'")
    Dim ViewName As String, PrimaryKey As String
    ViewName = props.fields("ViewName")
    PrimaryKey = props.fields("PrimaryKey")
    
    Dim frm As Form
    Set frm = CreateForm
    frm.recordSource = ViewName
    frm.Caption = props.fields("DataEntryCaption")
    
    frm.OnCurrent = "=SetFocusOnForm([Form],""" & props.fields("SetFocus") & """)"
    frm.BeforeUpdate = "=SaveFormData([Form],""" & recordSource & """,""" & PrimaryKey & """)"
    frm.OnLoad = "=SetDefaultUserID([Form])"
    
    ''Set the Form Properties
    SetFormProperties "DataEntry", frm
    
    ''Establish the Recordset
    Dim rsDef As Object
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    If DoesPropertyExists(CurrentDb.TableDefs, ViewName) Then
        Set rsDef = dbs.TableDefs(ViewName)
    Else
        Set rsDef = dbs.QueryDefs(ViewName)
    End If
    
    Dim x As Long, y As Long
    'x is the starting left, y is the starting top
    x = 800: y = 600
    
    Dim fld As field
    
    For Each fld In rsDef.fields
        
        
        Select Case fld.Name
            Case props.fields("PrimaryKey"), "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
'        If fld.Name = props.Fields("PrimaryKey") Or fld.Name = "Timestamp" Or fld.Name Then
'            GoTo NextField
'        End If
        
        ''Create the label first
        Dim ctl As control
        ''Add a filed Caption when it does not exist
        Dim fldCaption
        fldCaption = SetCaption(fld)
        
        Set ctl = CreateControl(frm.Name, acLabel, , fld.Name, fldCaption, x, y)
        
        Select Case fld.Name
            Case "RecordTimestamp", "UserID":
                ctl.Visible = False
        End Select
        
        ''Label Properties
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCtlProps WHERE CtlPropType = ""100""")
        
        Do Until rs.EOF
            ctl.Properties(rs.fields("CtlPropName")) = rs.fields("CtlPropVal")
            rs.MoveNext
        Loop
        
        ''Generate the control field
        y = y + 380
        
        Dim ControlTypeID
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeID = acTextBox
        Else
            ControlTypeID = fld.Properties("DisplayControl")
        End If
        
'        If ControlTypeID = 106 Then
'            Set ctl = CreateControl(frm.Name, ControlTypeID, , "", fld.Name, x + 1750, y - 370, 3000)
'            ctl.Name = fld.Name
'        Else
'            Set ctl = CreateControl(frm.Name, ControlTypeID, , "", fld.Name, x, y, 3000)
'            ctl.Name = fld.Name
'        End If

        Set ctl = CreateControl(frm.Name, ControlTypeID, , "", fld.Name, x, y, 3000)
        ctl.Name = fld.Name
        
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCtlProps WHERE isNull(CtlPropType)")
        rs.MoveFirst
        
        
        Do Until rs.EOF
            If DoesPropertyExists(ctl.Properties, rs.fields("CtlPropName")) Then
                ctl.Properties(rs.fields("CtlPropName")) = rs.fields("CtlPropVal")
            End If
            rs.MoveNext
        Loop
        
        y = y + 400
        
        If y > 15000 Then
            y = 600
            x = 2500
        End If
        
        
        InsertToFields fld, recordSource, fldCaption

NextField:
        
        
    Next fld
    
    ''Create the Timestamp and CreatedBy field (Hidden Fields)
    Set ctl = CreateControl(frm.Name, acTextBox, , "", "Timestamp", 0, 0, 0)
    ctl.Name = "Timestamp"
    
    Set ctl = CreateControl(frm.Name, acComboBox, , "", "CreatedBy", 0, 0, 0)
    ctl.Name = "CreatedBy"
    
    y = y + 400
    ''New Record
    RenderButton x, y, "Cancel", 23, frm, "Cancel"
    x = x + 1300
    RenderButton x, y, "New", 23, frm, "New"
    x = x + 1300
    ''Save Record
    RenderButton x, y, "Save", 23, frm, "SaveClose"
    x = x + 1300
    ''Delete Record
    RenderButton x, y, "Delete", 24, frm, "Delete"
    
    frm.cmdCancel.OnClick = "=CancelEdit([Form])"
    frm.cmdNew.OnClick = "=Save([Form],'" & recordSource & "',0)"
    frm.cmdSaveClose.OnClick = "=Save([Form],'" & recordSource & "',1)"
    frm.cmdDelete.OnClick = "=DeleteRecord([Form], '" & props.fields("PrimaryKey") & "', '" & recordSource & "')"
    
    ''Set background color
    ''RGB(31, 73, 125) Green
    ''RGB(31, 73, 125) Blue
    frm.Detail.BackColor = RGB(31, 73, 125)
    
    frm("Timestamp").Visible = False
    frm("CreatedBy").Visible = False
    
    Dim frmName As String, customFrmName As String, i As Integer
    frmName = frm.Name: customFrmName = "frm" & props.fields("FormName")
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    Do Until Not FrmExist(customFrmName)
        customFrmName = customFrmName & "_1"
    Loop
    
    DoCmd.Rename customFrmName, acForm, frmName
    
End Function

Public Function FrmExist(sFrmName As String) As Boolean
    On Error GoTo Error_Handler
    Dim frm                   As Access.AccessObject
 
    For Each frm In Application.CurrentProject.AllForms
        If sFrmName = frm.Name Then
            FrmExist = True
            Exit For    'We know it exist so let leave, no point continuing
        End If
    Next frm
 
Error_Handler_Exit:
    On Error Resume Next
    Set frm = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Error Source: FrmExist" & vbCrLf & _
           "Error Description: " & Err.description, _
           vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function RptExist(sFrmName As String) As Boolean
    On Error GoTo Error_Handler
    Dim frm                   As Access.AccessObject
 
    For Each frm In Application.CurrentProject.AllReports
        If sFrmName = frm.Name Then
            RptExist = True
            Exit For    'We know it exist so let leave, no point continuing
        End If
    Next frm
 
Error_Handler_Exit:
    On Error Resume Next
    Set frm = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Error Source: FrmExist" & vbCrLf & _
           "Error Description: " & Err.description, _
           vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function ImportFieldsToTable(recordSource As String)
    
    ''Establish the Recordset
    Dim rsDef As Object
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    If DoesPropertyExists(CurrentDb.TableDefs, recordSource) Then
        Set rsDef = dbs.TableDefs(recordSource)
    Else
        Set rsDef = dbs.QueryDefs(recordSource)
    End If
    
    Dim fld As field
    
    For Each fld In rsDef.fields
    
        Dim fldCaption
        fldCaption = SetCaption(fld)
    
        InsertToFields fld, recordSource, fldCaption
        
    Next fld
    
End Function

Public Function InsertToFields(fld As field, recordSource As String, fldCaption As Variant)
    If Not isPresent("tblFormFields", "TableName = '" & recordSource & "' And FieldName = '" & fld.Name & "'") Then
        Dim fields(4): Dim fieldValues(4)
        fields(0) = "FieldName"
        fields(1) = "FieldCaption"
        fields(2) = "FieldTypeID"
        fields(3) = "ValidationString"
        fields(4) = "TableName"
        
        fieldValues(0) = """" & fld.Name & """"
        fieldValues(1) = """" & fldCaption & """"
        fieldValues(2) = fld.Type
        
        If fld.required = True Then
            fieldValues(3) = """required"""
        Else
            fieldValues(3) = "Null"
        End If
        
        fieldValues(4) = """" & recordSource & """"
        
        InsertData "tblFormFields", fields, fieldValues
    End If
End Function


