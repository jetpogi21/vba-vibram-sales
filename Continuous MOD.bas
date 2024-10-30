Attribute VB_Name = "Continuous MOD"
Option Compare Database
Option Explicit

''CreateContinuousForm(Forms("frmModels"))
Public Function CreateContinuousForm(frm As Object, Optional RemainedGrouped As Boolean = False, Optional CreateAsReport As Boolean = False)
    
    ''Declare Model Variables
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, SubformName, UserQueryFields, IsSystemTable, PrimaryKey
    
    ModelID = frm("ModelID"): If ExitIfTrue(IsNull(ModelID), "Please select a record..") Then Exit Function
    Model = frm("Model")
    VerboseName = frm("VerboseName")
    VerbosePlural = frm("VerbosePlural")
    MainField = frm("MainField")
    TableWideValidation = frm("TableWideValidation")
    FormColumns = frm("FormColumns")
    SetFocus = frm("SetFocus")
    IsKeyVisible = frm("IsKeyVisible")
    QueryName = frm("QueryName")
    OnFormCreate = frm("OnFormCreate")
    SubformName = frm("SubformName")
    UserQueryFields = frm("UserQueryFields")
    IsSystemTable = frm("IsSystemTable")
    PrimaryKey = frm("PrimaryKey")
    
    Dim ModelName: ModelName = IIf(IsNull(VerboseName), Model, VerboseName)
    
    Dim rsName: rsName = GetTableName(Model, VerbosePlural)
    ''Override the rsName if the QueryName isn't empty
    If Not IsNull(QueryName) Then rsName = QueryName
    Dim tblName: tblName = rsName
    
'    Dim tblName: tblName = IIf(IsNull(VerbosePlural), "tbl" & Model & "s", "tbl" & VerbosePlural)
    If CreateAsReport Then
        Set frm = CreateReport
    Else
        Set frm = CreateForm
    End If
    Dim frmName: frmName = frm.Name
    frm.recordSource = tblName
    frm.DefaultView = 1
    frm.ScrollBars = 2
    
    Dim VerboseCaption: VerboseCaption = ELookup("tblSupplementalModels", "ModelID = " & ModelID, "VerboseCaption")
    Dim frmCaption: frmCaption = GetFieldCaption(VerboseName, Model, VerboseCaption)
    
    frm.Caption = concat(frmCaption, " Form")
    If Not CreateAsReport Then frm.OnCurrent = "=SetFocusOnForm([Form],""" & SetFocus & """)"
    
    ''If the PrimaryKey is null then the PrimaryKey will come from the Model
    If IsNull(PrimaryKey) Then PrimaryKey = GetPrimaryKeyFromTable(ModelID)
    
    If Not CreateAsReport Then
        frm.BeforeUpdate = "=SaveFormData2([Form]," & Esc(Model) & ")"
        frm.OnLoad = "=DefaultFormLoad([Form]," & Esc(PrimaryKey) & ")"
    End If
    
    ''Activate Header in Form
    DoCmd.RunCommand acCmdFormHdrFtr
    frm.Section(acHeader).BackColor = RGB(255, 255, 255)
   
    ''SetFocus, IsKeyVisible, QueryName, PrimaryKey
    RenderFormControls ModelID, frm, RemainedGrouped
    
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, frm, 8
    End If
    
    frm.Section(acFooter).Height = 0
    frm.Section(acFooter).Height = frm.Section(acFooter).Height + 50
    frm.Section(acHeader).Height = 0
    frm.Section(acDetail).Height = 0
    frm.Section(acDetail).AlternateBackColor = vbWhite
    frm.Section(acDetail).BackColor = vbWhite
    frm.Width = 0: frm.Width = frm.Width + 50
    
    If CreateAsReport Then
        frm.Section(acPageHeader).Height = 0
        frm.Section(acPageFooter).Height = 0
        frm.Section(acDetail).Height = 0
        frm.Width = 0
    End If
    
    ''Rename the form here
    Dim origName: origName = frm.Name
    Dim newName: newName = IIf(CreateAsReport, "srpt", "cont") & GetFormSuffix(ModelID)
    DoCmd.Close IIf(CreateAsReport, acReport, acForm), origName, acSaveYes
    DoCmd.Rename newName, IIf(CreateAsReport, acReport, acForm), origName
    LoadSavedFormLayout newName, False
    
    If CreateAsReport Then
        DoCmd.OpenReport newName, acDesign
    Else
        DoCmd.OpenForm newName, acDesign
    End If
    
    
End Function

Private Sub RenderFormControls(ModelID, frm As Object, Optional RemainGrouped As Boolean = False)
    
    Dim rs As Recordset
    Dim ModelFieldID, ModelField, FieldTypeID, VerboseName, ValidationString, ForeignKey, possibleValues, DefaultValue, IsIndexed
    Dim FieldOrder, Columns, ColumnBreak, ColumnWidth, HideSubformFromParent, ParentModelID, VerboseChildName, SubPageOrder, FieldSource, subformSource, IsAnExpression, ControlSource, FieldFormat, ReportFieldOrder

    Set rs = ReturnRecordset("SELECT * FROM qryModelFields WHERE ModelID = " & ModelID & " AND FieldOrder <> 0 ORDER BY FieldOrder")
    Do Until rs.EOF
        RenderFormControl frm, rs
        rs.MoveNext
    Loop
    
    DoCmd.RunCommand acCmdTabularLayout
    DoCmd.RunCommand acCmdControlPaddingNone
    
    Dim ControlHeight
    Dim ctl As control, items As New clsArray
    For Each ctl In frm.controls
        If ctl.ControlType = acLabel Then
            ControlHeight = ctl.Height
            Exit For
        End If
    Next ctl
    
    Dim TextboxReference
    If Not isFalse(ControlHeight) Then
        For Each ctl In frm.controls
            If ctl.ControlType = acTextBox Then
                 ctl.Height = ControlHeight
                 TextboxReference = ctl.Name
                Exit For
            End If
        Next ctl
    End If
    
    If Not RemainGrouped Then DoCmd.RunCommand acCmdRemoveFromLayout
    
    For Each ctl In frm.controls
        If ctl.ControlType = acCheckBox Then
            ctl.Height = InchToTwip(0.15)
            ctl.Width = InchToTwip(0.18)
            ctl.Left = ctl.Left + GetLeftPosition(frm("lbl" & ctl.Name).Width, ctl.Width)
            If Not isFalse(TextboxReference) Then
                CenterVertically frm.Name, TextboxReference, ctl.Name
            End If
            Exit For
        End If
    Next ctl
    
    ''Do the anchoring here
    AnchorControls frm, ModelID
    
End Sub

Public Sub AnchorControls(frm As Object, ModelID)

    Dim rs As Recordset, ctl As control
    Set rs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelID = " & ModelID & " AND FieldOrder <> 0 ORDER BY FieldOrder")
    
    Dim HasHorizontalAnchoredAlready As Boolean: HasHorizontalAnchoredAlready = False
    Dim PreviousControlName: PreviousControlName = ""
    
    Do Until rs.EOF
    
        Dim ColumnWidth: ColumnWidth = rs.fields("ColumnWidth")
        Dim Columns: Columns = rs.fields("Columns")
        Dim ModelField: ModelField = rs.fields("ModelField")
        
        Dim lblName: lblName = "lbl" & ModelField
        Set ctl = frm(lblName)

        ''Setting the column widths here
        ctl.Width = ctl.Width * Columns
        frm(ModelField).Width = frm(ModelField).Width * Columns
        
        If Not IsNull(ColumnWidth) Then
            ctl.HorizontalAnchor = acHorizontalAnchorBoth
            frm(ModelField).HorizontalAnchor = acHorizontalAnchorBoth
            HasHorizontalAnchoredAlready = True
            GoTo NextControl
        End If
        
        If HasHorizontalAnchoredAlready Then
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            frm(ModelField).HorizontalAnchor = acHorizontalAnchorRight
        End If
        
        If Not isFalse(PreviousControlName) Then
            ctl.Left = frm(PreviousControlName).Width + frm(PreviousControlName).Left
            frm(ModelField).Left = ctl.Left
        End If
        
        PreviousControlName = lblName
NextControl:
        rs.MoveNext
    Loop
    
End Sub

''RenderFormControl(Forms("Form1"),
Private Sub RenderFormControl(frm As Object, rs As Recordset)
    
    Dim frmName: frmName = frm.Name
    Dim ModelField: ModelField = rs.fields("ModelField")
    Dim FieldTypeID: FieldTypeID = rs.fields("FieldTypeID")
    Dim ParentModelID: ParentModelID = rs.fields("ParentModelID")
    Dim VerboseName: VerboseName = rs.fields("VerboseName")
    Dim ColumnWidth: ColumnWidth = rs.fields("ColumnWidth")
    Dim possibleValues: possibleValues = rs.fields("PossibleValues")
    
    ''Control Rendering
    ''See if what's to be rendered is a checkbox,textbox or combobox
    Dim ctl As control, ctlName
    If FieldTypeID = 1 Then ''checkbox
        Set ctl = CreateFlexControl(frmName, acCheckBox, acDetail, , ModelField)
        SetCommonControlProperties ctl, frm, rs
        Exit Sub
    End If
    
    If Not IsNull(ParentModelID) Then
        Set ctl = CreateFlexControl(frmName, acComboBox, acDetail, , ModelField)
        SetComboBoxSQL ctl, ParentModelID
        SetCommonControlProperties ctl, frm, rs
        Exit Sub
    End If
    
    If Not IsNull(possibleValues) Then
        Set ctl = CreateFlexControl(frmName, acComboBox, acDetail, , ModelField)
        SetCommonControlProperties ctl, frm, rs
        ctl.ColumnWidths = IIf(ctl.ColumnWidths Like "*cm*", "1cm", "1""")
        ctl.ColumnCount = 1
        ctl.RowSourceType = "Value List"
        ctl.RowSource = QuoteAndJoin(possibleValues)
        ctl.LimitToList = -1
        ctl.AllowValueListEdits = 0
        
        Exit Sub

    End If
    
    ''Textbox here
    Set ctl = CreateFlexControl(frmName, acTextBox, acDetail, , ModelField)
    SetCommonControlProperties ctl, frm, rs
    
   
End Sub


Private Sub SetCommonControlProperties(ByVal ctl As control, frm As Object, rs As Recordset)
    
    Dim ModelField: ModelField = rs.fields("ModelField")
    Dim VerboseName: VerboseName = rs.fields("VerboseName")
    Dim ModelID: ModelID = rs.fields("ModelID")
    Dim TableName: TableName = GetTableNameFromModelID(ModelID)
    Dim FieldCaption: FieldCaption = GetFieldCaption(VerboseName, ModelField)
    Dim FieldTypeEnum: FieldTypeEnum = rs.fields("FieldTypeEnum")
    
    FieldCaption = Coalesce(VerboseName, GetCaptionPropertyFromTable(TableName, ModelField), FieldCaption)
    ctl.Name = ModelField
    SetControlPropertiesFromTemplate ctl, frm
    ctl.InSelection = True
    
    If FieldTypeEnum = "dbDouble" Then
        ctl.Format = "Standard"
    End If
    
    Set ctl = CreateFlexControl(frm.Name, acLabel, acHeader, , FieldCaption)
    ctl.Name = "lbl" & ModelField
    SetControlPropertiesFromTemplate ctl, frm
    ctl.InSelection = True
    ctl.TextAlign = 2
    ''ctl.ForeColor = RGB(0, 32, 96)
    
    
End Sub

Private Sub SetComboBoxSQL(ctl As control, ParentModelID)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ParentModelID)
    
    ''MainField
    Dim MainField: MainField = rs.fields("MainField"): MainField = IIf(Left(MainField, 1) = "=", Right(MainField, Len(MainField) - 1), MainField)
    Dim Model: Model = rs.fields("Model")
    Dim VerbosePlural: VerbosePlural = rs.fields("VerbosePlural")
    Dim tblName: tblName = IIf(IsNull(VerbosePlural), "tbl" & Model & "s", "tbl" & VerbosePlural)
    Dim Pk: Pk = Model & "ID"
    Dim sqlStr
    
    sqlStr = "SELECT " & Pk & "," & MainField & " AS MainField FROM " & tblName & " ORDER BY " & MainField
    ctl.RowSource = sqlStr
On Error GoTo Err_Handler:
    
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0"",1"""
    Exit Sub
    
Err_Handler:
    If Err.Number = 2145 Then
        ctl.ColumnWidths = "0cm;1cm"
        Exit Sub
    End If
    
    MsgBox Err.description, vbCritical
    
    Exit Sub
    
End Sub

'Private Sub RenderFormButton(frm As Object, ByVal x, y, CustomReportID)
'
'    y = y + 100
'
'    Dim ctl As Control
'    Set ctl = createControl(frm.Name, acCommandButton, , , , x, y)
'    ctl.Caption = "Cancel"
'    ctl.OnClick = "=CancelEdit([Form])"
'    SetControlPropertiesFromTemplate ctl, frm
'
'    x = x + ctl.Width + 100
'    Set ctl = createControl(frm.Name, acCommandButton, , , , x, y)
'    ctl.Caption = "Preview Report"
'    ctl.OnClick = "=PreviewCustomReport([Form]," & CustomReportID & ")"
'    SetControlPropertiesFromTemplate ctl, frm
'
'End Sub
'
'Private Sub RenderFormControl(frm As Object, CustomReportField, VerboseName, ByRef x, ByRef y, recordsetName)
'
'    Dim ctl As Control
'    Set ctl = createControl(frm.Name, acLabel, , , VerboseName, x, y)
'    ctl.Name = "lbl" & CustomReportField
'    ctl.Width = 1440 * 3
'    SetControlPropertiesFromTemplate ctl, frm
'    y = y + ctl.height
'
'    Set ctl = createControl(frm.Name, acComboBox, , , , x, y)
'    ctl.Name = CustomReportField
'    ctl.Width = 1440 * 3
'    SetControlPropertiesFromTemplate ctl, frm
'    y = y + ctl.height + 200
'
'    Dim sqlStr
'    sqlStr = "SELECT " & CustomReportField & " FROM " & recordsetName & " GROUP BY " & CustomReportField & " ORDER BY " & CustomReportField
'    ctl.RowSource = sqlStr
'
'End Sub

'Private Sub CreateCustomReportControl(rpt As Report, CustomReportField, FieldTypeID, VerboseName)
'
'    Dim ctl As Control, maxX
'    maxX = GetMaxX(rpt)
'    Set ctl = CreateReportControl(rpt.Name, acTextBox, , , CustomReportField, maxX, 0)
'    ctl.Name = CustomReportField
'    CustomReportControlFormat ctl
'
'    If IsNull(VerboseName) Then VerboseName = AddSpaces(CustomReportField)
'
'    Set ctl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , VerboseName, maxX, 0)
'    ctl.Name = "lbl" & CustomReportField
'    ctl.TextAlign = 2
'    CustomReportControlFormat ctl
'
'End Sub

''DoCmd.RunCommand acCmdTabularLayout
''DoCmd.RunCommand acCmdControlPaddingNone


