Attribute VB_Name = "Model Mod"
Option Compare Database
Option Explicit

Public Function CreateFormSet(frm2 As Form)
    
    CreateDEForm frm2
    CreateDSForm frm2
    CreateMainForm frm2
    CreateSimpleDEForm frm2
    AddMainFormTo_tblMainMenus frm2
    
    Dim ModelID: ModelID = frm2("ModelID")
    Dim FormSuffix: FormSuffix = GetFormSuffix(ModelID)
    DoCmd.Close acForm, "dsht" & FormSuffix, acSaveNo
    
End Function

Public Function AddMainFormTo_tblMainMenus(frm As Form)
    
    Dim ModelID: ModelID = frm("ModelID")
    Dim Model: Model = frm("Model")
    Dim VerbosePluralName: VerbosePluralName = GetVerbosePluralName(Model)
    Dim FormSuffix: FormSuffix = GetFormSuffix(ModelID)
    
    Dim MenuOrder: MenuOrder = ELookup("tblMainMenus", "MainMenuID > 0", "MenuOrder", "MenuOrder DESC")
    
    If isFalse(MenuOrder) Then
        MenuOrder = 1
    Else
        MenuOrder = CDbl(MenuOrder) + 1
    End If
    
    Dim fields As New clsArray: fields.arr = "MenuCaption,FormName,MenuOrder"
    Dim fieldValues As New clsArray
    Set fieldValues = New clsArray
    fieldValues.Add VerbosePluralName
    fieldValues.Add "main" & FormSuffix
    fieldValues.Add MenuOrder
    
    UpsertRecord "tblMainMenus", fields, fieldValues, "MenuCaption = " & Esc(VerbosePluralName)

End Function

Public Function frmModels_OnLoad(frm As Form)
    
    DefaultFormLoad frm, "ModelID"
    ''frm("subModelFields").Form("FieldSource").SetFocus
    ''DoCmd.RunCommand acCmdFreezeColumn
    Dim FieldTypeID: FieldTypeID = ELookup("tblFieldTypes", "FieldTypeEnum = " & Esc("dbText"), "FieldTypeID")
    frm("subModelFields").Form.FieldTypeID.DefaultValue = "=" & FieldTypeID
    
End Function

Public Function CreateCustomModule(frm2 As Form)
    
    Dim VBProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim lineNum As Long: lineNum = 4
    
    Set VBProj = Application.VBE.VBProjects(Application.GetOption("Project Name"))
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, SubformName, Timestamp, CreatedBy, PrimaryKey
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    SubformName = frm2("SubformName")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    PrimaryKey = frm2("PrimaryKey")
    
    Dim moduleName: moduleName = concat(Model, " Mod")
    
    InsertToModelRelatedObjects ModelID, acModule, moduleName
    
    If DoesPropertyExists(VBProj.VBComponents, moduleName) Then
        Set vbComp = VBProj.VBComponents(moduleName)
    Else
    
        Set vbComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
        vbComp.Name = moduleName
        
        Set CodeMod = vbComp.CodeModule
        
        InsertLines CodeMod, lineNum, concat("Public Function ", Model, "Create(frm AS Object, FormTypeID)")
        InsertLines CodeMod, lineNum, ""
        
        InsertLines CodeMod, lineNum, vbTab & "Select Case FormTypeID"
        InsertLines CodeMod, lineNum, vbTab & vbTab & "Case 4: ''Data Entry Form"
        InsertLines CodeMod, lineNum, vbTab & vbTab & "Case 5: ''Datasheet Form"
        InsertLines CodeMod, lineNum, vbTab & vbTab & "Case 6: ''Main Form"
        InsertLines CodeMod, lineNum, vbTab & vbTab & "Case 7: ''Tabular Report"
        InsertLines CodeMod, lineNum, vbTab & vbTab & "Case 8: ''Cont Form"
        InsertLines CodeMod, lineNum, vbTab & vbTab & "Case 9: ''Selector Form"
        InsertLines CodeMod, lineNum, vbTab & vbTab & vbTab & "Dim contFrm As Form: Set contFrm = frm(""subform"").Form"
        InsertLines CodeMod, lineNum, vbTab & "End Select"
        
        InsertLines CodeMod, lineNum, ""
        
        InsertLines CodeMod, lineNum, "End Function"
        
        
    End If
    
    frm2("OnFormCreate") = concat(Model, "Create")
    
    DoCmd.Save acModule, moduleName
    
End Function

Private Sub InsertLines(CodeMod As Object, lineNum As Long, CodeStr)
    
    CodeMod.InsertLines lineNum, CodeStr
    lineNum = lineNum + 1
    
End Sub

Public Function GenerateDatasheetControls(frm As Form)
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, SubformName, UserQueryFields, IsSystemTable
    ModelID = frm("ModelID")
    
    If ExitIfTrue(IsNull(ModelID), "Please select a record.") Then Exit Function
    
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
    
    Dim dshtName
    If Not IsNull(SubformName) Then
        dshtName = concat("dsht", SubformName)
    Else
        If Not IsNull(VerbosePlural) Then
            dshtName = concat("dsht", VerbosePlural)
        Else
            dshtName = concat("dsht", Model, "s")
        End If
    End If
    
    DoCmd.OpenForm dshtName, acDesign
    
    Dim dshtFrm As Form, ctl As control, ControlName, ControlCaption, ControlOrder As Integer
    Set dshtFrm = Forms(dshtName)
    
    ControlOrder = 1
    For Each ctl In dshtFrm.Section(acFooter).controls
        ControlName = ctl.Name
        ControlCaption = AddSpaces(replace(ControlName, "Sum", ""))
        If Not isPresent("tblDatasheetTotals", "controlName = " & EscapeString(ControlName) & " AND ModelID = " & ModelID) Then
            RunSQL "INSERT INTO tblDatasheetTotals (controlName,ControlCaption,ControlOrder,ModelID) VALUES (" & _
                    EscapeString(ControlName) & "," & _
                    EscapeString(ControlCaption) & "," & _
                    ControlOrder & "," & _
                    ModelID & ")"
        End If
        ControlOrder = ControlOrder + 1
    Next ctl
    
    DoCmd.Close acForm, dshtFrm.Name, acSaveNo
    
    If DoesPropertyExists(frm, "subDatasheetTotals") Then
        frm("subDatasheetTotals").Form.Requery
    End If
    
    MsgBox "Datasheet Control Successfully Imported.."
    
End Function

Private Function OverrideProperties(ModelID, FormTypeID, frm As Object)
    
    Dim filterArr As New clsArray
    filterArr.Add "FormTypeID = " & FormTypeID
    If FormTypeID = 4 Or FormTypeID = 5 Then
        filterArr.Add "FormTypeID = 8"
    End If
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyOverrides WHERE ModelID = " & ModelID & " And (" & filterArr.JoinArr(" OR ") & ")")
    
    Dim PropertyOverrideID, ControlName, propertyName, PropertyValue
    
    Do Until rs.EOF
        PropertyOverrideID = rs.fields("PropertyOverrideID")
        ControlName = rs.fields("ControlName")
        propertyName = rs.fields("PropertyName")
        PropertyValue = rs.fields("PropertyValue")
        If Not IsNull(ControlName) Then
            If DoesPropertyExists(frm, ControlName) Then
                If DoesPropertyExists(frm(ControlName).Properties, propertyName) Then
                    frm(ControlName).Properties(propertyName) = PropertyValue
                End If
            End If
        Else
            If DoesPropertyExists(frm.Properties, propertyName) Then
                frm.Properties(propertyName) = PropertyValue
            End If
        End If
        rs.MoveNext
    Loop
    
End Function

Public Function CreateTabularReport(frm2 As Form)
        
    Dim rpt As Report, rs As Recordset, rsName, frmCaption, PrimaryKey, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth
    Dim ctl As control
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible
    Dim QueryName, OnFormCreate, SubformName, UserQueryFields, Timestamp, CreatedBy
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    SubformName = frm2("SubformName")
    UserQueryFields = frm2("UserQueryFields")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    '''Create the controls
    '''Fields, default buttons, additional buttons and controls
    
    ''Create the form
    Set rpt = CreateReport
    rsName = GetTableName(Model, VerbosePlural)
    
    Dim tblName
    tblName = rsName
    
    If Not IsNull(QueryName) Then
        rsName = QueryName
    End If
    
    rpt.recordSource = rsName
    frmCaption = GetFieldCaption(VerboseName, Model)
    rpt.Caption = concat(frmCaption, " List")
    rpt.PopUp = True
    rpt.AutoCenter = True


    DoCmd.RunCommand acCmdReportHdrFtr
    
    Dim rsObj As Object, db As DAO.Database, fld As DAO.field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    x = 0
    y = 0
    
    CurrentCol = 0: isMemo = False
    
    Dim sqlStr As String
    sqlStr = "SELECT * FROM tblModelFields WHERE ReportFieldOrder IS NOT NULL AND ModelID = " & ModelID
    
    If Not UserQueryFields Then
        sqlStr = sqlStr & " AND FieldSource = " & EscapeString(rsName)
    End If
    
    Set rs = ReturnRecordset(sqlStr & " ORDER BY ReportFieldOrder ASC")
    
    Do Until rs.EOF
    
        ''Look First if the ControlSource is not Null
        If Not IsNull(rs.fields("ControlSource")) Then
           If isPresent("qryModelFieldProperties", "Property = ""onDatasheetOnly"" And ModelFieldID = " & rs.fields("ModelFieldID")) Then
               
                GoTo NextField
            Else
                ''Create a control at the footer of the form
                fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
                Set ctl = CreateReportControl(rpt.Name, acTextBox, , "", "", x, y, fldWidth)
                SetControlPropertiesFromTemplate ctl, rpt
                ctl.Name = rs.fields("ModelField")
                ctl.ControlSource = rs.fields("ControlSource")
                
                Select Case rs.fields("FieldTypeID")
                    Case dbMemo:
                        ctl.Height = 900
                        isMemo = True
                    Case dbDouble
                        ctl.Format = "Standard"
                End Select
            
                ''Generate the label just above the control
                Dim CustomVerboseName
                If IsNull(rs.fields("VerboseName")) Then
                    CustomVerboseName = AddSpaces(rs.fields("ModelField"))
                Else
                    CustomVerboseName = rs.fields("VerboseName")
                End If
                Set ctl = CreateReportControl(rpt.Name, acLabel, , rs.fields("ModelField"), CustomVerboseName, x, y - 300)
                SetControlPropertiesFromTemplate ctl, rpt
                ctl.Width = fldWidth
                
                GoTo SetVariables
            End If
            
        End If
        
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"))
        
        If Not DoesPropertyExists(rsObj.fields, fldName) Then
            GoTo NextField
        End If
        
        Set fld = rsObj.fields(fldName)
        
        If Not IsKeyVisible And fld.Name = PrimaryKey Then
            GoTo NextField
        End If
        
        Select Case fld.Name
            Case "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
        Dim ControlTypeValue
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeValue = acTextBox
        Else
            ControlTypeValue = fld.Properties("DisplayControl")
        End If
        
        ''Generate the control first before the label
        ''Get the width depending on the Number of columns but make sure that the CurrentCol + the columns will not exceed
        ''the FormColumn
        If fld.Type = dbMemo Then
            fldWidth = 2500
        Else
            fldWidth = 1200
        End If
        
        ''Setting the FieldOrder to 0 here will make the fld hidden..
        If rs.fields("FieldOrder") = 0 Then
'            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, 0, 0, 0)
'            ctl.Name = fld.Name
'            SetControlPropertiesFromTemplate ctl, frm
            GoTo NextField
        End If
        
        If fld.Type = dbBoolean Then
            Set ctl = CreateReportControl(rpt.Name, acTextBox, , "", fld.Name, x, y, fldWidth)
            SetControlPropertiesFromTemplate ctl, rpt
            ctl.FontName = "Wingdings"
            ctl.Format = "?;\?"
            ctl.fontSize = 12
            ctl.TextAlign = 2
        Else
            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, x, y, fldWidth)
            SetControlPropertiesFromTemplate ctl, rpt
        End If
        
        ctl.BottomMargin = 50
        ctl.LeftMargin = 50
        ctl.RightMargin = 50
        ctl.TopMargin = 50
        ctl.Name = fld.Name
        ctl.BackStyle = 0
        
        ''Set control property based on ControlTypeValue
        
        'ctl.BorderStyle = 0
    
        Select Case fld.Type
             Case dbMemo:
                 'ctl.height = 900
                 isMemo = True
             Case dbDouble:
                 ctl.Format = "Standard"
         End Select
         
         ctl.CanGrow = True
         ctl.InSelection = True
        
        ''Generate the label just above the control
        Dim ControlCaption
        If Not DoesPropertyExists(fld.Properties, "Caption") Then
            If IsNull(rs.fields("VerboseName")) Then
                ControlCaption = AddSpaces(rs.fields("ModelField"))
            Else
                ControlCaption = rs.fields("VerboseName")
            End If
        Else
            ControlCaption = fld.Properties("Caption")
        End If
        
        Set ctl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , ControlCaption, x, y)

        SetControlPropertiesFromTemplate ctl, rpt
        ctl.Width = fldWidth
        ctl.TextAlign = 2
        'ctl.Height = 200
        ctl.InSelection = True
        ctl.BackStyle = 1
        ctl.BackColor = 49407
        ctl.ForeColor = RGB(31, 73, 125)
        ctl.Top = 2000
        ctl.BorderStyle = 1

SetVariables:
    
        CurrentCol = CurrentCol + rs.fields("Columns")
        
        x = x + fldWidth
                 
NextField:
        
        rs.MoveNext
        
    Loop
    
    DoCmd.RunCommand acCmdTabularLayout
    DoCmd.RunCommand acCmdControlPaddingNone
    
    For Each ctl In rpt.controls
        If ctl.InSelection Then
            ctl.Left = 0
            Exit For
        End If
    Next ctl
    
    Dim rptWidth
    rptWidth = (8.5 - (0.25 * 2)) * 1440
    
    ''Write the pageheader (The Report Title & Current Date Time)
    Set ctl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , , 0, 0, 4800, 400)
    SetControlPropertiesFromTemplate ctl, rpt
    ctl.fontSize = 12
    ctl.Height = 340
    ctl.Caption = AddSpaces(GetModelPlural(Model, VerbosePlural, ""))
    ctl.ForeColor = RGB(31, 73, 125)
    
    ''Current Date Time
    fldWidth = 3000
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageHeader, , , rptWidth - fldWidth, 0, fldWidth, 400)
    SetControlPropertiesFromTemplate ctl, rpt
    ctl.ControlSource = "=Now()"
    ctl.BorderStyle = 0
    
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageHeader, , , 0, ctl.Top + ctl.Height + 200, rptWidth, 570)
    SetControlPropertiesFromTemplate ctl, rpt
    ctl.fontSize = 12
    ctl.Height = 570
    ctl.ForeColor = RGB(254, 254, 254)
    ctl.BackColor = RGB(31, 73, 125)
    ctl.ControlSource = "=" & EscapeString("SOME CAPTION")
    ctl.TextAlign = 2
    ctl.FontBold = True
    ctl.TopMargin = 100
    
    ''At pagefooter (The Page x of y)
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageFooter, , , 0, 0, fldWidth, 400)
    SetControlPropertiesFromTemplate ctl, rpt
    ctl.ControlSource = "=""Page "" & [Page] & "" of "" & [Pages]"
    ctl.BorderStyle = 0
    
    rpt.Section(acHeader).Height = 0
    rpt.Section(acHeader).BackColor = RGB(254, 254, 254)
    rpt.Section(acFooter).Height = 0
    rpt.Section(acPageFooter).Height = 0
    rpt.Section(acPageHeader).Height = 0
    rpt.Section(acDetail).Height = 0
    rpt.Section(acDetail).AlternateBackColor = RGB(250, 243, 232)
    rpt.Section(acDetail).BackColor = RGB(254, 254, 254)
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    frmName = rpt.Name
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("rptTab", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("rptTab", Model, "s")
    End If
    
    If Not IsNull(SubformName) Then
        baseFormName = concat("rptTab", SubformName)
    End If
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, rpt, 7
    End If
    
    ''Override
    OverrideProperties ModelID, 7, rpt
    
    DoCmd.Close acReport, rpt.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not RptExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    
    ''Insert the newly created form to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acReport, customFrmName
    
    DoCmd.Rename customFrmName, acReport, frmName
    DoCmd.OpenReport customFrmName, acViewPreview
    
End Function

Public Function CreateColumnReport(frm2 As Form)
    
    Dim rpt As Report, rs As Recordset, rsName, frmCaption, PrimaryKey, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth
    Dim ctl As control
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible
    Dim QueryName, OnFormCreate, SubformName, UserQueryFields, Timestamp, CreatedBy
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    SubformName = frm2("SubformName")
    UserQueryFields = frm2("UserQueryFields")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    '''Create the controls
    '''Fields, default buttons, additional buttons and controls
    
    ''Create the form
    Set rpt = CreateReport
    rsName = GetTableName(Model, VerbosePlural)
    
    Dim tblName
    tblName = rsName
    
    If Not IsNull(QueryName) Then
        rsName = QueryName
    End If
    
    rpt.recordSource = rsName
    frmCaption = GetFieldCaption(VerboseName, Model)
    rpt.Caption = concat(frmCaption, " Report")
    rpt.PopUp = True

    DoCmd.RunCommand acCmdReportHdrFtr
    
    Dim rsObj As Object, db As DAO.Database, fld As DAO.field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    x = 0
    y = 0
    
    CurrentCol = 0: isMemo = False
    
    Dim sqlStr As String
    sqlStr = "SELECT * FROM tblModelFields WHERE ReportFieldOrder IS NOT NULL AND ModelID = " & ModelID
    
    If Not UserQueryFields Then
        sqlStr = sqlStr & " AND FieldSource = " & EscapeString(rsName)
    End If
    
    Set rs = ReturnRecordset(sqlStr & " ORDER BY ReportFieldOrder ASC")
    
    Dim maxY
    
    Do Until rs.EOF
        
        maxY = GetMaxY(rpt) + 200
        ''Look First if the ControlSource is not Null
        If Not IsNull(rs.fields("ControlSource")) Then
           If isPresent("qryModelFieldProperties", "Property = ""onDatasheetOnly"" And ModelFieldID = " & rs.fields("ModelFieldID")) Then
               
                GoTo NextField
            Else
                ''Create a control at the footer of the form
                fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
                Set ctl = CreateReportControl(rpt.Name, acTextBox, , "", "", x, y, fldWidth)
                SetControlPropertiesFromTemplate ctl, rpt
                ctl.Name = rs.fields("ModelField")
                ctl.ControlSource = rs.fields("ControlSource")
                
                Select Case rs.fields("FieldTypeID")
                    Case dbMemo:
                        ctl.Height = 900
                        isMemo = True
                    Case dbDouble
                        ctl.Format = "Standard"
                End Select
            
                ''Generate the label just above the control
                Dim CustomVerboseName
                If IsNull(rs.fields("VerboseName")) Then
                    CustomVerboseName = AddSpaces(rs.fields("ModelField"))
                Else
                    CustomVerboseName = rs.fields("VerboseName")
                End If
                Set ctl = CreateReportControl(rpt.Name, acLabel, , rs.fields("ModelField"), CustomVerboseName, x, y - 300)
                SetControlPropertiesFromTemplate ctl, rpt
                ctl.Width = fldWidth
                
                GoTo SetVariables
            End If
            
        End If
        
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"))
        
        If Not DoesPropertyExists(rsObj.fields, fldName) Then
            GoTo NextField
        End If
        
        Set fld = rsObj.fields(fldName)
        
        If Not IsKeyVisible And fld.Name = PrimaryKey Then
            GoTo NextField
        End If
        
        Select Case fld.Name
            Case "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
        Dim ControlTypeValue
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeValue = acTextBox
        Else
            ControlTypeValue = fld.Properties("DisplayControl")
        End If
        
        ''Generate the control first before the label
        ''Get the width depending on the Number of columns but make sure that the CurrentCol + the columns will not exceed
        ''the FormColumn
        If fld.Type = dbMemo Then
            fldWidth = 2500
        Else
            fldWidth = 1200
        End If
        
        ''Setting the FieldOrder to 0 here will make the fld hidden..
        If rs.fields("FieldOrder") = 0 Then
'            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, 0, 0, 0)
'            ctl.Name = fld.Name
'            SetControlPropertiesFromTemplate ctl, frm
            GoTo NextField
        End If
        
        If fld.Type = dbBoolean Then
            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, x + 550, y, 200)
        Else
            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, x + fldWidth + 50, maxY, fldWidth)
        End If
        ctl.Name = fld.Name
        
        ''Set control property based on ControlTypeValue
        SetControlPropertiesFromTemplate ctl, rpt
        ctl.BorderStyle = 0
    
        Select Case fld.Type
             Case dbMemo:
                 ctl.Height = 900
                 isMemo = True
             Case dbDouble:
                 ctl.Format = "Standard"
         End Select
         
        ctl.CanGrow = True
        ctl.InSelection = True
        
        ''Generate the label just above the control
        Dim ControlCaption
        If Not DoesPropertyExists(fld.Properties, "Caption") Then
            If IsNull(rs.fields("VerboseName")) Then
                ControlCaption = AddSpaces(rs.fields("ModelField"))
            Else
                ControlCaption = rs.fields("VerboseName")
            End If
        Else
            ControlCaption = fld.Properties("Caption")
        End If
        
        Set ctl = CreateReportControl(rpt.Name, acLabel, , ctl.Name, ControlCaption & ":", x, maxY)

        SetControlPropertiesFromTemplate ctl, rpt
        ctl.Width = fldWidth
        ctl.TextAlign = 2
        ctl.InSelection = True
        ctl.BackStyle = 1


SetVariables:
    
        CurrentCol = CurrentCol + rs.fields("Columns")
        
        x = x + fldWidth
                 
NextField:
        
        rs.MoveNext
        
    Loop
    
    DoCmd.RunCommand acCmdStackedLayout
    DoCmd.RunCommand acCmdControlPaddingNone
    
    For Each ctl In rpt.controls
        If ctl.InSelection Then
            ctl.Left = 0
            Exit For
        End If
    Next ctl
    
    Dim rptWidth
    rptWidth = (8.5 - (0.25 * 2)) * 1440
    
    ''Write the pageheader (The Report Title & Current Date Time)
    Set ctl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , , 0, 0, 4800, 400)
    SetControlPropertiesFromTemplate ctl, rpt
    ctl.fontSize = 12
    ctl.Height = 1000
    ctl.Caption = AddSpaces(GetModelPlural(Model, VerbosePlural, ""))
    
    ''Current Date Time
    fldWidth = 3000
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageHeader, , , rptWidth - fldWidth, 0, fldWidth, 400)
    SetControlPropertiesFromTemplate ctl, rpt
    ctl.ControlSource = "=Now()"
    ctl.BorderStyle = 0
    
    ''At pagefooter (The Page x of y)
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageFooter, , , 0, 0, fldWidth, 400)
    SetControlPropertiesFromTemplate ctl, rpt
    ctl.ControlSource = "=""Page "" & [Page] & "" of "" & [Pages]"
    ctl.BorderStyle = 0
    
    rpt.Section(acHeader).Height = 0
    rpt.Section(acHeader).BackColor = RGB(254, 254, 254)
    rpt.Section(acFooter).Height = 0
    rpt.Section(acPageFooter).Height = 0
    rpt.Section(acPageHeader).Height = 0
    rpt.Section(acDetail).Height = 0
    rpt.Section(acDetail).AlternateBackColor = RGB(254, 254, 254)
    rpt.Section(acDetail).BackColor = RGB(254, 254, 254)
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    frmName = rpt.Name
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("rptCol", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("rptCol", Model, "s")
    End If
    
    If Not IsNull(SubformName) Then
        baseFormName = concat("rptCol", SubformName)
    End If
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, rpt, 7
    End If
    
    ''Override
    OverrideProperties ModelID, 7, rpt
    
    DoCmd.Close acReport, rpt.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not RptExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    
    ''Insert the newly created form to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acReport, customFrmName
    
    DoCmd.Rename customFrmName, acReport, frmName
    DoCmd.OpenReport customFrmName, acViewPreview
    
End Function
 
Public Function GetTableName(Model, Optional VerbosePlural, Optional QueryName = Null, Optional forceTable As Boolean = False) As String
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryModels WHERE Model = " & EscapeString(Model)
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If Not rs.EOF Then
        Dim ModelID: ModelID = rs.fields("ModelID")
        Dim TableName: TableName = rs.fields("TableName")
        If Not isFalse(TableName) Then
            GetTableName = TableName
            Exit Function
        End If
        GetTableName = GetTableNameFromModelID(ModelID, forceTable)
    End If
    
End Function

Public Function GetTableNameFromModelID(ModelID, Optional forceTable As Boolean = False)

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    
    Dim QueryName, Model, VerbosePlural, TableName
    QueryName = rs.fields("QueryName")
    Model = rs.fields("Model")
    VerbosePlural = rs.fields("VerbosePlural")
    TableName = ELookup("tblSupplementalModels", "ModelID = " & ModelID, "TableName")
    
    ''QueryName will supercede this TableName
    If isFalse(QueryName) And Not isFalse(TableName) Then
        GetTableNameFromModelID = TableName
        Exit Function
    End If
    
    If Not rs.EOF Then
        If Not IsNull(QueryName) And Not forceTable Then
            GetTableNameFromModelID = QueryName
            Exit Function
        End If
    
        If Not IsNull(VerbosePlural) And Not VerbosePlural = "" Then
            GetTableNameFromModelID = concat("tbl", replace(VerbosePlural, " ", ""))
        Else
            GetTableNameFromModelID = concat("tbl", Model, "s")
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Public Function GetModelPlural(Model, VerbosePlural, Optional prefix = "tbl", Optional VerbosePluralCaption = "") As String
    
    ''VerbosePluralCaption will in reality be a forced setting of the Caption
    If Not isFalse(VerbosePluralCaption) Then
        GetModelPlural = VerbosePluralCaption
    Else
        If Not IsNull(VerbosePlural) Then
            GetModelPlural = concat(prefix, replace(VerbosePlural, " ", ""))
        Else
            GetModelPlural = concat(prefix, Model, "s")
        End If
    End If

End Function


Public Sub CreatePrimaryKey(Model, tblDef As TableDef, Optional pkName = "", Optional FieldTypeID = Null)

On Error GoTo ErrHandler:
    Dim fld As DAO.field, idx As DAO.Index
    If isFalse(pkName) Then
        pkName = concat(Model, "ID")
    End If
    
    If isFalse(FieldTypeID) Or FieldTypeID = dbLong Then
        Set fld = AddField(tblDef, pkName, dbLong, dbAutoIncrField)
    Else
        Set fld = AddField(tblDef, pkName, FieldTypeID)
    End If
    
    If Not DoesPropertyExists(tblDef.indexes, pkName) Then
        Set idx = tblDef.CreateIndex(pkName)
    
        With idx
            .fields.Append .CreateField(pkName)
            .Primary = True
        End With
        
        tblDef.indexes.Append idx
    End If

    Exit Sub
ErrHandler:
    If Err.Number = 3283 Then
        Exit Sub
    End If

End Sub

Public Function GetFieldName(ForeignKey, ModelField, Optional ByPassFK As Boolean = False) As String
    
    If IsNull(ForeignKey) Then
        GetFieldName = ModelField
    Else
        If ByPassFK Then
            GetFieldName = ModelField
        Else
            GetFieldName = concat(ForeignKey, "ID")
        End If
    End If

End Function

Public Function GetFieldCaption(VerboseName, fldName, Optional VerboseCaption = Null) As String
    
    If Not IsNull(VerboseCaption) Then
        GetFieldCaption = VerboseCaption
        Exit Function
    End If
    
    If IsNull(VerboseName) Then
        GetFieldCaption = AddSpaces(fldName)
    Else
        GetFieldCaption = VerboseName
    End If
    
End Function

Public Function AddField(tblDef As TableDef, fldName, fldType, Optional fldAttr As Variant) As DAO.field
    
    Dim fld As DAO.field
    
    If Not DoesPropertyExists(tblDef.fields, fldName) Then
    
        Set fld = tblDef.CreateField(fldName, fldType)
        
        If Not IsMissing(fldAttr) Then
            fld.attributes = fld.attributes Or fldAttr
        End If
        
        With tblDef.fields
            .Append fld
            .Refresh
        End With
    Else
        
        Set fld = tblDef.fields(fldName)
        
    End If
    
    Set AddField = fld
    
End Function

Public Function SaveFormData2(frm As Object, Model As String, Optional CancelOnClose As Boolean = False) As Boolean

    If areDataValid2(frm, Model) Then
    
'        If validationSuccessCB <> "" Then
'            Run validationSuccessCB, frm
'        End If
        
        If Not frm.NewRecord Then
            UpdateFormData2 frm, Model
        End If
        
    Else
        
        If CancelOnClose Then
            frm.Undo
        End If
        
    End If
    
End Function

Public Function GetModelByPrimaryKey(PrimaryKey) As Recordset
    
    Dim Model, ModelID
    Model = Left(PrimaryKey, Len(PrimaryKey) - 2)
    ModelID = ELookup("tblModels", "Model = " & EscapeString(Model), "ModelID")
    
    Set GetModelByPrimaryKey = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    
End Function

Public Function GetTableNameByPrimaryKey(PrimaryKey)

    Dim Model, ModelID
    Model = Left(PrimaryKey, Len(PrimaryKey) - 2)
    ModelID = ELookup("tblModels", "Model = " & EscapeString(Model), "ModelID")
    
    GetTableNameByPrimaryKey = GetTableNameFromModelID(ModelID)
    
End Function


Public Function UpdateFormData2(frm As Object, Model As String)
    
    Dim rs As Recordset, ModelID, tblName As String
    Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE Model = """ & Model & """")
    tblName = GetTableName(Model, rs.fields("VerbosePlural"))
    ModelID = rs.fields("ModelID")
    
    Set rs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelID = " & ModelID)
    
    Dim ctl As control, RecordID As Variant
    
    Dim fieldName As String, FieldTypeID As Integer, currentValue, oldValue, PrimaryKey
    Dim updateStatement() As String, i As Integer
    
    Do Until rs.EOF
    
        fieldName = rs.fields("ModelField")
        FieldTypeID = rs.fields("FieldTypeID")
        PrimaryKey = GetPrimaryKeyFromTable(ModelID)
        
        If Not DoesPropertyExists(CurrentDb.TableDefs, tblName) Then
            GoTo NextField:
        End If
        
        If ControlExists(fieldName, frm) And DoesPropertyExists(CurrentDb.TableDefs(tblName), fieldName) Then
            ''Get the oldvalue from the table
            RecordID = frm(PrimaryKey)
            currentValue = frm(fieldName)
            oldValue = GetOldValue(tblName, fieldName, PrimaryKey, RecordID)
            If oldValue <> currentValue Or (Not IsNull(oldValue) Xor Not IsNull(currentValue)) Then
                ReDim Preserve updateStatement(i)
                updateStatement(i) = fieldName & " = " & ReturnStringBasedOnType(currentValue, FieldTypeID)
                Update_Log tblName, oldValue, currentValue, RecordID, fieldName
                i = i + 1
            End If
        End If
NextField:
        rs.MoveNext
    Loop
    
End Function

Public Function QueryProperty()

On Error Resume Next
    Dim db As DAO.Database, qDef As DAO.QueryDef, fld As DAO.field, prop As DAO.property
    
    Set db = CurrentDb
    Set qDef = db.QueryDefs("qryEmployees")
    
    For Each fld In qDef.fields
        'Debug.Print fld.SourceTable
    Next fld
    

End Function

Public Function areDataValid2(frm As Object, Optional Model As String) As Boolean
    
    ''Fetch the tblModelFields from the specific form
    Dim rs As Recordset, ModelID
    Dim ModelField, VerboseName, ValidationString, ControlCaption, PrimaryKey, QueryName
    ModelID = ELookup("tblModels", "Model = " & EscapeString(Model), "ModelID")
    PrimaryKey = ELookup("tblModels", "ModelID = " & ModelID, "PrimaryKey")
    QueryName = ELookup("tblModels", "ModelID = " & ModelID, "QueryName")
    If QueryName = "" Then QueryName = Null
    If PrimaryKey = "" Then PrimaryKey = concat(Model, "ID")
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblModelFields where ModelID = " & ModelID & " ORDER BY FieldOrder")
    
    Dim ValidationArr As New clsArray, ValidationRule As Variant
    Dim ctl As control
    
    Do Until rs.EOF
        ModelField = rs.fields("ModelField")
        VerboseName = rs.fields("VerboseName")
        ValidationString = rs.fields("ValidationString")
        
        ControlCaption = GetFieldCaption(VerboseName, ModelField)
    
        If ControlExists(ModelField, frm) Then
            Set ctl = frm.controls(ModelField)
            
            If Not IsNull(ValidationString) Then
                ValidationArr.arr = Split(ValidationString, " ")
                For Each ValidationRule In ValidationArr.arr
                    Select Case Trim(ValidationRule)
                        Case "required":
                            If IsNull(ctl) Or ctl = "" Then
                                ShowError ControlCaption & " is a required field."
                                If ControlExists(ctl.Name, frm) And ctl.ColumnHidden = False And (ctl.Enabled = True And ctl.Locked = True) Then
                                    ctl.SetFocus
                                End If
                                areDataValid2 = False
                                DoCmd.CancelEvent
                                Exit Function
                            End If
                        Case "+":
                            If ctl < 0 Then
                                ShowError ControlCaption & " must be not be less than 0."
                                If ControlExists(ctl.Name, frm) And ctl.ColumnHidden = False Then
                                    ctl.SetFocus
                                End If
                                areDataValid2 = False
                                DoCmd.CancelEvent
                                Exit Function
                            End If
                    End Select
                Next ValidationRule
            End If
            
        End If
        rs.MoveNext
    Loop
    
    ''Validate Uniqueness of records
    Set rs = CurrentDb.OpenRecordset("SELECT * from tblModelFields where ModelID = " & ModelID & " And ValidationString Like '*unique*'")
    Dim i As Integer
    Dim filterStr() As String
    Dim fieldCaptions() As String
    Dim fieldValue As String
    If Not rs.EOF Then
        Do Until rs.EOF
            fieldValue = frm(rs.fields("ModelField"))
            ReDim Preserve filterStr(i)
            Select Case rs.fields("FieldTypeID")
                Case 10:
                    fieldValue = EscapeString(fieldValue)
            End Select
            filterStr(i) = rs.fields("ModelField") & " = " & fieldValue
            ReDim Preserve fieldCaptions(i)
            fieldCaptions(i) = GetFieldCaption(rs.fields("VerboseName"), rs.fields("ModelField"))
            i = i + 1
            rs.MoveNext
        Loop
        
        Dim filterStmt As String
        filterStmt = Join(filterStr, " And ")
        
        ''If not a new record then disregard this record from the filter
        If Not frm.NewRecord Then
            filterStmt = filterStmt & " And " & PrimaryKey & " <> " & frm(PrimaryKey)
        End If
        
        Dim errorMsg As String
        errorMsg = Join(fieldCaptions, " | ") & " is already present from the record list"
        
        Dim rsName As String
        rsName = GetTableName(Model, ELookup("tblModels", "ModelID = " & ModelID, "VerbosePlural"), QueryName)
        
        If isPresent(rsName, filterStmt) Then
            ShowError errorMsg
            DoCmd.CancelEvent
            If frm.Name Like "cont*" Or frm.Name Like "dsht*" Then
                ''frm.Undo
            End If
            areDataValid2 = False
            Exit Function
        End If
        
    End If
    
    ''Look for AdditionalValidation from tblTables
    Dim TableWideValidation
    TableWideValidation = ELookup("tblModels", "ModelID = " & ModelID, "TableWideValidation")
    If TableWideValidation <> "" Then
        If Not Application.Run(TableWideValidation, frm) Then
            areDataValid2 = False
            DoCmd.CancelEvent
            Exit Function
        End If
    End If
    
    If frm.NewRecord Then
        frm.OnClose = "=RequeryOnClose2('" & Model & "',True)"
    End If
    
    areDataValid2 = True
    
End Function

Public Function RequeryOnClose2(Model As String, Optional shouldRequery As Boolean)

    If shouldRequery Then
        'On Error Resume Next
        Dim requeryForms As Variant, requeryFormArray() As String, RequeryForm As Variant, PrimaryKey
        Dim rs As Recordset, frm As Form, rsClone As Recordset, VerbosePlural, PluralizedName, ModelID
        Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE Model = '" & Model & "'")
        
        VerbosePlural = rs.fields("VerbosePlural")
        PrimaryKey = rs.fields("PrimaryKey")
        If IsNull(PrimaryKey) Then PrimaryKey = concat(Model, "ID")
        ModelID = rs.fields("ModelID")
        
        If Not IsNull(VerbosePlural) Then
            PluralizedName = concat(replace(VerbosePlural, " ", ""))
        Else
            PluralizedName = concat(Model, "s")
        End If
        
        If IsFormOpen(concat("main", PluralizedName)) Then
            Forms(concat("main", PluralizedName)).subform.Requery
            
            On Error Resume Next
            Set frm = Forms(concat("main", PluralizedName)).subform.Form
                
            'ReturnMainForm(requeryForm).SetFocus
            Set rsClone = frm.RecordsetClone
            Dim pkType: pkType = rsClone.fields(PrimaryKey).Type
            Dim pkValue: pkValue = frm(PrimaryKey)
            If pkType = 10 Then pkValue = EscapeString(frm(PrimaryKey))
            
            rsClone.FindFirst PrimaryKey & " = " & pkValue
         
            If Not rsClone.NoMatch Then
                frm.Bookmark = rsClone.Bookmark
            End If
        End If
        
        'Eval "Forms!Main" & PluralizedName & ".Requery"
        
        ''Get all the foreign key models of this model
        Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = "tblModelFields"
            .AddFilter "ModelID = " & ModelID & " AND " & _
                       "ParentModelID IS NOT NULL"
            .fields = "ParentModelID"
            .GroupBy = "ParentModelID"
            sqlStr = .sql
        End With
        
        ''SELECT STATEMENT
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = "tblModels"
            .fields = "Model,VerbosePlural"
            .joins.Add GenerateJoinObj(sqlStr, "ModelID", "temp", "ParentModelID")
            Set rs = .Recordset
        End With
        
        Dim PluralizedName2 As String
        
        Do Until rs.EOF
        
            VerbosePlural = rs.fields("VerbosePlural")
            Model = rs.fields("Model")
            
            If Not IsNull(VerbosePlural) Then
                PluralizedName2 = concat(replace(VerbosePlural, " ", ""))
            Else
                PluralizedName2 = concat(Model, "s")
            End If
            
            Dim SubformName
            SubformName = concat("sub", PluralizedName)
On Error GoTo SkipForm:

            ''Check here if the form is existing and opened..
            If DoesPropertyExists(Forms, concat("frm", PluralizedName2)) Then
                
                Forms(concat("frm", PluralizedName2))(SubformName).Requery
                Set frm = Forms(concat("frm", PluralizedName2))(SubformName).Form
            
                'ReturnMainForm(requeryForm).SetFocus
                Set rsClone = frm.RecordsetClone
                pkType = rsClone.fields(PrimaryKey).Type
                pkValue = frm(PrimaryKey)
                If pkType = 10 Then pkValue = EscapeString(frm(PrimaryKey))
                
                rsClone.FindFirst PrimaryKey & " = " & pkValue
                If Not rsClone.NoMatch Then
                    frm.Bookmark = rsClone.Bookmark
                End If
            End If
            
            
SkipForm:
            rs.MoveNext
        Loop
        
    End If
    
    Exit Function
    
ErrHandler:
    
    If Err.Number = 2450 Then
        GoTo SkipForm
    Else
        ShowError concat(Err.Number, vbCrLf, Err.description)
    End If
    
End Function

Public Function SetFormProperties(FormTypeID, frm As Form)

    ''Set the Form Properties
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblFrmProps WHERE FormTypeID = " & FormTypeID)
    Do Until rs.EOF
        frm.Properties(rs.fields("FormProp")) = rs.fields("FormPropValue")
        rs.MoveNext
    Loop
    
    ''Also open the form template - detail section - backColor, alternateBackColor
    CopyControlTemplateProperties frm
    
End Function

Public Sub CopyControlTemplateProperties(obj As Object, Optional ctlName = "", Optional insideTab As Boolean = False)
    
    DoCmd.OpenForm "frmControlTemplate", acDesign, , , , acHidden
    Dim frm1 As Form: Set frm1 = Forms("frmControlTemplate")
    
    ''If controlName is false then use the form's back color and alternate back color
    If isFalse(ctlName) Then
        obj.Section(acDetail).BackColor = frm1.Section(acDetail).BackColor
        obj.Section(acDetail).AlternateBackColor = frm1.Section(acDetail).AlternateBackColor
    ''Else use the properties listed from tblTemplateControls
    Else
        If obj(ctlName).Tag Like "*DontFormat*" Then Exit Sub
        
        Dim ControlType: ControlType = obj(ctlName).ControlType
        Select Case ControlType
            Case acTextBox:
                CopyProperties obj, ctlName, "TextControl", insideTab
            Case acComboBox:
                CopyProperties obj, ctlName, "ComboControl", insideTab
            Case acLabel:
                CopyProperties obj, ctlName, "LabelControl", insideTab
            Case acTabCtl:
                CopyProperties obj, ctlName, "TabControl", insideTab
            Case acCommandButton:
                CopyProperties obj, ctlName, "ButtonControl", insideTab
            Case acSubform:
                CopyProperties obj, ctlName, "SubformControl", insideTab
        End Select
    End If
    
End Sub



Public Sub CopyProperties(obj As Object, ctlName, masterControlName, Optional insideTab = False)
    ''TABLE: tblTemplateControls Fields: TemplateControlID|Name|ControlProperty|Timestamp|CreatedBy|Record
    If insideTab Then
        masterControlName = masterControlName & "InTab"
    End If
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblTemplateControls WHERE Name = " & EscapeString(masterControlName))
    
    If Not IsFormOpen("frmControlTemplate") Then
        DoCmd.OpenForm "frmControlTemplate", acDesign, , , , acHidden
    End If
    Dim frm1 As Form: Set frm1 = Forms("frmControlTemplate")
    Dim ControlProperty
    Do Until rs.EOF
        ControlProperty = rs.fields("ControlProperty")
        On Error Resume Next
        obj(ctlName).Properties(ControlProperty) = frm1.controls(masterControlName).Properties(ControlProperty)
        rs.MoveNext
    Loop
    
End Sub

Public Function AddTableDef(db As DAO.Database, tblName) As TableDef
    
    If Not DoesPropertyExists(db.TableDefs, tblName) Then
        Set AddTableDef = db.CreateTableDef(tblName)
    Else
        Set AddTableDef = db.TableDefs(tblName)
    End If
    
End Function

Private Function AddIndex(tblDef As TableDef, idxName, fldName, Optional IsUnique As Boolean = False)
    
    Dim idx As DAO.Index
    
    If Not DoesPropertyExists(tblDef.indexes, idxName) Then
        
        Set idx = tblDef.CreateIndex(idxName)
        With idx
            .fields.Append .CreateField(fldName)
            .Unique = IsUnique
        End With
        
        tblDef.indexes.Append idx
        
    End If
        
End Function

Public Function CreateProperty(fld As DAO.field, propertyName, PropertyType, PropertyValue)
    
On Error GoTo Err_Handler:
    
    If Not DoesPropertyExists(fld.Properties, propertyName) Then
        fld.Properties.Append fld.CreateProperty(propertyName, PropertyType, PropertyValue)
    Else
        fld.Properties(propertyName) = PropertyValue
    End If
Err_Handler:
    Exit Function

End Function



'Private Sub cmdCreateTableDef_Click()
'
'    Dim tblName As String
'    Dim db As DAO.Database
'    Dim tblDef As DAO.TableDef, fld As DAO.Field, idx As DAO.Index
'    Dim rel As DAO.Relation, relName, primaryField, foreignField, primaryTable, foreignTable
'    Dim pkName, fldName, idxName, MainField
'    Dim rs As Recordset, fldCaption, rowSourceSQL, ForeignKey, rs2 As Recordset
'
'    ''Get the name of the TableDef
'    tblName = GetTableName(Model, VerbosePlural)
'
'    Set db = CurrentDb
'    Set tblDef = AddTableDef(db, tblName)
'
'    ''Create the table fields using
'    ''tblDef.fields.Append .CreateField("FirstName", dbText)
'    ''If the field already exists then skip if not then append
'    ''First create the primary key (this is an autonumber field)
'    CreatePrimaryKey Model, tblDef
'
'    ''Add Custom Fields here via loop
'    Set rs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelID = " & ModelID & " ORDER BY FieldOrder ASC")
'
'    Do Until rs.EOF
'
'        fldName = GetFieldName(rs.Fields("ForeignKey"), rs.Fields("ModelField"))
'
'
'        Set fld = AddField(tblDef, fldName, rs.Fields("FieldTypeID"))
'
'        rs.MoveNext
'    Loop
'
'    ''Also add the Timestamp and CreatedBy field
'    ''Timestamp Field
'    ''Created by will be set into a combo box looked up into tblUsers with Username as its field
'    ''And also create index for this field..
'    fldName = "Timestamp"
'    Set fld = AddField(tblDef, fldName, dbDate)
'    AddIndex tblDef, fldName, fldName
'
'    ''CreatedBy Field
'    fldName = "CreatedBy"
'    Set fld = AddField(tblDef, fldName, dbLong)
'    AddIndex tblDef, fldName, fldName
'
'    ''RecordImportID Field
'    fldName = "RecordImportID"
'    Set fld = AddField(tblDef, fldName, dbLong)
'    AddIndex tblDef, fldName, fldName
'
'    If Not DoesPropertyExists(db.TableDefs, tblName) Then
'        db.TableDefs.Append tblDef
'    End If
'
'    ''Set field properties here
'    rs.MoveFirst
'    Do Until rs.EOF
'
'
'        fldName = GetFieldName(rs.Fields("ForeignKey"), rs.Fields("ModelField"))
'        fldCaption = GetFieldCaption(rs.Fields("VerboseName"), fldName)
'
'        Set fld = tblDef.Fields(fldName)
'
'        ''Set the Caption
'        CreateProperty fld, "Caption", dbText, fldCaption
'
'        ''Set the index
'        If rs.Fields("IsIndexed") Then
'            AddIndex tblDef, fldName, fldName
'        End If
'
'        ''Set the foreign key
'        If Not IsNull(rs.Fields("ForeignKey")) Then
'
'            CreateProperty fld, "DisplayControl", dbInteger, acComboBox
'
'            ForeignKey = rs.Fields("ForeignKey")
'            MainField = ELookup("tblModels", Concat("Model = ", EscapeString(ForeignKey)), "MainField")
'            foreignTable = tblName: foreignField = Concat(ForeignKey, "ID")
'
'            Set rs2 = ReturnRecordset("SELECT * FROM tblModels WHERE Model = " & EscapeString(ForeignKey))
'
'            ''Get the name of the TableDef
'            primaryTable = GetTableName(rs2.Fields("Model"), rs2.Fields("VerbosePlural"))
'
'            primaryField = Concat(ForeignKey, "ID")
'
'            rowSourceSQL = Concat("SELECT ", primaryField, ",", MainField, " FROM ", primaryTable, " ORDER BY ", MainField)
'
'            CreateProperty fld, "RowSource", dbText, rowSourceSQL
'            CreateProperty fld, "ColumnCount", dbInteger, 2
'            CreateProperty fld, "ColumnWidths", dbText, "0;1"
'
'            ''Create relationship with the primaryTable
'            relName = Concat(primaryTable, primaryField, "_", foreignTable, foreignField)
'            If DoesPropertyExists(db.Relations, relName) Then
'                db.Relations.Delete relName
'            End If
'
'            Set rel = db.CreateRelation(relName, primaryTable, foreignTable, &H2000 + dbRelationUpdateCascade)
'            Set fld = rel.CreateField(primaryField)
'            fld.foreignName = foreignField
'
'            rel.Fields.Append fld
'            db.Relations.Append rel
'
'        End If
'
'        ''Set the default value property
'        If Not IsNull(rs.Fields("DefaultValue")) Then
'
'            fld.Properties("DefaultValue") = rs.Fields("DefaultValue")
'
'        End If
'
'        ''Set the value list
'        If Not IsNull(rs.Fields("PossibleValues")) Then
'
'            CreateProperty fld, "DisplayControl", dbInteger, acComboBox
'            CreateProperty fld, "RowSourceType", dbText, "Value List"
'            CreateProperty fld, "RowSource", dbText, rs.Fields("PossibleValues")
'            CreateProperty fld, "LimitToList", dbBoolean, True
'
'        End If
'
'        ''Default Format
'        If rs.Fields("FieldTypeID") = dbDouble Then
'
'             CreateProperty fld, "Format", dbText, "Standard"
'
'        End If
'
'        rs.MoveNext
'    Loop
'
'    fldName = "Timestamp"
'    Set fld = tblDef.Fields(fldName)
'    fld.Properties("DefaultValue") = "=Now()"
'
'    fldName = "CreatedBy"
'    Set fld = tblDef.Fields(fldName)
'    CreateProperty fld, "Caption", dbText, "Created By"
'    CreateProperty fld, "DisplayControl", dbInteger, acComboBox
'    CreateProperty fld, "RowSource", dbText, "SELECT UserID, UserName FROM tblUsers ORDER BY UserName"
'    CreateProperty fld, "ColumnCount", dbInteger, 2
'    CreateProperty fld, "ColumnWidths", dbText, "0;1"
'
'    idxName = fldName
'    AddIndex tblDef, idxName, fldName
'
'    ''Create relationship with tblUsers
'    foreignTable = tblName: foreignField = "CreatedBy": primaryTable = "tblUsers": primaryField = "UserID"
'    relName = Concat(primaryTable, primaryField, "_", foreignTable, foreignField)
'    If DoesPropertyExists(db.Relations, relName) Then
'        db.Relations.Delete relName
'    End If
'    Set rel = db.CreateRelation(relName, primaryTable, foreignTable, &H2000 + dbRelationUpdateCascade)
'    Set fld = rel.CreateField(primaryField)
'    fld.foreignName = foreignField
'
'    rel.Fields.Append fld
'    db.Relations.Append rel
'
'    MsgBox "Table Def successfully created.."
'
'End Sub

Public Function DeclareVariables(rsName, Optional encloser As String = "frm")
    
    Dim tblDef As Object, db As DAO.Database, fld As DAO.field
    Set db = CurrentDb
    
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set tblDef = db.TableDefs(rsName)
    Else
        Set tblDef = db.QueryDefs(rsName)
    End If
    
    ''The Line where the DIM variables are declared
    Dim fieldArr As New clsArray
    For Each fld In tblDef.fields
        Select Case fld.Name
            Case "Timestamp", "CreatedBy", "RecordImportID":
                
            Case Else:
                fieldArr.Add fld.Name
        End Select
    Next fld
    
    Dim lines As New clsArray
    Debug.Print "Dim " & fieldArr.JoinArr
    lines.Add "Dim " & fieldArr.JoinArr
    
    Dim fieldItem As Variant
    
    For Each fieldItem In fieldArr.arr
        Select Case fieldItem
            Case "Timestamp", "CreatedBy", "RecordImportID":
                
            Case Else:
                Debug.Print fieldItem & " = " & encloser & "(" & EscapeString(fieldItem) & ")"
                lines.Add fieldItem & " = " & encloser & "(" & EscapeString(fieldItem) & ")"
        End Select
        
    Next fieldItem
    
    CopyToClipboard lines.JoinArr(vbNewLine)
    
End Function

Public Function MakeProductionCopy()
    
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    ''Make copy of the current database
    ''Make sure that the name is different from the current one
    Dim currentPath, dbName, copyDbName, dbNameArr As New clsArray
    currentPath = CurrentProject.path
    dbName = CurrentProject.Name
    dbNameArr.arr = Split(dbName, ".")
    copyDbName = concat(dbNameArr.arr(0), "-Prod.", dbNameArr.arr(1))
    
    fso.CopyFile concat(currentPath, "\", dbName), concat(currentPath, "\", copyDbName), True
    
End Function

Public Function RemoveNonSystemTables()
    

    If MsgBox("Are you sure you want to remove all the non-system related objects?", vbYesNo) = vbNo Then
        Exit Function
    End If
    
    DoCmd.Close acForm, "mainModels", acSaveNo
    
    Dim rs As Recordset, rel As DAO.Relation, db As DAO.Database, relName
    Set rs = ReturnRecordset("SELECT * FROM qryModelRelatedObjects WHERE IsSystemTable = 0")
    Set db = CurrentDb
    
On Error GoTo ErrHandler:
    Dim relationArr As New clsArray, modelArr As New clsArray, modelArrItem
    Do Until rs.EOF
        
        If rs.fields("ObjectTypeID") = acTable Then
            For Each rel In db.Relations
                If rel.foreignTable = rs.fields("ObjectName") Then
                    relationArr.Add rel.Name
                End If
                If rel.Table = rs.fields("ObjectName") Then
                    relationArr.Add rel.Name
                End If
            Next rel
            
      
        End If
        
        'Debug.Print rs.Fields("ObjectName")
        modelArr.Add rs.fields("ModelID")
        'DoCmd.DeleteObject rs.Fields("ObjectTypeID"), rs.Fields("ObjectName")
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    If relationArr.count > 0 Then
        For Each relName In relationArr.arr
            If DoesPropertyExists(db.Relations, relName) Then db.Relations.Delete relName
        Next relName
    End If
    
    Set rs = ReturnRecordset("SELECT * FROM qryModelRelatedObjects WHERE IsSystemTable = 0")
    Do Until rs.EOF
        DoCmd.DeleteObject rs.fields("ObjectTypeID"), rs.fields("ObjectName")
        rs.MoveNext
    Loop

    For Each modelArrItem In modelArr.arr
        RunSQL "DELETE FROM tblModels WHERE ModelID = " & modelArrItem
    Next modelArrItem
    
    DoCmd.OpenForm "mainModels"
    
    Exit Function

ErrHandler:
    If Err.Number = 7874 Then
       Resume Next
    Else
        MsgBox Err.Number & vbCrLf & Err.description
    End If

End Function

Public Function CreateTableDef(frm As Object, Optional notifySuccess As Boolean = True)
    
    Dim tblName As String
    Dim db As DAO.Database
    Dim tblDef As DAO.TableDef, fld As DAO.field, idx As DAO.Index
    Dim rel As DAO.Relation, relName, primaryField, foreignField, primaryTable, foreignTable
    Dim pkName, fldName, idxName
    Dim rs As Recordset, fldCaption, rowSourceSQL, ForeignKey, rs2 As Recordset
    
    ''TABLE: tblModels Fields: ModelID|Model|VerboseName|VerbosePlural|MainField|TableWideValidation|FormColumns
    ''SetFocus|IsKeyVisible|QueryName|OnFormCreate|SubformName|UserQueryFields|IsSystemTable|Timestamp|CreatedBy
    ''RecordImportID|PrimaryKey|VerbosePluralCaption
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation
    Dim PrimaryKey, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, Timestamp, CreatedBy
    
    ModelID = frm("ModelID")
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
    Timestamp = frm("Timestamp")
    CreatedBy = frm("CreatedBy")
    PrimaryKey = frm("PrimaryKey")
        
    ''Get the name of the TableDef
    tblName = GetTableName(Model, VerbosePlural, , True)
    
    Set db = CurrentDb
    Set tblDef = AddTableDef(db, tblName)
    
    ''Create the table fields using
    ''tblDef.fields.Append .CreateField("FirstName", dbText)
    ''If the field already exists then skip if not then append
    ''First create the primary key (this is an autonumber field)
    ''Get the primary key's fieldtype if it's present from the ModelFields
    Dim FieldTypeID
    If Not isFalse(PrimaryKey) Then
        FieldTypeID = ELookup("tblModelFields", "ModelID = " & ModelID & " AND ModelField = " & Esc(PrimaryKey), "FieldtypeID")
    End If
    CreatePrimaryKey Model, tblDef, PrimaryKey, FieldTypeID
    
    ''Add Custom Fields here via loop
    Set rs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelID = " & ModelID & _
                             " AND (FieldSource IS NULL OR FieldSource = " & EscapeString(tblName) & ") " & _
                             "AND IsAnExpression = 0 AND ControlSource IS NULL ORDER BY FieldOrder ASC")
    
    Do Until rs.EOF
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"), True)
        Set fld = AddField(tblDef, fldName, rs.fields("FieldTypeID"))
        rs.MoveNext
    Loop
    
    ''Also add the Timestamp and CreatedBy field
    ''Timestamp Field
    ''Created by will be set into a combo box looked up into tblUsers with Username as its field
    ''And also create index for this field..
    fldName = "Timestamp"
    Set fld = AddField(tblDef, fldName, dbDate)
    AddIndex tblDef, fldName, fldName
    
    ''CreatedBy Field
    fldName = "CreatedBy"
    Set fld = AddField(tblDef, fldName, dbLong)
    AddIndex tblDef, fldName, fldName
    
    ''Create the RecordImportID
    fldName = "RecordImportID"
    Set fld = AddField(tblDef, fldName, dbText)
    AddIndex tblDef, fldName, fldName
    
    If Not DoesPropertyExists(db.TableDefs, tblName) Then
        db.TableDefs.Append tblDef
    End If
    
    ''Set field properties here
    If rs.recordCount > 0 Then rs.MoveFirst
    Do Until rs.EOF
        
        fldName = rs.fields("ModelField")
        fldCaption = GetFieldCaption(rs.fields("VerboseName"), fldName)
        
        Set fld = tblDef.fields(fldName)
    
        ''Set the Caption
        CreateProperty fld, "Caption", dbText, fldCaption
        
        ''Set the index
        If rs.fields("IsIndexed") Then
            Dim ValidationString: ValidationString = rs.fields("ValidationString")
            Dim IsUnique As Boolean: IsUnique = False
            Dim UniqueCount: UniqueCount = ECount("qryModelFields", "ValidationString Like '*unique*' AND ModelID = " & ModelID)
            If Not isFalse(ValidationString) Then
                If ValidationString Like "*unique*" And UniqueCount = 1 Then
                    IsUnique = True
                End If
            End If
            AddIndex tblDef, fldName, fldName, IsUnique
        End If
        
        ''Set the foreign key
        If Not IsNull(rs.fields("ParentModelID")) Then
            
            CreateProperty fld, "DisplayControl", dbInteger, acComboBox
            
            PrimaryKey = ELookup("tblModels", "ModelID = " & rs.fields("ParentModelID"), "PrimaryKey")
            ForeignKey = ELookup("tblModels", "ModelID = " & rs.fields("ParentModelID"), "Model")
            MainField = ELookup("tblModels", "ModelID = " & rs.fields("ParentModelID"), "MainField")
            
            Dim rs3 As Recordset: Set rs3 = ReturnRecordset("SELECT * FROM tblSupplementalModels WHERE ModelID = " & rs.fields("ParentModelID") & _
                " AND NOT UseAsModel IS NULL")
            Dim UseAsModel
            If Not rs3.EOF Then
                UseAsModel = rs3.fields("UseAsModel")
                ForeignKey = ELookup("tblModels", "ModelID = " & UseAsModel, "Model")
                MainField = ELookup("tblModels", "ModelID = " & UseAsModel, "MainField")
            End If
            
            If MainField Like "=*" Then
                MainField = replace(MainField, "=", "")
            End If
            
            foreignTable = tblName: foreignField = fldName
            
            Set rs2 = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & rs.fields("ParentModelID"))
            
            ''Get the name of the TableDef
            primaryTable = GetTableName(rs2.fields("Model"), rs2.fields("VerbosePlural"))
            If Not rs3.EOF Then
                primaryTable = GetTableNameFromModelID(UseAsModel)
            End If
            
            primaryField = concat(ForeignKey, "ID")
            
            If Not isFalse(PrimaryKey) Then
                primaryField = PrimaryKey
            End If
            
            rowSourceSQL = concat("SELECT ", primaryField, ",", MainField, " As MainField FROM ", primaryTable, " ORDER BY ", MainField)
            
            CreateProperty fld, "RowSource", dbText, rowSourceSQL
            CreateProperty fld, "ColumnCount", dbInteger, 2
            CreateProperty fld, "ColumnWidths", dbText, "0;1"
            ''CreateProperty fld, "ListItemsEditForm", dbText, GetModelPlural(rs2.fields("Model"), rs2.fields("VerbosePlural"), "frmSimple")
            
            ''Create relationship with the primaryTable
            relName = Left(concat(primaryTable, primaryField, "_", foreignTable, foreignField), 30)
            If DoesPropertyExists(db.Relations, relName) Then
                db.Relations.Delete relName
            End If
            
            Set rel = db.CreateRelation(relName, primaryTable, foreignTable, dbRelationUpdateCascade + dbRelationDeleteCascade)
            Set fld = rel.CreateField(primaryField)
            fld.foreignName = foreignField
 On Error Resume Next
            rel.fields.Append fld
            db.Relations.Append rel
            
        End If
        
        ''If type is Boolean set the display control to checkbox
        If rs.fields("FieldTypeID") = dbBoolean Then
            CreateProperty fld, "DisplayControl", dbInteger, acCheckBox
        End If
        
        If rs.fields("FieldTypeID") = dbText Then
            CreateProperty fld, "AllowZeroLength", dbBoolean, True
        End If
        
        ''Set the default value property
        If Not IsNull(rs.fields("DefaultValue")) Then
            
            fld.Properties("DefaultValue") = rs.fields("DefaultValue")

        End If
        
        ''Set the value list
        If Not isFalse(rs.fields("PossibleValues")) Then
            
            CreateProperty fld, "DisplayControl", dbInteger, acComboBox
            CreateProperty fld, "RowSourceType", dbText, "Value List"
            CreateProperty fld, "RowSource", dbText, QuoteAndJoin(rs.fields("PossibleValues"))
            CreateProperty fld, "LimitToList", dbBoolean, True
            CreateProperty fld, "AllowValueListEdits", dbBoolean, False
            
        End If
        
        ''Default Format
        If rs.fields("FieldTypeID") = dbDouble Then
        
             CreateProperty fld, "Format", dbText, "Standard"
             
        End If
        
        rs.MoveNext
    Loop
    
    fldName = "Timestamp"
    Set fld = tblDef.fields(fldName)
    fld.Properties("DefaultValue") = "=Now()"
    
    fldName = "CreatedBy"
    Set fld = tblDef.fields(fldName)
    CreateProperty fld, "Caption", dbText, "Created By"
    CreateProperty fld, "DisplayControl", dbInteger, acComboBox
    CreateProperty fld, "RowSource", dbText, "SELECT UserID, UserName FROM tblUsers ORDER BY UserName"
    CreateProperty fld, "ColumnCount", dbInteger, 2
    CreateProperty fld, "ColumnWidths", dbText, "0;1"

    idxName = fldName
    AddIndex tblDef, idxName, fldName
    
    ''Create relationship with tblUsers
On Error Resume Next
'    foreignTable = tblName: foreignField = "CreatedBy": primaryTable = "tblUsers": primaryField = "UserID"
'    relName = concat(primaryTable, primaryField, "_", foreignTable, foreignField)
'    If DoesPropertyExists(db.Relations, relName) Then
'        db.Relations.Delete relName
'    End If
'    Set rel = db.CreateRelation(relName, primaryTable, foreignTable, &H2000 + dbRelationUpdateCascade)
'    Set fld = rel.CreateField(primaryField)
'    fld.foreignName = foreignField
'    rel.fields.Append fld
'    db.Relations.Append rel
    
    ''Insert the newly created table to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acTable, tblName

    If notifySuccess Then MsgBox "Table Def successfully created.."
    
    Set fld = Nothing
    Set rel = Nothing
    Set tblDef = Nothing
    Set db = Nothing
    
    CreateUniqueIndex ModelID, tblName
    
End Function

Private Sub CreateUniqueIndex(ModelID, tblName)
   
On Error GoTo ErrHandler:
    ''Add unique index when there's two unique constraint in one table
    Dim fields: fields = Elookups("tblModelFields", "ModelID = " & ModelID & " AND ValidationString LIKE ""*unique*""", "ModelField")
    Dim fieldsArr As New clsArray: fieldsArr.arr = fields
     
    If fieldsArr.count <> 2 Then Exit Sub
    
    Dim sqlStr: sqlStr = "CREATE UNIQUE INDEX uq_[tblName] ON [tblName] ([Fields])"
    
    Dim values As New clsArray: values.arr = ""
    values.Add tblName
    values.Add fields
    
    sqlStr = GetReplacedString(sqlStr, "tblName,Fields", values)
    RunSQL sqlStr
    
    Exit Sub
    
ErrHandler:
    If Err.Number = 3375 Then
        Exit Sub
    End If
    
End Sub

Public Function OpenMainForm(frm As Form)
        
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns
    Dim SetFocus, IsKeyVisible, QueryName, OnFormCreate, SubformName, UserQueryFields, Timestamp, CreatedBy
    
    ModelID = frm("ModelID")
    If ExitIfTrue(IsNull(ModelID), "Please select a record..") Then Exit Function
    
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
    Timestamp = frm("Timestamp")
    CreatedBy = frm("CreatedBy")
    
    ''Open the mainForm of the record selected
    Dim MainFormName
    ''Check first if SubformName is not Null
    If Not IsNull(SubformName) Then
        MainFormName = concat("main", SubformName)
GoTo OpenMainForm:
    End If
    
    If Not IsNull(VerbosePlural) Then
        MainFormName = concat("main", RemoveSpaces(VerbosePlural))
    Else
        MainFormName = concat("main", Model, "s")
    End If
    
OpenMainForm:
    DoCmd.OpenForm MainFormName
    
End Function

Public Function CreateDEUploadForm(frm As Object, ModelID)
    
    Dim ModelFieldID
    ModelFieldID = ELookup("qryModelFieldProperties", "ModelID = " & ModelID & " And Property = " & EscapeString("imageType"), "ModelFieldID")
    
    If ModelFieldID = "" Then Exit Function
    
    Dim modelFieldRS As Recordset
    Set modelFieldRS = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & ModelFieldID)
    
    Dim x, y: x = GetMaxX(frm): y = 600
    
    ''Render the image control
    Dim ctl As control
    Set ctl = CreateFlexControl(frm.Name, acImage, , , , x + 200, y, 3000, 3000)
    ctl.Name = concat(modelFieldRS.fields("ModelField"), "Img")
    SetControlPropertiesFromTemplate ctl, frm
    ctl.Picture = "placeholder"
    
    Dim ControlCaption
    If IsNull(modelFieldRS.fields("VerboseName")) Then
        ControlCaption = AddSpaces(modelFieldRS.fields("ModelField"))
    Else
        ControlCaption = modelFieldRS.fields("VerboseName")
    End If
    
    Set ctl = CreateFlexControl(frm.Name, acCommandButton, , , , x + 200, y - 300)
    SetControlPropertiesFromTemplate ctl, frm
    CopyControlTemplateProperties frm, ctl.Name
    ctl.Name = concat("cmd", modelFieldRS.fields("ModelField"))
    ctl.Caption = concat("Upload ", ControlCaption)
    ctl.OnClick = "=UploadImage([Form]," & EscapeString(modelFieldRS.fields("ModelField")) & ")"
    ctl.Height = 300
    ctl.Width = 3000
    
    ''Also render the textbox control of the ModelField
    Set ctl = CreateFlexControl(frm.Name, acTextBox, , , modelFieldRS.fields("ModelField"), 0, 0, 3000, 3000)
    SetControlPropertiesFromTemplate ctl, frm
    CopyControlTemplateProperties frm, ctl.Name
    ctl.Name = modelFieldRS.fields("ModelField")
    ctl.Visible = False
    
End Function

Public Function FollowFormHyperlink(frm, fieldName, Optional WithStreetAddress As Boolean = False, Optional AbsoluteLink = Null)
    
    Dim fileName, PropertyListID
    If isFalse(AbsoluteLink) Then fileName = frm(fieldName)
    If WithStreetAddress Then PropertyListID = frm("PropertyListID")
    
    If IsNull(fileName) Then
        ShowError "The hyperlink is empty..."
        Exit Function
    End If
    
    Dim assetDir, fs As Object, uploadDirectory
    
    uploadDirectory = GetAttachmentsDirectory
    
    If WithStreetAddress Then
        Dim StreetAddress
        StreetAddress = ELookup("tblPropertyList", "PropertyListID = " & PropertyListID, "StreetAddress")
        uploadDirectory = GetAttachmentsDirectory(StreetAddress)
    End If
    
    assetDir = uploadDirectory
    
    ''Override assetDir value with AbsoluteLink if it's not Null
    If Not isFalse(AbsoluteLink) Then
        assetDir = AbsoluteLink
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim filePath
    filePath = concat(assetDir, fileName)
    
    If Not fs.fileExists(filePath) Then
        MsgBox "File does not exist at: " & EscapeString(filePath)
        Exit Function
    End If
    
    On Error Resume Next
    FollowHyperlink filePath
    
End Function

Public Function SelectDirectory(frm As Object, ModelField)
    
    Dim fs As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPath
    With FileDialog(msoFileDialogFolderPicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Show the dialog box
        If ExitIfTrue(Not .Show, "Directory Selection cancelled..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        
    End With
    
    frm(ModelField) = fullPath
    
End Function

Public Function UploadMultiFile(frm As Object, ModelField, Optional WithStreetAddress As Boolean = False)

    Dim FileType, EntityID, PropertyListID
    FileType = frm("FileType")
    EntityID = frm("EntityID")
    PropertyListID = frm("PropertyListID")
    
    If ExitIfTrue(isFalse(FileType), "Select a valid file type..") Then Exit Function
    If ExitIfTrue(isFalse(EntityID), "One of the required fields is empty..") Then Exit Function
    
    Dim uploadDirectory, strFolderExists
    
    uploadDirectory = GetAttachmentsDirectory
    
    If WithStreetAddress Then
        Dim StreetAddress
        StreetAddress = ELookup("tblPropertyList", "PropertyListID = " & PropertyListID, "StreetAddress")
        uploadDirectory = GetAttachmentsDirectory(StreetAddress)
    End If
    
   
    strFolderExists = Dir(uploadDirectory, vbDirectory)
    
    ''Create the directory if it doesn't exist
    If strFolderExists = "" Then
        MkDir uploadDirectory
    End If
    
    Dim fs As Object, filePath, fileName
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim FileSelected
    With FileDialog(msoFileDialogFilePicker)
        'This will allow multi file selection
        .AllowMultiSelect = True
        .InitialFileName = uploadDirectory
        .filters.Clear
        If ExitIfTrue(Not .Show, "File upload cancelled..") Then Exit Function
        
        For Each FileSelected In .SelectedItems
            
           filePath = FilenameToBeUploaded(FileSelected, uploadDirectory, fileName)
           
           fs.CopyFile FileSelected, filePath, True
        
           InsertToEntityFiles fileName, FileType, EntityID, PropertyListID
           
           frm(ModelField) = filePath
            
        Next
        
    End With

End Function

Private Function InsertToEntityFiles(filePath, FileType, EntityID, PropertyListID)
    
    RunSQL "DELETE FROM tblEntityFiles WHERE EntityFileLink = " & EscapeString(filePath)
    RunSQL "INSERT INTO tblEntityFiles (EntityID,FileType,EntityFileLink,PropertyListID) VALUES (" & EntityID & "," & EscapeString(FileType) & "," & EscapeString(filePath) & "," & PropertyListID & ")"
    
End Function

Private Function FilenameToBeUploaded(FileSelected, uploadDirectory, fileName)
    
    With CreateObject("Scripting.FileSystemObject")
    
        Dim extName, baseName, fileExists
        fileName = .GetFileName(FileSelected)
        extName = .GetExtensionName(FileSelected)
        baseName = .GetBaseName(FileSelected)
        
        FilenameToBeUploaded = uploadDirectory & fileName
        
'        '''Check if the file already exists to the upload directory
'        fileExists = Dir(FilenameToBeUploaded)
'
'        If fileExists <> "" Then
'            ''Change the file name
'            fileName = baseName & Format(Now(), "_yyyy_mm_dd_hh_MM_ss") & "." & extName
'            FilenameToBeUploaded = uploadDirectory & fileName
'        End If
 
    End With
    
    Debug.Print FilenameToBeUploaded
    
End Function

Public Function GetAttachmentsDirectory(Optional StreetAddress = Null)
    
    GetAttachmentsDirectory = CurrentProject.path & "\Files\"
    If Environ("computername") <> "LAPTOP-4EL19IO4" Then GetAttachmentsDirectory = "Z:\MY PANDA APP\Attachments\"
    
    If Not isFalse(StreetAddress) Then
        StreetAddress = replace(StreetAddress, "\", " ")
        StreetAddress = replace(StreetAddress, "/", " ")
        GetAttachmentsDirectory = GetAttachmentsDirectory & StreetAddress & "\"
    End If
    
End Function

Public Function UploadFile(frm As Object, ModelField)
    
    Dim assetDir, fs As Object
    assetDir = GetApplicationSetting("Asset Directory")
    
    If assetDir = "" Then assetDir = CurrentProject.path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(assetDir) Then assetDir = CurrentProject.path
    
    Dim fullPath, fileName
    With FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        .filters.Clear
        'Show the dialog box
        If ExitIfTrue(Not .Show, "File upload cancelled..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        fileName = fs.GetFileName(fullPath)
        
    End With
    
    fs.CopyFile fullPath, concat(assetDir, "\", fileName), True
    
    frm(ModelField) = fileName
    
End Function

Public Function GetFilePath(frm As Object, Optional ModelField = "") As String
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPath, fileName
    With FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        .filters.Clear
        'Show the dialog box
        If ExitIfTrue(Not .Show, "File selection cancelled..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        fileName = fs.GetFileName(fullPath)
        
    End With
    
    If Not isFalse(ModelField) Then frm(ModelField) = fullPath
    GetFilePath = fullPath
    
End Function

Public Function UploadImage(frm As Object, ModelField)
    
    Dim assetDir, fs As Object
    assetDir = GetApplicationSetting("Asset Directory")
    
    If assetDir = "" Then assetDir = CurrentProject.path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(assetDir) Then assetDir = CurrentProject.path
    
    Dim fullPath, fileName
    With FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .filters.Add "Image Files", "*.jpg; *.png", 1
        'Show the dialog box
        If ExitIfTrue(Not .Show, "File upload cancelled..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        fileName = fs.GetFileName(fullPath)
        
    End With
    
    fs.CopyFile fullPath, concat(assetDir, "\", fileName), True
    
    frm(ModelField) = fileName
    frm(concat(ModelField, "Img")).Picture = concat(assetDir, "\", fileName)
    
End Function

Public Function CreateDEFileUploadControl(frm As Object, fldName, ByVal x, ByVal y, fldWidth, Optional asFilePath As Boolean = False)
    
    Dim proportionArr As New clsArray, controlArr As New clsArray, proportionTotal, totalWidth, colSpaceWidth, i, proportion
    
    colSpaceWidth = 50
    totalWidth = fldWidth
    
    Dim ctl As control
    Set ctl = CreateControl(frm.Name, acLabel, , fldName, "Select " & AddSpaces(fldName), x, y - 300)
    SetControlPropertiesFromTemplate ctl, frm
    ctl.Width = totalWidth
    
    Set ctl = CreateControl(frm.Name, acTextBox, , , fldName, 0, 0, 0, 300) ''Texbox Portion
    SetControlPropertiesFromTemplate ctl, frm
    ctl.Name = fldName
    ctl.OnClick = "=FollowFormHyperlink([Form]," & EscapeString(fldName) & ")"
    
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0) ''Button Portion
    SetControlPropertiesFromTemplate ctl, frm
    ctl.Name = concat("cmd", fldName)
    ctl.Caption = "Browse..."
    ctl.OnClick = IIf(asFilePath, "=GetFilePath([Form]," & EscapeString(fldName) & ")", "=UploadFile([Form]," & EscapeString(fldName) & ")")
    ctl.Height = 300
    
    ''Render the Filter buttons
    ''Filter and Clear
    proportionArr.arr = "10,2"
    controlArr.arr = fldName & "," & concat("cmd", fldName)
    proportionTotal = GetProportionTotal(proportionArr)
    
    For i = 0 To proportionArr.count - 1
    
        proportion = CDbl(proportionArr.arr(i)) / proportionTotal
        frm(controlArr.arr(i)).Left = x
        frm(controlArr.arr(i)).Top = y
        frm(controlArr.arr(i)).Width = (totalWidth - ((proportionArr.count - 1) * colSpaceWidth * 2)) * proportion
        frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
       
        x = x + (colSpaceWidth * 2) + frm(controlArr.arr(i)).Width
       
    Next i
    
    
End Function

Private Function GetTableOrQueryDef(rsName, db As Database) As Object
    Dim def As Object, fld As field
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set GetTableOrQueryDef = db.TableDefs(rsName)
    Else
        Set GetTableOrQueryDef = db.QueryDefs(rsName)
    End If
End Function

''SimpleOnly flag is useful when creating a simple data entry form without any children to be used by NotInList event of comboboxes.
Public Function CreateDEForm(frm2 As Form, Optional CreateAsReport As Boolean = False, Optional SimpleOnly As Boolean = False)
    
    Dim rs As Recordset, fldName
    Dim maxWidth As Double
    Dim ctl As control
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible
    Dim QueryName, OnFormCreate, SubformName2, UserQueryFields, Timestamp, CreatedBy, PrimaryKey
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    SubformName2 = frm2("SubformName")
    UserQueryFields = frm2("UserQueryFields")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    PrimaryKey = frm2("PrimaryKey")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    ''This will insert the fields not present from the model field table but present from the recordset
    GenerateFields frm2
    
    '''Create the controls
    '''Fields, default buttons, additional buttons and controls
    ''Create the form
    Dim frm As Object
    If CreateAsReport Then
        Set frm = CreateReport
        SetCommonReportProperties frm
    Else
        Set frm = CreateForm
        If SimpleOnly Then frm.DataEntry = True
    End If
    
    Dim rsName: rsName = GetTableName(Model, VerbosePlural)
    ''Override the rsName if the QueryName isn't empty
    If Not IsNull(QueryName) Then rsName = QueryName
    Dim tblName: tblName = rsName
    
    ''Set the form's property
    frm.recordSource = rsName
    Dim VerboseCaption: VerboseCaption = ELookup("tblSupplementalModels", "ModelID = " & ModelID, "VerboseCaption")
    Dim frmCaption: frmCaption = GetFieldCaption(VerboseName, Model, VerboseCaption)
    
    frm.Caption = concat(frmCaption, IIf(CreateAsReport, " Report", " Form"))
    If Not CreateAsReport Then frm.OnCurrent = "=SetFocusOnForm([Form],""" & SetFocus & """)"
    
    ''If the PrimaryKey is null then the PrimaryKey will come from the Model
    If IsNull(PrimaryKey) Then PrimaryKey = GetPrimaryKeyFromTable(ModelID)
    If Not CreateAsReport Then
        frm.BeforeUpdate = "=SaveFormData2([Form]," & Esc(Model) & ")"
        frm.OnLoad = "=DefaultFormLoad([Form]," & Esc(PrimaryKey) & ")"
    End If
    
    If Not CreateAsReport Then SetFormProperties 4, frm
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    
    frmName = frm.Name
    
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("frm", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("frm", Model, "s")
    End If
    
    If Not IsNull(SubformName2) Then
        baseFormName = concat("frm", SubformName2)
    End If
    
    If CreateAsReport Then
        baseFormName = replace(baseFormName, "frm", "rpt")
    End If
    
    ''If Not CreateAsReport Then RenderFilterForm frm, ModelID, True, baseFormName
    
    Dim db As Database: Set db = CurrentDb
    Dim rsObj As Object: Set rsObj = GetTableOrQueryDef(rsName, db)
    Dim x, y: x = 400: y = GetMaxY(frm) + InchToTwip(0.5)
    Dim CurrentCol: CurrentCol = 0:
    Dim isMemo As Boolean
    
    ''Render the non boolean fields first.. I think boolean fields should be rendered separately
    ''dbBoolean is 1, btw as FieldTypeID
    Dim sqlStr As String: sqlStr = "SELECT * FROM tblModelFields WHERE FieldOrder IS NOT NULL AND ModelID = " & ModelID & " AND FieldTypeID <> 1"
    ''Get only the native fields from the table if Not UserQueryFields
    If Not UserQueryFields Then sqlStr = sqlStr & " AND FieldSource = " & EscapeString(rsName)
    sqlStr = sqlStr & " ORDER BY FieldOrder ASC"
    Set rs = ReturnRecordset(sqlStr)
    
    ''Loop through the non boolean fields
    Dim fldWidth, fld As field, ControlHeight
    
    
    Do Until rs.EOF
        ControlHeight = rs.fields("ControlHeight")
        If isFalse(ControlHeight) Then ControlHeight = 3000
        ''Skip the field if there is an imageType proprety on qryModelFieldProperties
        If isPresent("qryModelFieldProperties", "Property = " & Esc("imageType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            GoTo NextField
        End If
    
        ''Look First if the ControlSource is not Null
        If Not IsNull(rs.fields("ControlSource")) Then
            ''If onDatasheetOnly property is on then skip this for the data entry form
            If isPresent("qryModelFieldProperties", "Property = ""onDatasheetOnly"" And ModelFieldID = " & rs.fields("ModelFieldID")) Then
                GoTo NextField
            Else
                ''Create a control at the footer of the form
                fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
                Set ctl = CreateFlexControl(frm.Name, acTextBox, , "", "", x, y, fldWidth)
                SetControlPropertiesFromTemplate ctl, frm
                ctl.Name = rs.fields("ModelField")
                ctl.ControlSource = rs.fields("ControlSource")
                
                Select Case rs.fields("FieldTypeID")
                    Case dbMemo:
                        ctl.Height = ControlHeight
                        ctl.EnterKeyBehavior = 1
                        isMemo = True
                    Case dbDouble
                        ctl.Format = "Standard"
                End Select
            
                ''Generate the label just above the control
                Dim CustomVerboseName
                If IsNull(rs.fields("VerboseName")) Then
                    CustomVerboseName = AddSpaces(rs.fields("ModelField"))
                Else
                    CustomVerboseName = rs.fields("VerboseName")
                End If
                Set ctl = CreateFlexControl(frm.Name, acLabel, , rs.fields("ModelField"), CustomVerboseName, x, y - 300)
                SetControlPropertiesFromTemplate ctl, frm
                ctl.Width = fldWidth
                
                GoTo SetVariables
            End If
        End If
    
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"), True)
        fldWidth = 3000
        
        If Not DoesPropertyExists(rsObj.fields, fldName) Then GoTo NextField
        
        Set fld = rsObj.fields(fldName)
        
        ''This skips the primary key
        If Not IsKeyVisible And fld.Name = PrimaryKey Then
            Select Case fld.Type
                Case dbText, dbDate:
                Case Else:
                    GoTo NextField
            End Select
        End If
  
        Select Case fld.Name
            Case "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
        Dim ControlTypeValue
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeValue = acTextBox
        Else
            ControlTypeValue = fld.Properties("DisplayControl")
        End If
        
        ''Generate the control first before the label
        ''Get the width depending on the Number of columns but make sure that the CurrentCol + the columns will not exceed
        ''the FormColumn
        fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
        
        ''Setting the FieldOrder to 0 here will make the fld hidden..
        If rs.fields("FieldOrder") = 0 Then
            Set ctl = CreateFlexControl(frm.Name, ControlTypeValue, , "", fld.Name, 0, 0, 0)
            ctl.Name = fld.Name
            ctl.Visible = False
            SetControlPropertiesFromTemplate ctl, frm
            GoTo NextField
        End If
        
        ''Check if the current field has fileType property. This will tell us that we need to use an upload form rather than a memo field
        ''Skip the field if there is an imageType proprety on qryModelFieldProperties
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("fileType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            CreateDEFileUploadControl frm, fld.Name, x, y, fldWidth
            GoTo SetVariables
        End If
        
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("filePath") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            CreateDEFileUploadControl frm, fld.Name, x, y, fldWidth, True
            GoTo SetVariables
        End If
        
        ''Path Only Here
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("folderType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            CreateDEFolderControl frm, fld.Name, x, y, fldWidth
            GoTo SetVariables
        End If
        
        Set ctl = CreateFlexControl(frm.Name, ControlTypeValue, , "", fld.Name, x, y, fldWidth)
        ctl.Name = fld.Name
        
        ''Set control property based on ControlTypeValue
        SetControlPropertiesFromTemplate ctl, frm
        
        Select Case fld.Type
            Case dbMemo:
                ctl.Height = ControlHeight
                ctl.EnterKeyBehavior = 1
                isMemo = True
            Case dbDouble:
                ctl.Format = "Standard"
        End Select
        
        ''Generate the label just above the control
        Dim ControlCaption
        If Not IsNull(rs.fields("VerboseName")) Then
            ControlCaption = rs.fields("VerboseName")
        Else
            If Not DoesPropertyExists(fld.Properties, "Caption") Then
                ControlCaption = AddSpaces(rs.fields("ModelField"))
            Else
                ControlCaption = fld.Properties("Caption")
            End If
        End If
'        If Not DoesPropertyExists(fld.Properties, "Caption") Then
'            If IsNull(rs.fields("VerboseName")) Then
'                ControlCaption = AddSpaces(rs.fields("ModelField"))
'            Else
'                ControlCaption = rs.fields("VerboseName")
'            End If
'        Else
'            ControlCaption = fld.Properties("Caption")
'        End If
        
        Set ctl = CreateFlexControl(frm.Name, acLabel, , fld.Name, ControlCaption, x, y - 300)
        SetControlPropertiesFromTemplate ctl, frm
        ctl.Name = concat("lbl", fld.Name)
        ctl.Width = fldWidth

SetVariables:
    
        CurrentCol = CurrentCol + rs.fields("Columns")
        
        If CurrentCol = FormColumns Then
            
            If x + 3000 + 400 > maxWidth Then
                maxWidth = x + 3000 + 400
            End If
            
            CurrentCol = 0
            x = 400
            If Not isMemo Then
                y = y + 700
            Else
                isMemo = False
                y = y + ControlHeight + 600
            End If
            
        Else
        
            x = x + (3200 * rs.fields("Columns"))
        End If
NextField:
        
        rs.MoveNext
    Loop
    
    ''Render the boolean fields here-->
    RenderBooleanFieldsToDEForm frm, ModelID
    
    ''Create the Timestamp and CreatedBy field (Hidden Fields)
    If Not CreateAsReport Then
        Set ctl = CreateFlexControl(frm.Name, acTextBox, , "", "Timestamp", 0, 0, 0)
        ctl.Name = "Timestamp"
        SetControlPropertiesFromTemplate ctl, frm
        
        Set ctl = CreateFlexControl(frm.Name, acComboBox, , "", "CreatedBy", 0, 0, 0)
        ctl.Name = "CreatedBy"
        SetControlPropertiesFromTemplate ctl, frm
        
        frm("Timestamp").Visible = False
        frm("CreatedBy").Visible = False
    End If

    ''Create Upload Form if there's any (Simple PictureBox) + Button at the bottom
    CreateDEUploadForm frm, ModelID
    
    ''Create the child forms here
    ''Child sub + Plural name of the Child Model
    x = 400
    y = y + 600
    
    
    fldWidth = GetMaxX(frm) - 400
    Dim minimumWidth: minimumWidth = (FormColumns * 3000)
    If minimumWidth > fldWidth Then fldWidth = minimumWidth
    
    
    If SimpleOnly Then
        GoTo EndOfSubformRendering:
    End If
    
    Dim childModels
    Set rs = ReturnRecordset("SELECT * " & _
        "FROM qryModelFields WHERE ParentModelID = " & ModelID & " And HideSubformFromParent = 0 " & _
        "ORDER BY SubPageorder ASC")
        
    If rs.EOF Then
        childModels = 0
    Else
        rs.MoveLast
        rs.MoveFirst
        childModels = rs.recordCount
    End If
    
    Dim pg As Page, tbCtl As TabControl, pgCaption, x1, y1 As Long, SubformName, SubModel, maxY, pgName, subTblName, ModelFieldID, maxX As Long
    
    Do Until rs.EOF
        
        If Not DoesPropertyExists(frm, "tabCtl") Then
        
            For Each ctl In frm.controls
                If ctl.Top + ctl.Height > maxY Then
                    maxY = ctl.Top + ctl.Height
                End If
            Next ctl
            
            Dim tabCtlHeight: tabCtlHeight = 7000
            Set ctl = CreateFlexControl(frm.Name, acTabCtl, , , , x, maxY + 400, fldWidth, tabCtlHeight)
            ctl.Name = "tabCtl"
            SetControlPropertiesFromTemplate ctl, frm
            
            'frm.Width = (FormColumns * 3000) + (FormColumns * 400) - 200
        End If
        
        Set tbCtl = frm.tabCtl
        
        Do Until tbCtl.Pages.count > childModels
            tbCtl.Pages.Add
        Loop
        
        For Each pg In tbCtl.Pages
                
            If pg.PageIndex > childModels - 1 Then
                tbCtl.Pages.Remove pg.PageIndex
            Else
                If Not IsNull(rs.fields("VerbosePlural")) Then
                    pgCaption = AddSpaces(rs.fields("VerbosePlural"))
                    SubModel = concat(replace(rs.fields("VerbosePlural"), " ", ""))
                    pgName = concat("pg", SubModel)
                    subTblName = concat("tbl", SubModel)
                Else
                    pgCaption = AddSpaces(concat(rs.fields("Model"), "s"))
                    SubModel = concat(rs.fields("Model"), "s")
                    pgName = concat("pg", SubModel)
                    subTblName = concat("tbl", SubModel)
                End If
                
                If Not IsNull(rs.fields("VerboseChildName")) Then
                    pgCaption = rs.fields("VerboseChildName")
                    SubModel = RemoveSpaces(pgCaption)
                    pgName = concat("pg", SubModel)
                End If
                
                If Not isFalse(rs.fields("VerbosePluralCaption")) Then
                    pgCaption = rs.fields("VerbosePluralCaption")
                End If
                
                pg.Caption = pgCaption
                pg.Name = pgName
                
                ''Add the Buttons
                maxY = frm.controls("tabCtl").Top - 400
                x1 = 600
                y1 = maxY + 400 + 500
                
                ModelFieldID = rs.fields("ModelFieldID")
                
                Dim frmToOpen
                frmToOpen = GetModelPlural(rs.fields("Model"), rs.fields("VerbosePlural"), "frm")
                
                ''New Button
                
                If Not isPresent("qryModelFieldProperties", "Property = ""pgDeleteNewEditHidden"" And ModelFieldID = " & ModelFieldID) Then
                    If Not isPresent("qryModelFieldProperties", "Property = ""pgNewHidden"" And ModelFieldID = " & ModelFieldID) Then
                        RenderButton x1, y1, "New", frm, concat("Add", SubModel), pg.Name
                        frm(concat("cmdAdd", SubModel)).OnClick = "=OpenFormFromMain(" & EscapeString(frmToOpen) & ","""","""",[Form]," & EscapeString(PrimaryKey) & ")"
                        x1 = x1 + (3200 * 0.46)
                    End If
                    
                    ''Edit/View Button
                    If Not isPresent("qryModelFieldProperties", "Property = ""pgEditHidden"" And ModelFieldID = " & ModelFieldID) Then
                        RenderButton x1, y1, "Edit/View", frm, concat("Edit", SubModel), pg.Name
                        frm(concat("cmdEdit", SubModel)).OnClick = "=OpenFormFromMain(" & EscapeString(frmToOpen) & ", " & EscapeString(concat("sub", SubModel)) & ", " & _
                                EscapeString(concat(rs.fields("Model"), "ID")) & ",[Form])"
                        x1 = x1 + (3200 * 0.46)
                        
                    End If
                    
                    ''Delete Button
                    If Not isPresent("qryModelFieldProperties", "Property = ""pgDeleteHidden"" And ModelFieldID = " & ModelFieldID) Then
                        RenderButton x1, y1, "Delete", frm, concat("Delete", SubModel), pg.Name
                        frm(concat("cmdDelete", SubModel)).OnClick = "=DeleteRecord([Form], " & EscapeString(concat(rs.fields("Model"), "ID")) & ", " & _
                                EscapeString(subTblName) & "," & EscapeString(concat("sub", SubModel)) & ")"
                        x1 = x1 + (3200 * 0.46)
                    End If
                End If
                
                RenderChildModelButtons frm, pg.Name, rs.fields("ModelID"), y1, SubModel
                ''Insert the buttons here.. a property should be enabled so that other pages wouldn't render their own button
                ''Use this function format: =RunFunctionFromSubform([Form],"subform","OpenButtonModule")
                ''Position the button at the leftmost + ?
                ''Property name is pgRenderAdditionalButtons - ModelFieldProperties
                ''TABLE: qryModelFieldProperties Fields: ModelFieldPropertyID | ModelFieldID | PropertyID | Property | PropertyDescription
                '' Timestamp | CreatedBy | ModelID
                ''TABLE: tblModelFields -> ParentModelID is equal the current ModelID
                ''TABLE: tblModelButtons Fields: ModelButtonID | ModelID | ModelButton | FunctionName | TableWideFunction
                '' Timestamp | CreatedBy | ModelButtonOrder | HideOnMain | HideOnForm | RecordImportID | Note | TemplateName
                
                ''Get the x + length of the leftmost button
                maxY = frm.controls("tabCtl").Top - 400
                For Each ctl In frm.controls
                    If ctl.ControlType = acCommandButton Then
                        If ctl.Top + ctl.Height > maxY Then
                            maxY = ctl.Top + ctl.Height
                        End If
                    End If
                Next ctl
                
                If x1 > 600 Then
                    y1 = y1 + 400 + 100
                End If
                
                x1 = 600
                
                Dim pgCtlHeight, pgCtlTop
                pgCtlHeight = frm.controls("tabCtl").Height
                pgCtlTop = frm.controls("tabCtl").Top
                pgCtlHeight = pgCtlTop + pgCtlHeight - y1 - 200
                
                Set ctl = CreateFlexControl(frm.Name, acSubform, , concat("pg", SubModel), , x1, y1, fldWidth - 400, pgCtlHeight)
                ctl.Name = concat("sub", SubModel)
                ctl.Properties("RightPadding") = 100
            
        
                If Not IsNull(rs.fields("SubformSource")) Then
                    ctl.SourceObject = rs.fields("SubformSource")
                Else
                    ctl.SourceObject = "dsht" & SubModel
                End If
                
                ctl.HorizontalAnchor = acHorizontalAnchorBoth
                ctl.VerticalAnchor = acVerticalAnchorBoth
                
                ''Join the subform using the PrimaryKey
                ctl.LinkMasterFields = PrimaryKey
                ctl.LinkChildFields = rs.fields("ModelField")
                
                ''Option button goes here ===>

                ''GenerateAdditionalOptionButton frm, ModelFieldID, Concat("sub", subModel), pg.Name
                
                            
            End If
            
            If Not rs.EOF Then
            
                rs.MoveNext
            
            End If
            
        Next pg
        
        
    Loop
    
    If childModels > 0 Then
        y = y + 7000
    End If

    ''Any subform totals will be placed after the subform
    Set rs = ReturnRecordset("SELECT * FROM tblSubformControls WHERE IsVisible = -1 And ModelID = " & ModelID & " ORDER BY FieldOrder ASC")
    
    ''This will loop inside the tblSubformControls for this model
    If Not rs.EOF Then
        maxY = 0
        CurrentCol = 0
        For Each ctl In frm.controls
            If (ctl.Top + ctl.Height) > maxY Then
                maxY = ctl.Top + ctl.Height
            End If
        Next ctl
        
        x = 400
        y = maxY + 800
        
        Dim ctlName
        
        Do Until rs.EOF
            
            ctlName = concat(rs.fields("SubformName"), rs.fields("ControlName"))
            
            fldWidth = 3000 * (0.5) + (200 * (0.5 - 1))
            Set ctl = CreateFlexControl(frm.Name, acTextBox, , "", "", x, y, fldWidth)
            ctl.Name = ctlName
            ctl.ControlSource = concat("=IfError(", rs.fields("SubformName"), "!SUM", rs.fields("ControlName"), ")")
            ctl.Format = "Standard"
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
            
            ''Set control property based on ControlTypeValue
            SetControlPropertiesFromTemplate ctl, frm
            
            ''Generate the label just above the control
            Set ctl = CreateFlexControl(frm.Name, acLabel, , ctl.Name, rs.fields("ControlCaption"), x, y - 500)
            
            SetControlPropertiesFromTemplate ctl, frm
            
            ctl.Height = 400
            ctl.Width = fldWidth
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
            CurrentCol = CurrentCol + 0.5
        
            If CurrentCol = FormColumns Then
                
                CurrentCol = 0
                x = 400
                y = y + 900
            Else
            
                x = x + (3200 * 0.5)
                
            End If
            
            
            rs.MoveNext
        Loop
    End If

EndOfSubformRendering:
    ''Buttons
    maxY = 0
    For Each ctl In frm.controls
        If (ctl.Top + ctl.Height) > maxY Then
            maxY = ctl.Top + ctl.Height
        End If
    Next ctl
    
    x = 400
    y = GetMaxY(frm) + 400
    
    Dim buttonMultiplier
    buttonMultiplier = 0.46
    CurrentCol = 0
    ''Cancel Button
    
    If Not CreateAsReport Then
        If Not isPresent("qryModelProperties", "Property = ""frmCancelNewSaveDeleteHidden"" And ModelID = " & ModelID) Then
            If Not isPresent("qryModelProperties", "Property = ""frmCancelHidden"" And ModelID = " & ModelID) Then
                RenderButton x, y, "Cancel", frm, "Cancel"
                frm.cmdCancel.OnClick = "=CancelEdit([Form])"
                frm.cmdCancel.HorizontalAnchor = 0
                frm.cmdCancel.VerticalAnchor = 1
                x = x + (3200 * buttonMultiplier)
                CurrentCol = CurrentCol + 0.5
            End If
            
            ''New Button
            If Not isPresent("qryModelProperties", "Property = ""frmNewHidden"" And ModelID = " & ModelID) Then
                RenderButton x, y, "New", frm, "New"
                frm.cmdNew.OnClick = "=Save2([Form],'" & Model & "',0)"
                frm.cmdNew.HorizontalAnchor = 0
                frm.cmdNew.VerticalAnchor = 1
                x = x + (3200 * buttonMultiplier)
                CurrentCol = CurrentCol + 0.5
            End If
            
            ''Save Button
            If Not isPresent("qryModelProperties", "Property = ""frmSaveHidden"" And ModelID = " & ModelID) Then
                RenderButton x, y, "Save", frm, "SaveClose"
                frm.cmdSaveClose.OnClick = "=Save2([Form],'" & Model & "',1)"
                frm.cmdSaveClose.HorizontalAnchor = 0
                frm.cmdSaveClose.VerticalAnchor = 1
                x = x + (3200 * buttonMultiplier)
                CurrentCol = CurrentCol + 0.5
            End If
            
            ''Delete Button
            If Not isPresent("qryModelProperties", "Property = ""frmDeleteHidden"" And ModelID = " & ModelID) Then
                RenderButton x, y, "Delete", frm, "Delete"
                frm.cmdDelete.OnClick = "=DeleteRecord([Form], '" & PrimaryKey & "', '" & rsName & "')"
                frm.cmdDelete.HorizontalAnchor = 0
                frm.cmdDelete.VerticalAnchor = 1
                x = x + (3200 * buttonMultiplier)
                CurrentCol = CurrentCol + 0.5
            End If
        End If
    End If

    maxX = x
    
    RenderAdditionalButtonOnDEForm ModelID, frm, y, maxX, CurrentCol, buttonMultiplier, FormColumns
    
    ''Align the collapsed controls if there's any
    If DoesPropertyExists(frm.controls, "cboFormActions") Then
        maxX = GetMaxX(frm)
        frm("cmdRunFormActions").Left = maxX - frm("cmdRunFormActions").Width
        frm("cboFormActions").Left = frm("cmdRunFormActions").Left - 55 - frm("cboFormActions").Width
        frm("lblFormActions").Left = frm("cboFormActions").Left - 55 - frm("lblFormActions").Width
    End If
    

    ''Set the Form Width
    maxX = 0
    For Each ctl In frm.controls
        If ctl.Left + ctl.Width > maxX Then
            maxX = ctl.Left + ctl.Width
        End If
    Next ctl
    
    frm.Width = GetMaxX(frm) + 400
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, frm, 4
    End If
    
    frm.Section("Detail").Height = GetMaxY(frm) + 400
    
    If CreateAsReport Then
        CleanUpReportProperties frm
    End If
    
    ''Override Properties Here
    OverrideProperties ModelID, 4, frm
    
    DoCmd.Close IIf(CreateAsReport, acReport, acForm), frm.Name, acSaveYes
    
    If SimpleOnly Then
        baseFormName = RegExReplace(baseFormName, "^frm", "frmSimple")
    End If
    
    customFrmName = baseFormName
    
'    Do Until Not FrmExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    InsertToModelRelatedObjects ModelID, acForm, customFrmName
    
    DoCmd.Rename customFrmName, IIf(CreateAsReport, acReport, acForm), frmName
    
    ''Load Save Form Layout
    LoadSavedFormLayout customFrmName
    
    If CreateAsReport Then
        DoCmd.OpenReport customFrmName, acViewDesign
    Else
        DoCmd.OpenForm customFrmName
    End If

    InsertFormInFormForRights customFrmName, Model
    
End Function

Private Function GetFromSupplementalModels(ModelID, PropName)
    ''PropName can be BooleanWidth, BooleanProportion
    ''See if supplemental property for this model has been set if not get the default value
    ''using tblModelFields -> DefaultValue property
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblSupplementalModels WHERE ModelID = " & ModelID)
    
    Do Until rs.EOF
        GetFromSupplementalModels = rs.fields(PropName)
        Exit Do
        rs.MoveNext
    Loop
    
    If isFalse(GetFromSupplementalModels) Then
        ''624 is the ModelID of SupplementalModel
        GetFromSupplementalModels = ELookup("tblModelFields", "ModelID = 624 And ModelField = " & Esc(PropName), "DefaultValue")
    End If
    
End Function

Private Function RenderBooleanFieldsToDEForm(frm As Object, ModelID)
    
    Dim isReport: isReport = IsObjectAReport(frm)
    ''Get the starting x and y positions
    ''x would be the maximum x value from the form
    ''y would be the starting position of controls from the form vertically -> would be computed for each control
    Dim Width: Width = GetFromSupplementalModels(ModelID, "BooleanWidth") ''See the model and get the field from there
    Dim x: x = GetMaxX(frm) + 300
    
    ''TABLE: tblModelFields Fields: ModelFieldID|ModelID|ModelField|FieldTypeID|VerboseName|ValidationString
    ''ForeignKey|PossibleValues|DefaultValue|IsIndexed|FieldOrder|Columns|ColumnBreak|ColumnWidth|HideSubformFromParent
    ''ParentModelID|VerboseChildName|SubPageOrder|FieldSource|SubformSource|IsAnExpression|ControlSource|FieldFormat
    ''ReportFieldOrder|Timestamp|CreatedBy|RecordImportID

    Dim sqlStr: sqlStr = "SELECT * FROM tblModelFields WHERE ModelID = " & ModelID & " AND Not FieldOrder IS NULL " & _
        " AND FieldTypeID = 1 ORDER BY FieldOrder ASC"
        
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim proportions As New clsArray: proportions.arr = GetFromSupplementalModels(ModelID, "BooleanProportion")
    Dim ControlNames As New clsArray
    Do Until rs.EOF
    
        Set ControlNames = New clsArray
        Dim ModelField: ModelField = rs.fields("ModelField")
        Dim VerboseName: VerboseName = rs.fields("VerboseName")
        Dim labelCaption: labelCaption = GetFieldCaption(VerboseName, ModelField)
        Dim labelName: labelName = "lbl" & ModelField
        ''Get the fields -> Render the label and the checkbox
        Dim ctl As control
        Set ctl = CreateNamedControl(frm, acCheckBox, ModelField) ''--> checkbox first since label's parent will be the name of the cb
        ctl.Name = ModelField
        ctl.ControlSource = ModelField
        CopyProperties frm, ModelField, "CheckboxControl", False
        ''checkbox control source will be the fieldName -> name will be the fieldName
        ''label Caption will be the verbose name of the field -> label name would be "lbl" & fieldName
        Set ctl = CreateNamedControl(frm, acLabel, labelName, ModelField)
        ctl.Caption = labelCaption
        CopyProperties frm, labelName, "LabelControl", False
        ''Reposition controls
        ControlNames.Add ModelField
        ControlNames.Add labelName
        Dim y: y = GetMaxY(frm, , x, Width)
        RepositionControlsInRow frm, proportions.JoinArr, ControlNames.JoinArr, x, Width, , 600
        frm(ModelField).HorizontalAnchor = acLeft
        frm(labelName).HorizontalAnchor = acLeft
        
        rs.MoveNext
    Loop

End Function

Private Function RenderChildModelButtons(frm As Object, pgName, ModelID, y, SubModel)
    
    Dim x, btnName, cmdBtnName
    
    If isPresent("qryModelFieldProperties", "ModelID = " & ModelID & " AND Property = ""pgRenderAdditionalButtons""") Then
        ''Render this button
        ''Get the x and y axis of the control
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset("SELECT * FROM tblModelButtons WHERE ModelID = " & ModelID)
        Do Until rs2.EOF
            Dim ModelButton: ModelButton = rs2.fields("ModelButton"): If ExitIfTrue(isFalse(ModelButton), "ModelButton is empty..") Then Exit Function
            Dim FunctionName: FunctionName = rs2.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
            x = GetMaxX(frm, y, pgName)
            If x < 600 Then x = 600
            btnName = concat(FunctionName, SubModel)
            cmdBtnName = "cmd" & btnName
            RenderButton x + 60, y, ModelButton, frm, btnName, pgName
            frm(cmdBtnName).OnClick = "=RunFunctionFromSubform([Form]," & Esc("sub" & SubModel) & "," & Esc(FunctionName) & ")"
            rs2.MoveNext
        Loop
    
    End If
        
    
    ''Insert the buttons here.. a property should be enabled so that other pages wouldn't render their own button
    ''Use this function format: =RunFunctionFromSubform([Form],"subform","OpenButtonModule")
    ''Position the button at the leftmost + ?
    ''Property name is pgRenderAdditionalButtons - ModelFieldProperties
    ''TABLE: qryModelFieldProperties Fields: ModelFieldPropertyID | ModelFieldID | PropertyID | Property | PropertyDescription
    '' Timestamp | CreatedBy | ModelID
    ''TABLE: tblModelFields -> ParentModelID is equal the current ModelID
    ''TABLE: tblModelButtons Fields: ModelButtonID | ModelID | ModelButton | FunctionName | TableWideFunction
    '' Timestamp | CreatedBy | ModelButtonOrder | HideOnMain | HideOnForm | RecordImportID | Note | TemplateName
    ''Use this function RenderButton x1, y1, "Delete", frm, concat("Delete", SubModel), pg.name
'    RenderButton x1, y1, "Delete", frm, concat("Delete", SubModel), pg.name
'                        frm(concat("cmdDelete", SubModel)).OnClick = "=DeleteRecord([Form], " & EscapeString(concat(rs.fields("Model"), "ID")) & ", " & _
'                                EscapeString(subTblName) & "," & EscapeString(concat("sub", SubModel)) & ")"
    
End Function


Private Function GetFormType(customFrmName) As String
    
    Dim frmName
    frmName = customFrmName
    
    GetFormType = "DataEntry"
    
    If frmName Like "main*" Then
    
        GetFormType = "MainForm"
        Exit Function
        
    ElseIf frmName Like "dsht*" Then
            
        GetFormType = "DataSheet"
        Exit Function
            
    End If
    
End Function

Private Function InsertFormInFormForRights(customFrmName, Model)
    
    Dim frmType
    frmType = GetFormType(customFrmName)
    
    ''TABLE: tblFormForRights Fields: FormForRightsID|ModelName|FormName|FormType|Timestamp|CreatedBy|RecordImportID|ModelCaption
    
    RunSQL "INSERT INTO tblFormForRights (ModelName, FormName, FormType) VALUES ('" & Model & "','" & customFrmName & "','" & frmType & "')"
    
End Function

Public Function CreateDEFolderControl(frm As Object, fldName, ByVal x, ByVal y, fldWidth)
    
    Dim proportionArr As New clsArray, controlArr As New clsArray, proportionTotal, totalWidth, colSpaceWidth, i, proportion
    
    colSpaceWidth = 50
    totalWidth = fldWidth
    
    Dim ctl As control
    ''This is the label
    Set ctl = CreateControl(frm.Name, acLabel, , fldName, "Select " & AddSpaces(fldName), x, y - 300)
    SetControlPropertiesFromTemplate ctl, frm
    ctl.Width = totalWidth
    
    ''This is the textbox -> Need to add anchoring
    Set ctl = CreateControl(frm.Name, acTextBox, , , fldName, 0, 0, 0, 300) ''Texbox Portion
    SetControlPropertiesFromTemplate ctl, frm
    ctl.Name = fldName
    
    ''ctl.OnClick = "=FollowFormHyperlink([Form]," & EscapeString(fldName) & ")"
    
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0) ''Button Portion
    ''SetControlPropertiesFromTemplate ctl, frm
    CopyProperties frm, ctl.Name, "ButtonControl"
    ctl.Name = concat("cmd", fldName)
    ctl.Caption = "Browse..."
    ctl.OnClick = "=SelectDEDirectory([Form]," & EscapeString(fldName) & ")"
    ctl.Height = 300
    
    
    ''Render the Filter buttons
    ''Filter and Clear
    proportionArr.arr = "10,2"
    controlArr.arr = fldName & "," & concat("cmd", fldName)
    proportionTotal = GetProportionTotal(proportionArr)
    
    For i = 0 To proportionArr.count - 1
    
        proportion = CDbl(proportionArr.arr(i)) / proportionTotal
        frm(controlArr.arr(i)).Left = x
        frm(controlArr.arr(i)).Top = y
        frm(controlArr.arr(i)).Width = (totalWidth - ((proportionArr.count - 1) * colSpaceWidth * 2)) * proportion
        frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
       
        x = x + (colSpaceWidth * 2) + frm(controlArr.arr(i)).Width
       
    Next i
    
    frm.controls(fldName).HorizontalAnchor = acHorizontalAnchorLeft
    frm.controls(concat("cmd", fldName)).HorizontalAnchor = acHorizontalAnchorLeft
    
End Function

Public Function CreateSimpleDEForm(frm2 As Form)
    
    CreateDEForm frm2, False, True
    
End Function

Private Sub InsertToModelRelatedObjects(ModelID, ObjectTypeID, objectName)
    
    If Not isPresent("tblModelRelatedObjects", "ObjectName = " & EscapeString(objectName)) Then
        RunSQL "INSERT INTO tblModelRelatedObjects (ModelID, ObjectTypeID, ObjectName) VALUES (" & _
               ModelID & "," & _
               ObjectTypeID & "," & _
               EscapeString(objectName) & ")"
    End If

End Sub

''This function will set the properties of a control in a form by inheriting the properties from frmControlTemplate
Public Sub SetControlPropertiesFromTemplate(ctl As control, obj As Object, Optional insideTab As Boolean = False)
    
    SetControlProperties ctl
    CopyControlTemplateProperties obj, ctl.Name, insideTab
     
End Sub

Public Sub SetControlProperties(ctl As control)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryControlTypes WHERE ControlTypeValue = " & ctl.ControlType)
    
    Do Until rs.EOF
        ctl.Properties(rs.fields("ControlPropValue")) = rs.fields("ControlProp")
        rs.MoveNext
    Loop
    
End Sub

Private Function RenderButton(x, y, Caption, frm As Form, cmdName, Optional parent = "")
        
    Dim ctl As control, fldWidth
    
    fldWidth = 3000 * (0.5) + (200 * (0.5 - 1))
    
    Set ctl = CreateControl(frm.Name, acCommandButton, , parent, , x, y, fldWidth)
    With ctl
    
        .Name = "cmd" & cmdName
        .Properties("Caption") = Caption
        
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM qryControlTypes WHERE ControlTypeValue = 104")
        
        Do Until rs.EOF
            If DoesPropertyExists(.Properties, rs.fields("ControlPropValue")) Then
                .Properties(rs.fields("ControlPropValue")) = rs.fields("ControlProp")
            End If
            rs.MoveNext
        Loop
        
        '.Properties("UseTheme") = False
        .Properties("CursorOnHover") = 1
        
    End With
    
    CopyProperties frm, ctl.Name, "ButtonControl", Not isFalse(parent)
    
End Function

Public Function GetMaxX(frm As Object, Optional atYPosition = Null, Optional pgName = "") As Double
    
    ''MAX X doesn't take into the allowance or margin
    Dim ctl As control, x As Double
    Dim ctls As controls
    Dim maxX As Double
'    If Not isFalse(pgName) Then
'        Set ctls = frm(pgName).controls
'    Else
'        Set ctls = frm.controls
'    End If

    Set ctls = frm.controls
    
    For Each ctl In ctls
        If ctl.Left + ctl.Width > x Then

            If Not isFalse(pgName) Then
                If ctl.ControlType = acCommandButton Then
                    If IsNull(atYPosition) Then
                    x = ctl.Left + ctl.Width
                    Else
                        If ctl.Top = atYPosition Then
                            x = ctl.Left + ctl.Width
                        End If
                    End If
                End If
            Else
                If IsNull(atYPosition) Then
                    x = ctl.Left + ctl.Width
                Else
                    If ctl.Top = atYPosition Then
                        x = ctl.Left + ctl.Width
                    End If
                End If
            End If

        End If
    Next ctl

'    For Each ctl In ctls
'        If ctl.Left + ctl.Width > maxX Then
'            If ctl.ControlType = acCommandButton And (IsNull(atYPosition) Or ctl.Top = atYPosition) Then
'                maxX = ctl.Left + ctl.Width
'            End If
'        End If
'    Next ctl
    
    GetMaxX = x
    
End Function

Public Function GetMaxY(frm As Object, Optional objectSection = Null, Optional x = Null, Optional totalWidth = Null, Optional pgName = "", Optional minY = 0) As Double
    
    Dim ctl As control, y As Double
    Dim frmControls As Object
    If Not IsNull(objectSection) Then
        Set frmControls = frm.Section(objectSection).controls
    Else
        Set frmControls = frm.controls
    End If
    
    For Each ctl In frmControls
    
        If Not IsNull(x) Then
            ''If ctl.Left + totalWidth >= x And ctl.Left + totalWidth <= x + totalWidth Then
            If ctl.Left + ctl.Width >= x And ctl.Left <= x + totalWidth Then
                If ctl.Top + ctl.Height > y Then
                    y = ctl.Top + ctl.Height
                End If
            End If
        Else
            If ctl.Top + ctl.Height > y Then
                y = ctl.Top + ctl.Height
            End If
        End If
        
    Next ctl
    
    If y < minY Then y = minY
    GetMaxY = y
    
End Function

Private Sub RenderDatasheetTotals(frm As Object, ModelID, maxWidth)
    
    Dim y, formMargin, x
    formMargin = 100
    y = GetMaxY(frm) + 700
    x = formMargin
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblDatasheetTotals WHERE ModelID = " & ModelID & " AND " & _
                             "IsHidden <> -1 ORDER BY ControlOrder ASC")
                             
    Dim ctlName, fldWidth, ctl As control
    fldWidth = 3000 * (0.5) + (200 * (0.5 - 1))
    
    If Not rs.EOF Then
        
        Do Until rs.EOF
            
            ctlName = rs.fields("ControlName")
            
            If x + fldWidth + 100 > maxWidth Then
                x = formMargin
                y = GetMaxY(frm) + 800
            End If
            
            Set ctl = CreateControl(frm.Name, acTextBox, , "", "", x, y, fldWidth)
            ctl.Name = ctlName
            ctl.ControlSource = concat("=IfError(subform!", ctlName, ")")
            ctl.Format = "Standard"
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
            
            ''Set control property based on ControlTypeValue
            SetControlPropertiesFromTemplate ctl, frm
            
            ''Generate the label just above the control
            Set ctl = CreateControl(frm.Name, acLabel, , ctl.Name, rs.fields("ControlCaption"), x, y - 500)
            
            SetControlPropertiesFromTemplate ctl, frm
            ctl.Height = 500
            ctl.Width = fldWidth
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
           
            x = x + fldWidth + 100
            
            rs.MoveNext
        Loop
        
    End If
    
End Sub

Private Sub RenderAdditionalButtonOnMainForm(ModelID, frm As Form, y)
    
    Dim rs As Recordset, sqlStr As String, ctl As control
    Dim maxX As Long, ModelButton, modelButtonName, FunctionName, TableWideFunction, cmdButtonName
    ''Check first if there's atleast one additional button
    sqlStr = "SELECT * FROM tblModelButtons WHERE ModelID = " & ModelID & _
             " AND HideOnMain <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
    Set rs = ReturnRecordset(sqlStr)

    If Not rs.EOF Then
        ''Check on what needs to be rendered, individual buttons or combo boxes.
        If isPresent("qryModelProperties", "ModelID = " & ModelID & " And Property = " & EscapeString("collapsedButtonOnMain")) Then
            ''Collapsed button so this is a combo box
            ''Create a combo box
            ''Left position should account for the label "Action:"
            ''55 is the space between controls,
            Dim lblWidth: lblWidth = 1000
            maxX = GetMaxX(frm) + 55 + lblWidth + 55
            Set ctl = CreateControl(frm.Name, acComboBox, , , , maxX, y, 3000, 400)
            ''Set the Default Control Properties Here
            SetControlPropertiesFromTemplate ctl, frm
            ''Additional Property make the RowSource to be the SQLStr, ColumnCount to 2, ColumnWidths to 0;1
            ''Set the Height to be the same height as the buttons
            sqlStr = "SELECT ModelButtonID,ModelButton FROM tblModelButtons WHERE ModelID = " & ModelID & _
                     " AND HideOnMain <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
            ctl.Name = "cboFormActions"
            ctl.RowSource = sqlStr
            ctl.ColumnCount = 2
            ctl.ColumnWidths = "0;1"
            ctl.Height = 400
            ctl.TopMargin = 75
            ctl.LeftMargin = 75
            ctl.FontBold = True
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            
            ''Render the label here
            Set ctl = CreateControl(frm.Name, acLabel, , "cboFormActions", , maxX - 55 - lblWidth, y, lblWidth, 400)
            ''Set the Default Control Properties Here
            SetControlPropertiesFromTemplate ctl, frm
            ctl.Name = "lblFormActions"
            ctl.Caption = "Actions: "
            ctl.TextAlign = 3
            ctl.Height = 400
            ctl.TopMargin = 75
            ctl.LeftMargin = 75
            ctl.FontBold = True
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            
            maxX = GetMaxX(frm) + 55
            RenderButton maxX, y, "Run", frm, "RunFormActions"
            frm("cmdRunFormActions").Width = frm("cmdRunFormActions").Width / 2
            frm("cmdRunFormActions").HorizontalAnchor = acHorizontalAnchorRight
            
            frm("cmdRunFormActions").OnClick = "=RunFormActions([Form],[cboFormActions])"
            
        Else
            Do Until rs.EOF
                ModelButton = rs.fields("ModelButton")
                modelButtonName = RemoveSpaces(ModelButton)
                cmdButtonName = concat("cmd", modelButtonName)
                FunctionName = rs.fields("FunctionName")
                TableWideFunction = rs.fields("TableWideFunction")
                
                maxX = GetMaxX(frm) + 55
                RenderButton maxX, y, ModelButton, frm, modelButtonName
                If Not IsNull(FunctionName) Then
                    If TableWideFunction Then
                        frm(cmdButtonName).OnClick = concat("=", FunctionName, "()")
                    Else
                        frm(cmdButtonName).OnClick = concat("=RunFunctionFromSubform([Form],""subform"",", EscapeString(FunctionName), ")")
                    End If
                End If
                
                maxX = maxX + (3200 * 0.45)
                rs.MoveNext
            Loop
        End If
        
        
    End If
    
End Sub

Private Sub RenderAdditionalButtonOnDEForm(ModelID, frm As Object, y, x, CurrentCol, buttonMultiplier, FormColumns)
    
    Dim rs As Recordset, sqlStr As String, ctl As control
    Dim maxX As Long, ModelButton, modelButtonName, FunctionName, TableWideFunction, cmdButtonName
    ''Check first if there's atleast one additional button
    sqlStr = "SELECT * FROM tblModelButtons WHERE ModelID = " & ModelID & _
             " AND HideOnForm <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
    Set rs = ReturnRecordset(sqlStr)

    If Not rs.EOF Then
        ''Check on what needs to be rendered, individual buttons or combo boxes.
        If isPresent("qryModelProperties", "ModelID = " & ModelID & " And Property = " & EscapeString("collapsedButtonOnForm")) Then
            ''Collapsed button so this is a combo box
            ''Create a combo box
            ''Left position should account for the label "Action:"
            ''55 is the space between controls,
            Dim lblWidth: lblWidth = 1000
            maxX = GetMaxX(frm, y) + 55 + lblWidth + 55
            Set ctl = CreateFlexControl(frm.Name, acComboBox, , , , maxX, y, 3000, 400)
            ''Set the Default Control Properties Here
            SetControlPropertiesFromTemplate ctl, frm
            ''Additional Property make the RowSource to be the SQLStr, ColumnCount to 2, ColumnWidths to 0;1
            ''Set the Height to be the same height as the buttons
            sqlStr = "SELECT ModelButtonID,ModelButton FROM tblModelButtons WHERE ModelID = " & ModelID & _
                     " AND HideOnForm <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
            ctl.Name = "cboFormActions"
            ctl.RowSource = sqlStr
            ctl.ColumnCount = 2
            ctl.ColumnWidths = "0;1"
            ctl.Height = 400
            ctl.TopMargin = 75
            ctl.LeftMargin = 75
            ctl.FontBold = True
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            ctl.VerticalAnchor = acVerticalAnchorBottom
            
            ''Render the label here
            Set ctl = CreateFlexControl(frm.Name, acLabel, , "cboFormActions", , maxX - 55 - lblWidth, y, lblWidth, 400)
            ''Set the Default Control Properties Here
            SetControlPropertiesFromTemplate ctl, frm
            ctl.Name = "lblFormActions"
            ctl.Caption = "Actions: "
            ctl.TextAlign = 3
            ctl.Height = 400
            ctl.TopMargin = 75
            ctl.LeftMargin = 75
            ctl.FontBold = True
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            ctl.VerticalAnchor = acVerticalAnchorBottom
            
            maxX = GetMaxX(frm, y) + 55
            RenderButton maxX, y, "Run", frm, "RunFormActions"
            frm("cmdRunFormActions").Width = frm("cmdRunFormActions").Width / 2
            frm("cmdRunFormActions").HorizontalAnchor = acHorizontalAnchorRight
            frm("cmdRunFormActions").VerticalAnchor = acVerticalAnchorBottom
            frm("cmdRunFormActions").OnClick = "=RunFormActionFromDE([Form],[cboFormActions])"
            
        ElseIf isPresent("qryModelProperties", "ModelID = " & ModelID & " And Property = " & EscapeString("listboxButtonOnForm")) Then
            
            ''RenderListBoxButton(frm As Object, ControlCreationHelperID, Optional xPosition, Optional parentFormName)
            maxX = GetMaxX(frm) + 300
            Dim Model: Model = ELookup("tblModels", "ModelID = " & ModelID, "Model")
            Dim ControlCreationHelperID: ControlCreationHelperID = ELookup("qryControlCreationHelper", "Model = " & Esc(Model) & _
                " AND CustomControlType = ""Listbox Button""", "ControlCreationHelperID")
            Dim parentFormName: parentFormName = frm.Name
            
            RenderListBoxButton frm, ControlCreationHelperID, maxX, parentFormName
            
        Else
            
            maxX = x
            
            Do Until rs.EOF
            
                ModelButton = rs.fields("ModelButton")
                modelButtonName = RemoveSpaces(ModelButton)
                cmdButtonName = concat("cmd", modelButtonName)
                FunctionName = rs.fields("FunctionName")
                TableWideFunction = rs.fields("TableWideFunction")
                
                RenderButton maxX, y, ModelButton, frm, modelButtonName
                frm(cmdButtonName).HorizontalAnchor = 0
                frm(cmdButtonName).VerticalAnchor = 1
                If Not IsNull(FunctionName) Then
                    If TableWideFunction Then
                        frm(cmdButtonName).OnClick = concat("=", FunctionName, "()")
                    Else
                        frm(cmdButtonName).OnClick = concat("=", FunctionName, "([Form])")
                    End If
                End If
                
                CurrentCol = CurrentCol + 0.5
            
                If CurrentCol = FormColumns Then
                    
                    CurrentCol = 0
                    maxX = 400
                    y = y + 600
                    
                Else
                
                    maxX = maxX + (3200 * buttonMultiplier)
                    
                End If
                
                rs.MoveNext
            Loop
            
        End If
        
        
    End If
    
End Sub

Public Function RunFormActionFromDE(frm, ModelButtonID)
    
    Dim rs As Recordset, sqlStr As String
    Dim ModelButton, modelButtonName, cmdButtonName, FunctionName, TableWideFunction
    sqlStr = "SELECT * FROM tblModelButtons WHERE ModelButtonID = " & ModelButtonID
    
    Set rs = ReturnRecordset(sqlStr)
    
    ModelButton = rs.fields("ModelButton")
    modelButtonName = RemoveSpaces(ModelButton)
    cmdButtonName = concat("cmd", modelButtonName)
    FunctionName = rs.fields("FunctionName")
    TableWideFunction = rs.fields("TableWideFunction")
    
    If TableWideFunction Then
        Run FunctionName
    Else
        Run FunctionName, frm
    End If

End Function

Public Function RunFormActions(frm, ModelButtonID, Optional SubformName = "subform")
    
    Dim rs As Recordset, sqlStr As String
    Dim ModelButton, modelButtonName, cmdButtonName, FunctionName, TableWideFunction
    sqlStr = "SELECT * FROM tblModelButtons WHERE ModelButtonID = " & ModelButtonID
    
    Set rs = ReturnRecordset(sqlStr)
    
    ModelButton = rs.fields("ModelButton")
    modelButtonName = RemoveSpaces(ModelButton)
    cmdButtonName = concat("cmd", modelButtonName)
    FunctionName = rs.fields("FunctionName")
    TableWideFunction = rs.fields("TableWideFunction")
    
    If TableWideFunction Then
        Run FunctionName
    Else
        Run "RunFunctionFromSubform", frm, SubformName, FunctionName
    End If

    
End Function

Private Function GetVariableSource(ModelID, frm As Form) As Object
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblSupplementalModels WHERE ModelID = " & ModelID)
    
    If rs.EOF Then GoTo UseForm:
    
    Dim UseAsModel: UseAsModel = rs.fields("UseAsModel")
    
    If isFalse(UseAsModel) Then GoTo UseForm:
    
    ''Use the Model indicated for this form. This is important so that the form name primary key name and other
    ''aspects can be assumed from the Model Itself -> But not all properties of the model shall be used.
    ''only the select parts.
    Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & UseAsModel)
    Set GetVariableSource = rs.fields
    
    Exit Function
    
UseForm:
    
    Set GetVariableSource = frm
    
End Function

Public Function CreateMainForm(frm2 As Form)
    
    Dim frm As Form, rs As Recordset, rsName, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth
    Dim tblDef As DAO.TableDef, db As DAO.Database, fld As DAO.field
    Dim ctl As control, baseName
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus
    Dim QueryName, OnFormCreate, SubformName, Timestamp, CreatedBy, PrimaryKey
    
    ''Declare the variables -> Look at the tblSupplementalModels and see if this model has a UseAsModel Property
    ''Get the object to be used -> Either frm2 or the rs.Fields
    ModelID = frm2("ModelID")
    Dim src As Object: Set src = GetVariableSource(ModelID, frm2)
    
    Model = src("Model")
    VerboseName = src("VerboseName")
    VerbosePlural = src("VerbosePlural")
    MainField = src("MainField")
    TableWideValidation = src("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    SubformName = frm2("SubformName")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    PrimaryKey = frm2("PrimaryKey")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    ''CreateDSForm frm2, True
    ''Create the form
    Set frm = CreateForm
    rsName = QueryName
    If IsNull(rsName) Then rsName = GetTableName(Model, VerbosePlural)
    Dim VerboseCaption: VerboseCaption = ELookup("tblSupplementalModels", "ModelID = " & ModelID, "VerboseCaption")
    Dim frmCaption: frmCaption = GetFieldCaption(VerboseName, Model, VerboseCaption)
    frm.Caption = concat(frmCaption, " List")
    
    If IsNull(PrimaryKey) Then PrimaryKey = concat(Model, "ID")
    frm.OnLoad = "=DefaultMainFormLoad([Form])"
    
    SetFormProperties 6, frm
    
    x = 100
    y = 100
    
    If Not IsNull(VerbosePlural) Then
        baseName = concat(replace(VerbosePlural, " ", ""))
    Else
        baseName = concat(Model, "s")
    End If
    
    ''frmName would be the variable to be used as the name of the data entry form
    ''Will be used on various buttons from the main form
    Dim frmName: frmName = "frm" & baseName
    
    If Not IsNull(SubformName) Then
        baseName = SubformName
    End If
    
    ''Add Button
    If Not isPresent("qryModelProperties", "Property = ""mainAddHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Add New", frm, "Add"
        frm.cmdAdd.OnClick = "=OpenFormFromMain(" & Esc(frmName) & ")"
        x = x + (3200 * 0.45)
    End If
    
    ''Edit Button
    If Not isPresent("qryModelProperties", "Property = ""mainEditHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "View/Edit", frm, "View"
        frm.cmdView.OnClick = "=OpenFormFromMain(" & Esc(frmName) & ", ""subform"", """ & PrimaryKey & """,[Form])"
        x = x + (3200 * 0.45)
    End If
    
    ''Save Button
    If Not isPresent("qryModelProperties", "Property = ""mainDeleteHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Delete", frm, "Delete"
        frm.cmdDelete.OnClick = "=DeleteRecord([Form], """ & PrimaryKey & """, """ & rsName & """, ""subform"")"
        x = x + (3200 * 0.45)
    End If
    
    ''Additional Buttons Here (tblModelButtons)
    RenderAdditionalButtonOnMainForm ModelID, frm, y
    
    Dim maxX As Long
    
    y = y + 500: x = 100
    
    ''Get the Max x to set the width of the subform
    ''must not be less than 10000
    ''Get the x + length of the leftmost button
    maxX = GetMaxX(frm) - 100
    If maxX < 10000 Then maxX = 10000
    
    Set ctl = CreateControl(frm.Name, acSubform, , , "subform", x, y, maxX, 7000)
    ctl.Name = "subform"
    ctl.Properties("RightPadding") = 100
    ctl.SourceObject = "dsht" & baseName
    ctl.HorizontalAnchor = acHorizontalAnchorBoth
    ctl.VerticalAnchor = acVerticalAnchorBoth
    
    RenderDatasheetTotals frm, ModelID, maxX
    
    ''Align the collapsed controls if there's any
    If DoesPropertyExists(frm.controls, "cboFormActions") Then
        maxX = GetMaxX(frm)
        frm("cmdRunFormActions").Left = maxX - frm("cmdRunFormActions").Width
        frm("cboFormActions").Left = frm("cmdRunFormActions").Left - 55 - frm("cboFormActions").Width
        frm("lblFormActions").Left = frm("cboFormActions").Left - 55 - frm("lblFormActions").Width
    End If
    
    ''Render the filterForm here
    RenderFilterForm frm, ModelID
    
    ''Resize the subform control if the cmdFilter is greater than the current height + top of the subform
    ResizeSubform frm
    
    frm.Section("Detail").Height = GetMaxY(frm) + 200
    frm.Width = GetMaxX(frm) + 100
    
    Dim customFrmName As String, baseFormName As String, i As Integer
    frmName = frm.Name
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("main", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("main", Model, "s")
    End If
    
    If Not IsNull(SubformName) Then
        baseFormName = concat("main", SubformName)
    End If
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, frm, 6
    End If
    
    ''Override Properties Here
    OverrideProperties ModelID, 6, frm
    
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not FrmExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    ''Insert the newly created form to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acForm, customFrmName
    
    DoCmd.Rename customFrmName, acForm, frmName
    
    ''Load Save Form Layout
    LoadSavedFormLayout customFrmName
    
    DoCmd.OpenForm customFrmName
    
    InsertFormInFormForRights customFrmName, Model

End Function

Private Function ResizeSubform(frm As Form)
    ''Get the subform's top, height
    ''See if cmdFilter control is present if not just leave the subform
    If DoesPropertyExists(frm.controls, "cmdFilter") Then
        Dim subformTop, subformHeight, subformBottom, subform As control
        Dim cmdTop, cmdHeight, cmdBottom, cmdFilter As control
        Set subform = frm("subform")
        Set cmdFilter = frm("cmdFilter")
        
        subformTop = subform.Top: subformHeight = subform.Height
        subformBottom = subformTop + subformHeight
        
        cmdTop = cmdFilter.Top: cmdHeight = cmdFilter.Height
        cmdBottom = cmdTop + cmdHeight
        
        Dim heightAdjustment: heightAdjustment = cmdBottom - subformBottom
        If heightAdjustment < 0 Then heightAdjustment = 0
        
        frm("subform").Height = subformHeight + heightAdjustment
        
    End If
    
End Function

Private Sub NewRow(frm As Object, x, ByRef y, originalY, totalWidth)

    y = GetMaxY(frm, , x, totalWidth) + 100
    If y = 100 Then y = originalY
    
End Sub

Private Sub SetComboBoxSQLForFilter(ctl As control, ModelFieldID)
    
    Dim parentModelRs As Recordset, modelFieldRS As Recordset
    Set modelFieldRS = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & ModelFieldID)
    
    Dim ParentModelID: ParentModelID = modelFieldRS.fields("ParentModelID")
    
    Dim sqlStr
    If isFalse(ParentModelID) Then
        Dim ModelField: ModelField = modelFieldRS.fields("ModelField")
        Dim ModelID: ModelID = modelFieldRS.fields("ModelID")
        Dim TableName: TableName = GetTableNameFromModelID(ModelID)
        sqlStr = "SELECT DISTINCT " & ModelField & " FROM " & TableName & " ORDER BY " & ModelField
    Else
        Set parentModelRs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & modelFieldRS.fields("ParentModelID"))
        
        Dim MainField: MainField = parentModelRs.fields("MainField")
        If MainField Like "=*" Then MainField = replace(MainField, "=", "")
        
        Dim primaryTable: primaryTable = GetTableName(parentModelRs.fields("Model"), parentModelRs.fields("VerbosePlural"))
        
        Dim primaryField: primaryField = concat(parentModelRs.fields("Model"), "ID")
        
        
        sqlStr = concat("SELECT ", primaryField, ",", MainField, " As MainField FROM ", primaryTable, " ORDER BY ", MainField)
        
        ctl.ColumnCount = 2
        ctl.ColumnWidths = "0;1"
        
        
    End If
    
    ctl.RowSource = sqlStr
    
End Sub

Private Function GenerateWildSearchFromRelatedFieldSQL(fltrArr As clsArray, ModelID, ctlValue) As String
    
    Dim filterValue
    ''Open the wildcards filterFields => Related Fields
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblRelatedFilterFields WHERE ModelID = " & ModelID & " And IncludeInWildcardSearch")
    
    Dim TableName, FieldToUse, fieldName, joinTableName
    If Not rs.EOF Then
        ''Check if the control is not null
        If Not IsNull(ctlValue) Then
            TableName = rs.fields("TableName")
            FieldToUse = rs.fields("FieldToUse")
            joinTableName = "temp" & TableName
            fieldName = GetSQLFieldName(FieldToUse, joinTableName)
            fltrArr.Add fieldName & " Like " & EscapeString("*" & ctlValue & "*")
        End If
    End If
    
End Function

Private Function GenerateWildSearchSQL(ctlValue, Optional rsName, Optional ModelID) As String

    Dim modelFieldRS As Recordset, fltrArr As New clsArray, fieldName
    ''Open the wildcard recordset for this model -> main table only
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE ModelID = " & ModelID & " AND IsWildSearch")
    
    Do Until rs.EOF
        Set modelFieldRS = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & rs.fields("ModelFieldID"))
        fieldName = GetSQLFieldName(modelFieldRS.fields("ModelField"), rsName)
        fltrArr.Add fieldName & " Like " & EscapeString("*" & ctlValue & "*")
        rs.MoveNext
    Loop
    
    GenerateWildSearchFromRelatedFieldSQL fltrArr, ModelID, ctlValue
    
    GenerateWildSearchSQL = "(" & fltrArr.JoinArr(" OR ") & ")"
    
End Function

Private Function GenerateNumericSearch(fieldName, fromValue, toValue) As String
    
    If IsNull(fromValue) And IsNull(toValue) Then Exit Function
    
    If Not IsNull(fromValue) And Not IsNull(toValue) Then
        GenerateNumericSearch = fieldName & " Between " & fromValue & " And " & toValue
    ElseIf Not IsNull(toValue) Then
        GenerateNumericSearch = fieldName & " <= " & toValue
    ElseIf Not IsNull(fromValue) Then
        GenerateNumericSearch = fieldName & " >= " & fromValue
    End If
    
End Function

Private Function GenerateDateSearch(fieldName, fromValue, toValue) As String
    
    If IsNull(fromValue) And IsNull(toValue) Then Exit Function
    
    If Not IsNull(fromValue) And Not IsNull(toValue) Then
        GenerateDateSearch = fieldName & " Between #" & fromValue & "# And #" & toValue & "#"
    ElseIf Not IsNull(toValue) Then
        GenerateDateSearch = fieldName & " <= #" & toValue & "#"
    ElseIf Not IsNull(fromValue) Then
        GenerateDateSearch = fieldName & " >= #" & fromValue & "#"
    End If
    
End Function


Private Function GenerateMonthYearSearch(fieldName, monthValue, yearValue) As String
    
    Dim fltrArr As New clsArray
    If IsNull(monthValue) And IsNull(yearValue) Then Exit Function
    
    If Not IsNull(monthValue) Then
        fltrArr.Add "Month(" & fieldName & ") = " & monthValue
    End If
    
    If Not IsNull(yearValue) Then
        fltrArr.Add "Year(" & fieldName & ") = " & yearValue
    End If
    
    If fltrArr.count > 0 Then
        GenerateMonthYearSearch = fltrArr.JoinArr(" AND ")
    End If
    
End Function

Private Sub ClearCBValue(ctl As control)
    On Error Resume Next
    ctl = ctl.DefaultValue
End Sub

Public Function ClearFilterSubform(frm As Object, ModelID, Optional useQuery As Boolean = False)
    
    Dim rsName
    rsName = GetTableNameFromModelID(ModelID)
    
    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.Name Like "fltr*" Then
            If ctl.ControlType = acSubform Then
                ''Run the onLoad event of the form if it is a subform
                Dim onLoadEvent: onLoadEvent = ctl.Form.OnLoad
                ''Remove the equals, extract the parameters other than the [Form]
                ''First item is the sqlStr, 2nd item is the TableName
                Dim FormParams As New clsArray
                Set FormParams = ExtractFilterContFormOnLoadParams(onLoadEvent)
                
                Run "FilterContFormOnLoad", ctl.Form, FormParams.items(0), FormParams.items(1)
                
            ElseIf ctl.ControlType <> acCheckBox And ctl.ControlType <> acOptionButton Then
                ctl = Null
            ElseIf ctl.ControlType = acOptionGroup Then
                ctl = 2
            ElseIf ctl.ControlType = acCheckBox Then
                ClearCBValue ctl
            End If
        End If
        
    Next ctl
    
    If useQuery Then
        SetSubformSQL frm, ModelID
        Exit Function
    End If
    
    If Not frm.subform.SourceObject Like "Report.*" Then
'        orderBy = frm.subform.Form.orderBy
'        If Not IsNull(orderBy) And orderBy <> "" Then sqlStr = sqlStr & " ORDER BY " & orderBy
'        frm.subform.Form.RecordSource = sqlStr
'        frm.subform.Form.orderBy = orderBy
'        frm.subform.Requery
        frm.subform.Form.FilterOn = False
    Else
'        orderBy = frm.subform.Form.orderBy
'        If Not IsNull(orderBy) And orderBy <> "" Then sqlStr = sqlStr & " ORDER BY " & orderBy
'        frm.subform.Report.RecordSource = sqlStr
'        frm.subform.Form.orderBy = orderBy
'        frm.subform.Requery
        frm.subform.Report.FilterOn = False
    End If
 
End Function

Private Function GetSQLFieldName(fieldName, Optional rsName)

    GetSQLFieldName = fieldName
    If Not isFalse(rsName) Then
        GetSQLFieldName = rsName & "." & fieldName
    End If
    
End Function

Public Function GetFilterArray(frm As Object, ModelID, Optional useQuery As Boolean = False) As clsArray
    
    ''Get the default TableName of the query
    Dim rsName
    If useQuery Then rsName = GetTableNameFromModelID(ModelID)

    ''Open the wildcards filterFields => Main Form
    ''Check first if there's a wildcard filter
    Dim containsWildcard: containsWildcard = CheckWildcardPresence(ModelID)
    
    ''Fetch the FilterFields
    Dim rs As Recordset, modelRS As Recordset, modelFieldRS As Recordset, ctl As control
    Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE ModelID = " & ModelID & " AND FilterOrder > 0")
    
    ''Check first if the recordset is empty
    If rs.EOF And Not containsWildcard Then Exit Function
   
    Dim ctlName: ctlName = "fltrWildSearch"
    Dim fltrArr As New clsArray, fieldName, filterValue
    If containsWildcard Then
        filterValue = frm(ctlName)
        ''Check if the control is not null
        If Not IsNull(filterValue) Then
            fltrArr.Add GenerateWildSearchSQL(filterValue, rsName, ModelID)
            RunSQL "INSERT INTO tblSearchHistorys (SearchTerm,FormName) VALUES (" & Esc(filterValue) & "," & Esc(frm.Name) & ")"
        End If
    End If
    
    
    ''User filter controls which is not a wildsearch (filter using fields from the table itself)
    Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE ModelID = " & ModelID & " ANd IsWildSearch = 0 AND FilterOrder > 0 Order By FilterOrder ASC")
    Do Until rs.EOF
    
        Set modelFieldRS = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & rs.fields("ModelFieldID"))

        ''Boolean Filter
        If modelFieldRS.fields("FieldTypeID") = dbBoolean Then
            
            fieldName = GetSQLFieldName(modelFieldRS.fields("ModelField"), rsName)
            ctlName = "ogFltr" & fieldName
            
            If Not IsNull(frm(ctlName)) Then
                Select Case frm(ctlName)
                    Case 2:
                        
                    Case Else:
                        fltrArr.Add fieldName & " = " & frm(ctlName)
                        
                End Select
            End If
            GoTo NextFilter:
        End If
        
        If Not IsNull(rs.fields("FilterOperator")) Then
            
            fieldName = GetSQLFieldName(modelFieldRS.fields("ModelField"), rsName)
            ctlName = "fltr" & fieldName
               
            If Not IsNull(frm(ctlName)) Then fltrArr.Add fieldName & " Like " & EscapeString("*" & frm(ctlName) & "*")
            
            GoTo NextFilter:
            
        End If
        
        If rs.fields("IsList") Then
            
            fieldName = GetSQLFieldName(modelFieldRS.fields("ModelField"), rsName)
            ctlName = "fltr" & fieldName
            Dim ctlValue: ctlValue = frm(ctlName)
            
            If Not IsNull(ctlValue) Then
            
                Dim Numbers As New clsArray: Numbers.arr = "3,4,5,6,7"
                If Not Numbers.InArray(CStr(modelFieldRS.fields("FieldTypeID"))) Then
                    ctlValue = EscapeString(ctlValue)
                End If
                
                fltrArr.Add fieldName & " = " & ctlValue
            End If
            
            GoTo NextFilter:
        End If

        ''Double Filter
        Dim resultingSQL
        If modelFieldRS.fields("FieldTypeID") = dbDouble Then
            
            fieldName = GetSQLFieldName(modelFieldRS.fields("ModelField"), rsName)
            ctlName = "fltr" & fieldName
            
            resultingSQL = GenerateNumericSearch(fieldName, frm(ctlName & "From"), frm(ctlName & "To"))

            If resultingSQL <> "" Then
                fltrArr.Add resultingSQL
            End If
            GoTo NextFilter:
        End If

        If rs.fields("IsMonthYear") Then
        
            fieldName = GetSQLFieldName(modelFieldRS.fields("ModelField"), rsName)
            ctlName = "fltr" & fieldName
            
            resultingSQL = GenerateMonthYearSearch(fieldName, frm(ctlName & "Month"), frm(ctlName & "Year"))
            
            If resultingSQL <> "" Then
                fltrArr.Add resultingSQL
            End If
            GoTo NextFilter:
        End If

        If rs.fields("IsBetween") Then
        
            fieldName = GetSQLFieldName(modelFieldRS.fields("ModelField"), rsName)
            ctlName = "fltr" & modelFieldRS.fields("ModelField")
            
            resultingSQL = GenerateDateSearch(fieldName, frm(ctlName & "From"), frm(ctlName & "To"))

            If resultingSQL <> "" Then
                fltrArr.Add resultingSQL
            End If
            GoTo NextFilter:

        End If
        
        ''If a checkbox or optionbutton
        ''if the field ahs parentModelId then dont do escapestring
        ''put the values in a parentheses joined by or
        If rs.fields("IsCheckboxList") Or rs.fields("IsOptionGroup") Then
        
            ''Cancel out fieldName if the filter has Caption
            fieldName = rs.fields("FilterCaption")
            If isFalse(fieldName) Then
                fieldName = modelFieldRS.fields("ModelField")
            End If
            
            Dim hasOption: hasOption = isPresent("tblFilterFieldOptions", "FilterFieldID = " & rs.fields("FilterFieldID"))
            
            resultingSQL = GetCbOrOGFilterStr(frm, rs.fields("IsDynamicList"), rs.fields("IsOptionGroup"), modelFieldRS, fieldName, rsName, hasOption)
            If resultingSQL <> "" Then
                fltrArr.Add resultingSQL
            End If
        End If
        
    
NextFilter:
        rs.MoveNext
    Loop
    
    AddUnsatisfiedRelatedFilterFields ModelID, fltrArr, frm
    
    ''Look at the tblRelatedFilterFields -> SatisfiesFilter false, FieldToUse false
    ''Add the result to the fltrArr -> IsNull(TableName!RightJoinKey)
    ''Dim subQueryAlias: subQueryAlias = "temp" & TableName
    
    Set GetFilterArray = fltrArr
    
End Function

Private Sub AddUnsatisfiedRelatedFilterFields(ModelID, fltrArr As clsArray, frm As Form)

    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblRelatedFilterFields WHERE ModelID = " & ModelID & " AND Not SatisfiesFilter AND FieldToUse IS NULL")
    
    Do Until rs.EOF
        Dim SatisfiesFilter: SatisfiesFilter = rs.fields("SatisfiesFilter")
        Dim TableName: TableName = rs.fields("TableName")
        Dim subqueryName: subqueryName = "temp" & TableName & "Unsatisfaction"
        Dim cbName: cbName = "fltr" & TableName & SatisfiesFilter
        Dim RightJoinKey: RightJoinKey = rs.fields("RightJoinKey")
        If frm(cbName) Then fltrArr.Add subqueryName & "!" & RightJoinKey & " IS NULL"
        rs.MoveNext
    Loop
    
End Sub

Public Function FilterSubform(frm As Object, ModelID, Optional ReturnMode = False)
    
    Dim filterStr, OrderBy, sqlStr
    Dim fltrArr As New clsArray: Set fltrArr = GetFilterArray(frm, ModelID)
    
    If Not ReturnMode Then
        If fltrArr.count > 0 Then
            filterStr = fltrArr.JoinArr(" AND ")
            If Not frm.subform.SourceObject Like "Report.*" Then
                frm.subform.Form.filter = filterStr
                frm.subform.Form.FilterOn = True
            Else
                frm.subform.Report.filter = filterStr
                frm.subform.Report.FilterOn = True
            End If
        End If
    End If
    
End Function

Private Function GetCbOrOGFilterStr(frm, IsDynamicList, IsOptionButton As Boolean, modelFieldRS As Recordset, fieldName, Optional rsName, Optional hasOption = False) As String
        
    Dim ParentModelID: ParentModelID = modelFieldRS.fields("ParentModelID")
    Dim FieldTypeID: FieldTypeID = CStr(modelFieldRS.fields("FieldTypeID"))
    Dim Numbers As New clsArray: Numbers.arr = "3,4,5,6,7"
    ''3,4,5,6,7 are numeric, 8 is a date type
    
    Dim filterArr As New clsArray, controlVal
    Dim ctlName: ctlName = "fltr" & fieldName
    fieldName = GetSQLFieldName(fieldName, rsName)
    Dim ctl As control
    ''Get the recordsetclone of the subform if IsDynamicList
    Dim filterStatement ''shall be used for the hasOptions filter fields
    If IsDynamicList Then
        Dim subform As Form: Set subform = frm(ctlName).Form
        Dim recordSource: recordSource = subform.recordSource
        Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM " & recordSource & " WHERE SELECTED")
        Do Until rs.EOF
            controlVal = rs.fields("Value")
            ''If hasOption is true then do a lookup from the tblFilterFieldOptions and look for the FilterStatement
            If hasOption Then
                filterStatement = ELookup("tblFilterFieldOptions", "FilterFieldOptionID = " & controlVal, "FilterStatement")
                filterArr.Add filterStatement
            Else
                ''Check if Selected then store in an array (the value)
                ''Get the recordSource -> e.g. tblModelButtons
                Dim formRecordSource: formRecordSource = frm("subform").Form.recordSource
                If IsNull(ParentModelID) Then
                    controlVal = EscapeString(controlVal)
                End If
                filterArr.Add formRecordSource & "." & fieldName & "=" & controlVal
            End If
            rs.MoveNext
        Loop
    Else
        ''If option group then trace the value from the option group
        ''loop trough each option and get the tag of the optionvalue that matches??
        If IsOptionButton Then
            Dim ogValue: ogValue = frm(ctlName)
            For Each ctl In frm.controls
                If ctl.Name Like ctlName & "_*" Then
                    If ctl.optionValue = ogValue Then
                        controlVal = ctl.Tag
                        ''If hasOption is true then use the filterstatement from that table
                        ''Do a lookup from that table using the controlVal variable
                        If hasOption Then
                            filterStatement = ELookup("tblFilterFieldOptions", "FilterFieldOptionID = " & controlVal, "FilterStatement")
                            filterArr.Add filterStatement
                        Else
                            If Not Numbers.InArray(FieldTypeID) Then
                                controlVal = EscapeString(controlVal)
                            End If
                            
                            filterArr.Add fieldName & "=" & controlVal
                        End If
                        
                    End If
                End If
            Next ctl
        Else
            ''If not then loop through each controls having the name of the filtercontrol
            For Each ctl In frm.controls
                If ctl.Name Like ctlName & "_*" Then
                    
                    If ctl.value Then
                        If hasOption Then
                            controlVal = ctl.Tag
                            filterStatement = ELookup("tblFilterFieldOptions", "FilterFieldOptionID = " & controlVal, "FilterStatement")
                            filterArr.Add filterStatement
                        Else
                            controlVal = ExtractFilterValue(ctl.Name)
                            If IsNull(ParentModelID) Then
                                controlVal = EscapeString(controlVal)
                            End If
                            filterArr.Add fieldName & "=" & controlVal
                        End If
                        
                    End If
                    
                End If
            Next ctl
        End If
    End If
    
    Dim filterStr As String
    If filterArr.count > 0 Then
        filterStr = "(" & filterArr.JoinArr(" OR ") & ")"
    End If
    
    GetCbOrOGFilterStr = filterStr
       
End Function

Public Sub FilterControlSetCommonProperties(ctl As control, frm As Form, Optional DataEntryMode As Boolean = False)
    
    ctl.HorizontalAnchor = acHorizontalAnchorRight
    CopyControlTemplateProperties frm, ctl.Name, False
    If ctl.ControlType = acLabel Then
        Select Case ctl.Caption
            Case "TO", "Month", "Year":
            Case "ALL", "YES", "NO":
                ctl.TextAlign = 2
            Case Else:
                CopyProperties frm, ctl.Name, IIf(DataEntryMode, "LabelControl", "FilterLabelControl"), False
        End Select
    End If
    
End Sub

Public Function CopyFromToToDate(frm As Object, ctlName)
    
    Dim ctlTo As control, ctlFrom As control
    Set ctlTo = frm(ctlName & "To")
    Set ctlFrom = frm(ctlName & "From")
    
'    If IsNull(ctlTo) Then
'        ctlTo = ctlFrom
'    Else
'        If ctlTo < ctlFrom Then
'            ctlTo = ctlFrom
'        End If
'    End If
    
    If ctlTo < ctlFrom Or isFalse(ctlTo) Or isFalse(ctlFrom) Then
        ctlTo = ctlFrom
    End If
    
End Function

Public Function CopyFromToIfEarlier(frm As Object, ctlName)

    Dim ctlTo As control, ctlFrom As control
    Set ctlTo = frm(ctlName & "To")
    Set ctlFrom = frm(ctlName & "From")
    
'    If IsNull(ctlFrom) Then
'        ctlFrom = ctlTo
'    Else
'        If ctlFrom > ctlTo Then
'            ctlFrom = ctlTo
'        End If
'    End If
    
    If ctlFrom > ctlTo Or isFalse(ctlTo) Or isFalse(ctlFrom) Then
        ctlFrom = ctlTo
    End If
    
End Function

Private Function CreateControlHelper(frmName As String, ctlType As Integer, ctlName As String, ctlTop As Integer, ctlLeft As Integer, ctlWidth As Integer) As control
    Set CreateControlHelper = CreateControl(frmName, ctlType, , , , ctlLeft, ctlTop, ctlWidth)
    With CreateControlHelper
        .Name = ctlName
        .FontName = "Arial"
        .fontSize = 10
        .ForeColor = vbBlack
    End With
End Function

Private Function GetLargeProportionWidth(proportionArr As clsArray) As Double
    
    Dim proportion
    For Each proportion In proportionArr.arr
        If CDbl(proportion) >= 10 Then
            GetLargeProportionWidth = GetLargeProportionWidth + CDbl(proportion)
        End If
    Next proportion
    
End Function


Public Function RepositionControlsInRow(frm As Object, proportionStr, ControlNames, x, totalWidth, Optional gap As Integer = 0, Optional startingY As Variant = Null, Optional forceY As Boolean = False, Optional yAllowance As Integer = 100)
    
    Dim isAReport: isAReport = IsObjectAReport(frm)
    Dim proportionArr As New clsArray: proportionArr.arr = proportionStr
    Dim controlArr As New clsArray: controlArr.arr = ControlNames
    
    Dim proportionTotal As Double, totalWidthExcludingLarge As Double
    Dim hasLargeProportion As Boolean: hasLargeProportion = False
    
    Dim proportion As Variant
    For Each proportion In proportionArr.arr
        If CDbl(proportion) >= 10 Then
            hasLargeProportion = True
        Else
            proportionTotal = proportionTotal + CDbl(proportion)
        End If
    Next proportion
    
    If Not hasLargeProportion And proportionTotal = 0 Then
        Exit Function ' no proportion to distribute
    End If

    totalWidthExcludingLarge = totalWidth
    If hasLargeProportion Then
        totalWidthExcludingLarge = totalWidth - GetLargeProportionWidth(proportionArr)
    End If
    
    Dim startX As Integer: startX = x
    
    Dim y As Integer: y = GetMaxY(frm, , startX, totalWidth) + yAllowance
    
    If Not forceY Then
        If Not IsNull(startingY) And y < startingY Then
            y = startingY
        End If
    Else
        y = startingY
    End If
    
    Dim totalGap As Integer
    If proportionArr.count > 0 Then
        totalGap = gap * (proportionArr.count - 1)
    End If
    
    Dim i As Integer, calculatedWidth As Integer: calculatedWidth = 0
    For i = 0 To proportionArr.count - 1
        proportion = CDbl(proportionArr.arr(i))
        Dim ctl As control
        If proportion >= 10 Then ' fixed width
            Set ctl = frm(controlArr.arr(i))
            ctl.Left = startX
            ctl.Top = y
            If ctl.ControlType = acCheckBox Or ctl.ControlType = acOptionButton Then
                ctl.Top = ctl.Top + 40
            End If
            ctl.Width = proportion
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            calculatedWidth = calculatedWidth + ctl.Width
            startX = startX + ctl.Width + gap
        Else ' proportional width
            Dim toBeDeductedFromWidth: toBeDeductedFromWidth = totalGap / proportionArr.count
            Dim controlWidth: controlWidth = ((totalWidthExcludingLarge) * proportion / proportionTotal) - toBeDeductedFromWidth
            If controlWidth < 0 Then controlWidth = 0 ' avoid negative width
            
            Set ctl = frm(controlArr.arr(i))
            ctl.Left = startX
            ctl.Top = y
            If ctl.ControlType = acCheckBox Or ctl.ControlType = acOptionButton Then
                ctl.Top = ctl.Top + 40
            End If
            ctl.Width = controlWidth
            If Not isAReport Then ctl.HorizontalAnchor = acHorizontalAnchorRight
            calculatedWidth = calculatedWidth + ctl.Width
            startX = startX + ctl.Width + gap
        End If
    Next i


End Function
    
    
Private Function GetBooleanFilterControlNames(ctlName, OptionStr) As String
    
    Dim ControlNameArr As New clsArray
    Dim optionItem, OptionArr As New clsArray: OptionArr.arr = OptionStr
    
    For Each optionItem In OptionArr.arr
        ControlNameArr.Add "lbl" & ctlName & Trim(optionItem)
        ControlNameArr.Add ctlName & Trim(optionItem)
    Next optionItem
    
    GetBooleanFilterControlNames = ControlNameArr.JoinArr(",")
    
End Function

Public Sub RenderCheckboxAndLabel(frm As Object, ctlName, optionValue As Integer, optionText As String, x, y, totalWidth As Integer)
    Dim ctl As control
    ''Render the checkbox
    Set ctl = CreateControl(frm.Name, acOptionButton, , "og" & ctlName, , 0, 0, totalWidth)
    ctl.Name = ctlName & optionText
    ctl.optionValue = optionValue
    FilterControlSetCommonProperties ctl, frm
    ''Render the Label
    Set ctl = CreateControl(frm.Name, acLabel, , ctlName & optionText, , x, y, totalWidth)
    ctl.Name = "lbl" & ctlName & optionText
    ctl.Caption = optionText
    FilterControlSetCommonProperties ctl, frm
End Sub

Public Function GetPrimaryKeyFromTable(ModelID, Optional tblName = Null) As String
    
    If isFalse(tblName) Then
        tblName = GetTableNameFromModelID(ModelID)
    End If

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.field

    Set db = CurrentDb
    
    If DoesPropertyExists(db.TableDefs, tblName) Then
        Set tdf = db.TableDefs(tblName)
    Else
        GetPrimaryKeyFromTable = ELookup("tblModels", "ModelID = " & ModelID, "Model") & "ID"
        Exit Function
    End If
    
    Dim ind As Index

    For Each ind In tdf.indexes
        If ind.Primary Then
            GetPrimaryKeyFromTable = ind.fields(0).Name
            Exit Function
        End If
    Next ind
    
    GetPrimaryKeyFromTable = ""
End Function

Private Sub RenderBasedOnPossibleValues(frm As Object, ctlName, xPosition, totalWidth, possibleValues, parent, ctlType, Direction, proportion)
    
    
    Dim possibleValuesArr As New clsArray: possibleValuesArr.arr = possibleValues
    Dim item
    Dim i As Integer: i = 1
    
    Dim pair As New clsArray: pair.arr = proportion
    Dim ControlNames As New clsArray, proportions As New clsArray
    
    For Each item In possibleValuesArr.arr
        
        Dim cbName: cbName = ctlName & "_" & item
        Dim lblName: lblName = "lbl" & cbName
        Dim ctl As control

        ''Render the checkbox
        Set ctl = CreateControl(frm.Name, ctlType, , parent, , 0, 0, 0)
        ctl.Name = cbName
        ''This means, if there is no parent then it must be a checkBox
        If Not isFalse(parent) Then
            ctl.optionValue = i
            i = i + 1
            ctl.Tag = item
        Else
            ctl.DefaultValue = False
        End If
        
        FilterControlSetCommonProperties ctl, frm
        ''Render the Label
        Set ctl = CreateControl(frm.Name, acLabel, , cbName, , 0, 0, 0)
        ctl.Name = lblName
        ctl.Caption = item
        CopyProperties frm, lblName, "LabelControl", False
        
        If Direction = "Vertical" Then
            ''Clear the controlnames and proportions variable first
            ControlNames.clearArr
            proportions.clearArr
        End If
        
        ControlNames.Add cbName
        ControlNames.Add lblName
        proportions.Add pair.items(0)
        proportions.Add pair.items(1)
        
        ''Reposition the control
        If Direction = "Vertical" Then
            RepositionControlsInRow frm, proportions.JoinArr, ControlNames.JoinArr, xPosition, totalWidth, 0, frm("subform").Top, , 50
        End If
        
    Next item
    
    If Direction = "Horizontal" Then
        RepositionControlsInRow frm, proportions.JoinArr, ControlNames.JoinArr, xPosition, totalWidth, 0, frm("subform").Top, , 50
    End If
    
End Sub

Private Sub RenderCheckboxListFilterControls(frm As Object, ctlName, xPosition, totalWidth, FilterFieldID, Optional optionGroupMode As Boolean = False)

    Dim ModelFieldID: ModelFieldID = ELookup("tblFilterFields", "FilterFieldID = " & FilterFieldID, "ModelFieldID")
    ''Get the name and model id from tblModelFields
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & ModelFieldID)
    Dim ModelID: ModelID = rs.fields("ModelID")
    Dim ModelField: ModelField = rs.fields("ModelField")
    Dim ParentModelID: ParentModelID = rs.fields("ParentModelID")
    Dim VerboseName: VerboseName = rs.fields("VerboseName")
    Dim possibleValues: possibleValues = rs.fields("PossibleValues")
    
    rs.Close
    
    ''Filter Caption
    Dim ctl As control
    Set ctl = CreateControl(frm.Name, acLabel, , , , 0, 0, 0)
    ctl.Name = "lbl" & ctlName
    ctl.Caption = GetFieldCaption(VerboseName, ModelField)
    CopyProperties frm, ctl.Name, "FilterLabelControl", False
    RepositionControlsInRow frm, "1", ctl.Name, xPosition, totalWidth, 0, frm("subform").Top
    
    Dim parent: parent = ""
    Dim ctlType: ctlType = acCheckBox
    
    ''If this is an option group then render the option group control
    If optionGroupMode Then
        parent = ctlName
        ctlType = acOptionButton
        ''Create the option group control
        Set ctl = CreateControl(frm.Name, acOptionGroup, , , , xPosition, 0, 0)
        ctl.Name = ctlName
        ctl.BorderStyle = 0
        ctl.SpecialEffect = 0
        ctl.HorizontalAnchor = acHorizontalAnchorRight
        
    End If
    
    ''Get the direction and proportion
    Dim Direction, proportion
    SetDirectionAndProportion Direction, proportion, FilterFieldID
    
    '''If possible values will override the ParentModelID check
    If Not IsNull(possibleValues) Then
        RenderBasedOnPossibleValues frm, ctlName, xPosition, totalWidth, possibleValues, parent, ctlType, Direction, proportion
        Exit Sub
    End If
    
    Dim TableName, sqlStr
    Dim hasFilterFieldOption: hasFilterFieldOption = ECount("tblFilterFieldOptions", "FilterFieldID = " & FilterFieldID) > 0
    ''If the ModelField is just a foreign key fetch the parent model's MainField
    If Not IsNull(ParentModelID) Then
        Dim MainField: MainField = ELookup("tblModels", "ModelID =" & ParentModelID, "MainField")
        Dim PrimaryKey: PrimaryKey = GetPrimaryKeyFromTable(ParentModelID)
        
        sqlStr = "SELECT " & PrimaryKey & "," & MainField & " AS MainField From " & TableName & " ORDER BY " & MainField
        
        ''If there's a filterfieldoption for this filterfieldID then use this sqlStr
    ElseIf hasFilterFieldOption Then
        sqlStr = "SELECT FilterFieldOptionID AS [Value], OptionCaption As Label FROM tblFilterFieldOptions WHERE FilterFieldID = " & FilterFieldID & " ORDER BY OptionOrder ASC"
    Else
        ''ctlName would be ctlName_id -> id to be used as filter
        TableName = GetTableNameFromModelID(ModelID)
        sqlStr = "SELECT " & ModelField & " From " & TableName & " GROUP BY " & ModelField & " ORDER BY " & ModelField
    End If
    
    Set rs = ReturnRecordset(sqlStr)
    Dim i As Integer: i = 1
    
    ''Get the size in pair of the cb and the label
    Dim pair As New clsArray: pair.arr = proportion
    Dim ControlNames As New clsArray, proportions As New clsArray
    Do Until rs.EOF
    
        Dim cbNameExtension: cbNameExtension = GetCheckBoxNameExtension(rs)
        Dim cbName: cbName = ctlName & "_" & cbNameExtension
        Dim cbLabelCaption: cbLabelCaption = GetCheckBoxLabelCaption(rs)
        Dim lblName: lblName = "lbl" & cbName

        ''Render the checkbox
        Set ctl = CreateControl(frm.Name, ctlType, , parent, , 0, 0, 0)
        ctl.Name = cbName
        
        If optionGroupMode Then
            ctl.optionValue = i
            i = i + 1
            ctl.Tag = GetCheckBoxNameExtension(rs)
        Else
            ctl.DefaultValue = False
        End If
        
        ''Overwrite the tag if hasFilterFieldOption. tag should be the id of the tblFilterFieldOptions
        If hasFilterFieldOption Then
            ctl.Tag = rs.fields(0) ''0 here is the value field of the recordset
        End If
        
        FilterControlSetCommonProperties ctl, frm
        ''Render the Label
        Set ctl = CreateControl(frm.Name, acLabel, , cbName, , 0, 0, 0)
        ctl.Name = lblName
        ctl.Caption = cbLabelCaption
        CopyProperties frm, lblName, "LabelControl", False
        
        If Direction = "Vertical" Then
            ''Clear the controlnames and proportions variable first
            ControlNames.clearArr
            proportions.clearArr
        End If
        
        ControlNames.Add cbName
        ControlNames.Add lblName
        proportions.Add pair.items(0)
        proportions.Add pair.items(1)
        
        ''Reposition the control -> If vertical then reposition as usual (one row per cb and label pairing)
        If Direction = "Vertical" Then
            RepositionControlsInRow frm, proportions.JoinArr, ControlNames.JoinArr, xPosition, totalWidth, 0, frm("subform").Top, , 50
        End If
        
        rs.MoveNext
    Loop
    
    If Direction = "Horizontal" Then
        RepositionControlsInRow frm, proportions.JoinArr, ControlNames.JoinArr, xPosition, totalWidth, 0, frm("subform").Top, , 50
    End If
       
End Sub

Private Sub SetDirectionAndProportion(Direction, proportion, FilterFieldID)

    ''default direction will be vertical, default proportion will be 1,9
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE FilterFieldID = " & FilterFieldID)
    Direction = rs.fields("Direction")
    proportion = rs.fields("Proportion")
    
    If isFalse(Direction) Then Direction = "Vertical"
    If isFalse(proportion) Then proportion = "1,9"
    
End Sub

Private Sub OverrideFromFilterFields(FilterFieldID, Caption)

    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE FilterFieldID = " & FilterFieldID)
    Dim FilterCaption: FilterCaption = rs.fields("FilterCaption")
    
    If Not isFalse(FilterCaption) Then
        Caption = FilterCaption
    End If
    
End Sub

Private Sub RenderCheckboxListFilterForm(frm As Object, ctlName, xPosition, totalWidth, FilterFieldID, Optional optionGroupMode As Boolean = False)

    Dim ModelFieldID: ModelFieldID = ELookup("tblFilterFields", "FilterFieldID = " & FilterFieldID, "ModelFieldID")
    ''Get the name and model id from tblModelFields
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & ModelFieldID)
    Dim ModelID: ModelID = rs.fields("ModelID")
    Dim ModelField: ModelField = rs.fields("ModelField")
    Dim ParentModelID: ParentModelID = rs.fields("ParentModelID")
    Dim VerboseName: VerboseName = rs.fields("VerboseName")
    Dim ModelName: ModelName = ELookup("tblModels", "ModelID = " & ModelID, "Model")

    rs.Close
    
    ''Set the subform height from the tblFilterFields
    Dim ControlHeight: ControlHeight = ELookup("tblFilterFields", "FilterFieldID = " & FilterFieldID, "ControlHeight")
    If isFalse(ControlHeight) Then ControlHeight = 3000
    
    ''Filter Caption
    Dim ctl As control, filterCaptionName, Caption
    Set ctl = CreateControl(frm.Name, acLabel, , , , 0, 0, 0)
    filterCaptionName = "lbl" & ctlName
    ctl.Name = filterCaptionName
    Caption = GetFieldCaption(VerboseName, ModelField)
    
    ''Override certain variables based on the conditions from tblFilterFields
    OverrideFromFilterFields FilterFieldID, Caption
    
    ctl.Caption = Caption
    CopyProperties frm, filterCaptionName, "FilterLabelControl", False
    
    ''Render the Textbox that will filter the subform
    Dim searchBoxName: searchBoxName = "txtSearch" & ctlName
    Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, 0)
    ctl.Name = searchBoxName
    ctl.Format = "@;""Search " & Caption & """"
    ctl.OnChange = "=FilterFilterSubform([Form], " & EscapeString(searchBoxName) & ", " & EscapeString(ctlName) & " )"
    CopyProperties frm, searchBoxName, "TextControl", False
    
    ''Render the clear button -> Will clear the searchBoxName
    Dim clearBtnName: clearBtnName = "clear" & searchBoxName
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0, 0)
    ctl.Name = clearBtnName
    ctl.Caption = "Clear"
    ctl.OnClick = "=ClearFilterFilterSubform([Form], " & EscapeString(searchBoxName) & ", " & EscapeString(ctlName) & " )"
    
    CopyProperties frm, clearBtnName, "ButtonControl", False
    ctl.Height = frm(searchBoxName).Height
    
    RepositionControlsInRow frm, "3,3,1", filterCaptionName & "," & searchBoxName & "," & clearBtnName, xPosition, totalWidth, 0, frm("subform").Top
    
    Dim ctlType: ctlType = acCheckBox
    If optionGroupMode Then
        ctlType = acOptionButton
    End If
    
    Dim TableName, sqlStr
    ''If the ModelField is just a foreign key fetch the parent model's MainField
    If Not IsNull(ParentModelID) Then
        Dim MainField: MainField = ELookup("tblModels", "ModelID =" & ParentModelID, "MainField")
        Dim PrimaryKey: PrimaryKey = GetPrimaryKeyFromTable(ParentModelID)
        TableName = GetTableNameFromModelID(ParentModelID)
        sqlStr = "SELECT " & PrimaryKey & "," & MainField & " AS MainField From " & TableName & " ORDER BY " & MainField
        
        ''ctlName would be ctlName_id -> id to be used as filter
        ''Check the filterFieldOption if there's some data
    ElseIf ECount("tblFilterFieldOptions", "FilterFieldID = " & FilterFieldID) > 0 Then
        sqlStr = "SELECT FilterFieldOptionID AS [Value], OptionCaption As Label FROM tblFilterFieldOptions WHERE FilterFieldID = " & FilterFieldID & " ORDER BY OptionOrder ASC"
    Else
        TableName = GetTableNameFromModelID(ModelID)
        sqlStr = "SELECT " & ModelField & " As [Value], " & ModelField & " AS Label From " & TableName & " GROUP BY " & ModelField & " ORDER BY " & ModelField
    End If
    
    ''Create the form here = must be a continious form with name of "cont" & ctlName
    CreateFilterContForm ctlName, VerboseName, optionGroupMode, sqlStr, ModelName
    
    ''Create the control here, must be a subform type. ControlSource would be "cont" & ctlName
    Set ctl = CreateControl(frm.Name, acSubform, , , , 0, 0, 0, 2000)
    ctl.Name = ctlName
    ctl.SourceObject = "cont" & ModelName & ctlName
    CopyProperties frm, ctl.Name, "SubformControl", False
    ctl.BorderStyle = 0
    ctl.Height = ControlHeight
    RepositionControlsInRow frm, "1", ctl.Name, xPosition, totalWidth, 0, frm("subform").Top
    
End Sub

Private Sub CreateFilterContForm(ctlName, VerboseName, optionGroupMode, sqlStr, ModelName)
    
    ''Create the table for the filter first -> tbl & ctlName
    Dim db As Database: Set db = CurrentDb
    Dim tblName: tblName = "tbl" & ModelName & ctlName
    Dim tblDef As TableDef: Set tblDef = AddTableDef(db, tblName)
    
    ''Pk id first so that the table would be valid
    CreatePrimaryKey "", tblDef, "ID"
    If Not DoesPropertyExists(db.TableDefs, tblName) Then
        db.TableDefs.Append tblDef
    End If
    
    ''id, checkbox, label, value
    Dim fld As DAO.field
    Set fld = AddField(tblDef, "FilterLabel", dbText)
    Set fld = AddField(tblDef, "Selected", dbBoolean)
    CreateProperty fld, "DisplayControl", dbInteger, acCheckBox
    Set fld = AddField(tblDef, "Value", dbText)
    ''Then create the form with that recordset
    ''Create the form here = must be a continious form with name of "cont" & ctlName
    Dim frm As Form: Set frm = CreateForm
    Dim frmName: frmName = "cont" & ModelName & ctlName
    frm.DefaultView = acDefViewContinuous
    frm.recordSource = tblName
    
    ''Create the form controls
    Dim ctlType: ctlType = acCheckBox
    If optionGroupMode Then
        ctlType = acOptionButton
    End If
    ''Create the option or checkbox
    Dim ctl As control
    Set ctl = CreateControl(frm.Name, ctlType, , , , 0, 0, 0)
    ctl.ControlSource = "Selected": ctl.Name = "Selected"
    
    ''Create the transparent button
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0, 0)
    ctl.Name = "cmdToggleValue"
    CopyProperties frm, ctl.Name, "TransparentButton", False
    
    ''Create the label
    Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, 0)
    ctl.ControlSource = "FilterLabel": ctl.Name = "FilterLabel"
    CopyProperties frm, ctl.Name, "LabelControl", False
    RepositionControlsInRow frm, "2,9", "Selected,FilterLabel", 100, 1440, 50, 0
    ''Lock and disable
    ctl.Locked = True
    ctl.Enabled = False
    
    ''Set the label and cbs top position and anchoring
    frm.controls("Selected").Top = 40
    frm.controls("FilterLabel").Top = 0
    ''Reposition the cmdToggleValue to match the FilterLabel position
    frm.controls("cmdToggleValue").Top = frm.controls("FilterLabel").Top
    frm.controls("cmdToggleValue").Left = frm.controls("FilterLabel").Left
    frm.controls("cmdToggleValue").Width = frm.controls("FilterLabel").Width
    frm.controls("cmdToggleValue").Height = frm.controls("FilterLabel").Height
    frm.controls("cmdToggleValue").InSelection = True
    DoCmd.RunCommand acCmdBringToFront
    ''Anchor should be left and both
    frm("Selected").HorizontalAnchor = acHorizontalAnchorLeft
    frm("FilterLabel").HorizontalAnchor = acHorizontalAnchorBoth
    frm("cmdToggleValue").HorizontalAnchor = acHorizontalAnchorBoth
    ''Detail Height must be zero
    frm.Section(acDetail).Height = 0
    ''Width should be 1440 -> 0 to autofit
    frm.Width = 0
    CopyControlTemplateProperties frm
    ''RecordSelector and NavigationButtons must be False
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    
    ''Add on Load event that will delete all record and requery this record to reflect the existing records
    ''FilterContFormOnLoad(frm As Object, sqlStr As String, tblName As String)
    frm.OnLoad = "=FilterContFormOnLoad([Form]," & EscapeString(sqlStr) & "," & EscapeString(tblName) & ")"
    ''If optiongroup uncheck all then just select the existing one.
    ''Attach event to the Selected control
    ''If checkbox then leave as is.
    If optionGroupMode Then
        frm("Selected").AfterUpdate = "=FilterContOptionOnChange([Form]," & EscapeString(tblName) & ")"
    End If
    
    ''Attach event to the button
    frm("cmdToggleValue").OnClick = "=ToggleFilterCB([Form]," & EscapeString(tblName) & ")"
    
    Dim OriginalFormName: OriginalFormName = frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename frmName, acForm, OriginalFormName
    
End Sub

Private Function GetCheckBoxLabelCaption(rs As Recordset)
On Error GoTo ErrHandler:
    GetCheckBoxLabelCaption = rs.fields(1)
    Exit Function
ErrHandler:
    GetCheckBoxLabelCaption = rs.fields(0)
End Function

Private Function GetCheckBoxNameExtension(rs As Recordset)
On Error GoTo ErrHandler:
    GetCheckBoxNameExtension = rs.fields(1)
    Exit Function
ErrHandler:
    GetCheckBoxNameExtension = rs.fields(0)
End Function

Private Function GetFilterFieldUnionSQL(ModelID) As String
    
    ''Get the FilterID, Order, FilterType
    Dim sqlStr: sqlStr = "SELECT FilterFieldID, FilterOrder, ""MainTable"" AS FilterType FROM tblFilterFields WHERE ModelID = " & ModelID & " AND " & _
        " Not IsWildSearch And FilterOrder > 0 ORDER BY FilterOrder UNION ALL "
    sqlStr = sqlStr & "SELECT RelatedFilterFieldID, FilterOrder, ""ForeignTable"" As FilterType FROM tblRelatedFilterFields WHERE ModelID = " & ModelID & " ORDER BY FilterOrder"
    
    GetFilterFieldUnionSQL = sqlStr
    
End Function

Private Function CheckForFilterField(ModelID) As Boolean

    Dim recordCount: recordCount = ECount("tblFilterFields", "ModelID = " & ModelID) + ECount("tblRelatedFilterFields", "ModelID = " & ModelID)
    CheckForFilterField = recordCount > 0
    
End Function

Public Function GetModelFormName(modelRS As Recordset, Optional frmType = "main") As String
    
    Dim VerbosePlural: VerbosePlural = modelRS.fields("VerbosePlural")
    Dim SubformName: SubformName = modelRS.fields("SubformName")
    Dim Model: Model = modelRS.fields("Model"): If ExitIfTrue(isFalse(Model), "Model is empty..") Then Exit Function
    
    If Not isFalse(SubformName) Then
        GetModelFormName = frmType & SubformName
        Exit Function
    End If
    
    If Not isFalse(VerbosePlural) Then
        GetModelFormName = frmType & VerbosePlural
        Exit Function
    End If
    
    GetModelFormName = frmType & Model & "s"
    
End Function


Private Sub SetSearchComboboxProperties(ctl As control, frm As Form, modelRS As Recordset)

    ctl.Properties("AllowValueListEdits") = False
    
    Dim frmName: frmName = GetModelFormName(modelRS)
    Dim sqlStr: sqlStr = "SELECT TOP 10 SearchTerm FROM tblSearchHistorys WHERE FormName = " & Esc(frmName) & " GROUP BY SearchTerm"
    ctl.RowSource = sqlStr
    
End Sub

Private Sub RenderFilterForm(frm As Object, ModelID, Optional DataEntryMode As Boolean = False, Optional frmName As String = "")
    
    ''Fetch the unionSQL of both types of filter
    Dim unionSQL: unionSQL = GetFilterFieldUnionSQL(ModelID)
    
    ''Fetch the FilterFields and the Model properties
    Dim rs As Recordset, modelRS As Recordset, modelFieldRS As Recordset, ctl As control
    Set modelRS = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    
    ''Check first if the filter is empty (meaning there's no filter form to be rendered)
    If Not CheckForFilterField(ModelID) Then Exit Sub
    
    ''Filter Form Property
    Dim totalWidth, colSpaceWidth, x, y, originalY, ctlHeight
    ''1 inch = 1440 twips (all measurement is in twips)
    totalWidth = ELookup("tblSupplementalModels", "ModelID = " & ModelID, "FilterFormWidth")
    If isFalse(totalWidth) Then totalWidth = 3500
    colSpaceWidth = 50
    ctlHeight = 300
    
    If DataEntryMode Then
        y = InchToTwip(0.25)
        x = 400
    Else
        y = frm("subform").Top
        x = frm("subform").Left + frm("subform").Width + 100
    End If
        
    originalY = y
    
    If DataEntryMode Then
        Set ctl = CreateControl(frm.Name, acLabel, , , , x, y, totalWidth)
        FilterControlSetCommonProperties ctl, frm, DataEntryMode
        ctl.Caption = "Search " & GetModelPlural(modelRS.fields("Model"), modelRS.fields("VerbosePlural"), "", modelRS.fields("VerbosePluralCaption"))
    End If
    
    ''Check wether a wildcard filter is present from the tblFilterFields and tblRelatedFilterFields.
    Dim isWildcardPresent: isWildcardPresent = CheckWildcardPresence(ModelID)
    If isWildcardPresent Then
        ''Render the wildcard seach box here the text box first
        ''Set ctl = CreateControl(frm.name, acTextBox, , , , x, y + ctlHeight + 100, totalWidth)
        Set ctl = CreateControl(frm.Name, acComboBox, , , , x, y + ctlHeight + 100, totalWidth)
        FilterControlSetCommonProperties ctl, frm
        SetSearchComboboxProperties ctl, frm, modelRS
        ctl.Name = "fltrWildSearch"
        ''Then the label
        Set ctl = CreateControl(frm.Name, acLabel, , "fltrWildSearch", , x, y, totalWidth)
        FilterControlSetCommonProperties ctl, frm
        ctl.Caption = "Search " & GetModelPlural(modelRS.fields("Model"), modelRS.fields("VerbosePlural"), "", modelRS.fields("VerbosePluralCaption"))
    End If
    
    ''Render other filter controls Not isWildSearch
    Dim proportionArr As New clsArray, proportionTotal, controlArr As New clsArray, i As Integer, proportion As Double, startX
    
    ''FilterType => can be MainTable or ForeignTable
    ''FilterFieldID, FilterOrder
    Set rs = ReturnRecordset(unionSQL)
    
    Dim filterType, FilterFieldID
    Do Until rs.EOF
        ''Check the type if MainTable do the usual else do the foreignTable
        filterType = rs.fields("FilterType")
        FilterFieldID = rs.fields("FilterFieldID")
        If filterType = "MainTable" Then
            ''RenderMainFilter(FilterFieldID, frm As Form, x, y, originalY, totalWidth, ctlHeight)
            RenderMainFilter FilterFieldID, frm, x, y, originalY, totalWidth, ctlHeight, DataEntryMode, frmName
        Else
            RenderForeignFilter FilterFieldID, frm, x, y, originalY, totalWidth, ctlHeight
        End If
        rs.MoveNext
    Loop
    
    ''Render the Filter buttons
    ''Filter and Clear
    If Not DataEntryMode Then
        proportionArr.arr = "6,6"
        controlArr.arr = "cmdFilter,cmdClear"
        proportionTotal = GetProportionTotal(proportionArr)
        NewRow frm, x, y, originalY, totalWidth
        
        RenderButton 0, 0, "Filter", frm, "Filter"
        ''Check if there's a related table for this ModelID. If there's one then
        Dim useQuery: useQuery = "False"
        If ECount("tblRelatedFilterFields", "ModelID = " & ModelID) = 0 Then
            frm("cmdFilter").OnClick = "=FilterSubform([Form]," & ModelID & ")"
        Else
            useQuery = "True"
            frm("cmdFilter").OnClick = "=SetSubformSQL([Form]," & ModelID & ")"
            frm.OnLoad = "=MainFormSQLOnLoad([Form]," & ModelID & ")"
        End If
        
        RenderButton 0, 0, "Clear Filter", frm, "Clear"
        frm("cmdClear").OnClick = "=ClearFilterSubform([Form]," & ModelID & "," & useQuery & ")"
        
        RepositionControlsInRow frm, "6,6", "cmdFilter,cmdClear", x, totalWidth, 100
    End If
    
    ''Add a line if DataEntryMode
    If DataEntryMode Then
        Dim FormColumns: FormColumns = ELookup("tblModels", "ModelID = " & ModelID, "FormColumns")
        Dim lineWidth: lineWidth = 3000 * (FormColumns) + (200 * (FormColumns - 1))
        Set ctl = CreateControl(frm.Name, acLine, , , , 400, GetMaxY(frm) + 200, lineWidth, 0)
        ctl.BorderStyle = 4
        ctl.BorderWidth = 2
    End If
    
End Sub

Public Function CheckWildcardPresence(ModelID) As Boolean

    CheckWildcardPresence = isPresent("tblFilterFields", "ModelID = " & ModelID & " AND IsWildSearch")
    
    If Not CheckWildcardPresence Then
        CheckWildcardPresence = isPresent("tblRelatedFilterFields", "ModelID = " & ModelID & " AND IncludeInWildcardSearch")
    End If
    
End Function

Public Function FilterFilterSubform(frm As Object, searchBoxName, SubformName)
    
    frm(searchBoxName).SetFocus
    Dim searchString: searchString = frm(searchBoxName).Text
    
    Dim filterStr As String
    If Not isFalse(searchString) And Not frm(searchBoxName).Format Like "*" & searchString & "*" Then
        filterStr = "FilterLabel Like " & EscapeString("*" & searchString & "*")
    End If
    
    Dim subfrm As Form: Set subfrm = frm(SubformName).Form
    
    If filterStr <> "" Then
        subfrm.filter = filterStr
        subfrm.FilterOn = True
    Else
        subfrm.FilterOn = False
    End If
    
End Function

Public Function ClearFilterFilterSubform(frm As Object, searchBoxName, SubformName)
    
    frm(searchBoxName) = Null
    FilterFilterSubform frm, searchBoxName, SubformName
    
End Function

Private Sub RenderPresenceFilter(FilterFieldID, frm As Object, x, totalWidth)

    ''x would be the starting position, totalWidth will be the width of the controls combined
    ''Render the checkbox and label -> cbName would be "fltr" & TableName & SatisfiesFilter
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblRelatedFilterFields WHERE RelatedFilterFieldID = " & FilterFieldID)
    ''TABLE: tblRelatedFilterFields Fields: RelatedFilterFieldID|ModelID|TableName|MainConnectorField|SubConnectorField
    ''FieldToUse|Timestamp|CreatedBy|RecordImportID|FilterOrder|IsOptionGroup|FilterOperation|LeftJoinKey|RightJoinKey
    ''IncludeInWildcardSearch|ShowPresent|FilterCaption|FilterString|SatisfiesFilter
    Dim SatisfiesFilter: SatisfiesFilter = rs.fields("SatisfiesFilter")
    Dim TableName: TableName = rs.fields("TableName")
    Dim FilterCaption: FilterCaption = rs.fields("FilterCaption")
    
    Dim cbName: cbName = "fltr" & TableName & SatisfiesFilter
    Dim lblName: lblName = "lbl" & cbName
    
    ''Render the checkbox
    Dim ctl As control
    Set ctl = CreateNamedControl(frm, acCheckBox, cbName)
    ctl.Name = cbName
    ctl.DefaultValue = False
    CopyProperties frm, cbName, "CheckboxControl", False
    
    ''Render the label
    Set ctl = CreateNamedControl(frm, acLabel, lblName, cbName)
    ctl.Name = lblName
    ctl.Caption = FilterCaption
    CopyProperties frm, lblName, "LabelControl", False
    
    RepositionControlsInRow frm, "1,8", cbName & "," & lblName, x, totalWidth
    
End Sub

Private Sub RenderForeignFilter(FilterFieldID, frm As Form, x, y, originalY, totalWidth, ctlHeight)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblRelatedFilterFields WHERE RelatedFilterFieldID = " & FilterFieldID)
    ''TABLE: tblRelatedFilterFields Fields: RelatedFilterFieldID|ModelID|TableName|MainConnectorField|SubConnectorField
    ''FieldToUse|Timestamp|CreatedBy|RecordImportID|FilterOrder
    
    ''How should we name the form? -> cont[Model][ctlName]
    Dim ModelID, TableName, MainConnectorField, SubConnectorField, FieldToUse, FilterOrder, IsOptionGroup
    Dim IncludeInWildcardSearch, ControlHeight
    
    ModelID = rs.fields("ModelID")
    TableName = rs.fields("TableName")
    MainConnectorField = rs.fields("MainConnectorField")
    SubConnectorField = rs.fields("SubConnectorField")
    FieldToUse = rs.fields("FieldToUse")
    FilterOrder = rs.fields("FilterOrder")
    IsOptionGroup = rs.fields("IsOptionGroup")
    IncludeInWildcardSearch = rs.fields("IncludeInWildcardSearch")
    ControlHeight = rs.fields("controlHeight")
    
    If isFalse(ControlHeight) Then ControlHeight = 3000
    
    ''If IncludeInWildcardSearch then exit sub
    If IncludeInWildcardSearch Then Exit Sub
    
    ''Hijack if FieldToUse is Null --> meaning this is a simple checkbox with a label field -->
    If isFalse(FieldToUse) Then
        RenderPresenceFilter FilterFieldID, frm, x, totalWidth
        Exit Sub
    End If
    
    ''ctlName = "fltr" & modelFieldRs.fields("ModelField")
    Dim ctlName: ctlName = "fltr" & FieldToUse
    Dim Model: Model = ELookup("tblModels", "ModelID = " & ModelID, "Model")
    Dim frmName: frmName = "cont" & Model & ctlName
    
    ''Render the filter Caption
    Dim Caption: Caption = GetCaptionPropertyFromTable(TableName, FieldToUse)
    Dim ctl As control, lblName: lblName = "lbl" & ctlName
    Set ctl = CreateControl(frm.Name, acLabel, , , , 0, 0, 0)
    ctl.Name = lblName
    ctl.Caption = Caption
    
    CopyProperties frm, ctl.Name, "FilterLabelControl", False
    
    Dim searchBoxName: searchBoxName = "txtSearch" & ctlName
    Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, 0)
    ctl.Name = searchBoxName
    ctl.Format = "@;""Search " & Caption & """"
    ctl.OnChange = "=FilterFilterSubform([Form], " & EscapeString(searchBoxName) & ", " & EscapeString(ctlName) & " )"
    CopyProperties frm, searchBoxName, "TextControl", False
    
    ''Render the clear button -> Will clear the searchBoxName
    Dim clearBtnName: clearBtnName = "clear" & searchBoxName
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0, 0)
    ctl.Name = clearBtnName
    ctl.Caption = "Clear"
    ctl.OnClick = "=ClearFilterFilterSubform([Form], " & EscapeString(searchBoxName) & ", " & EscapeString(ctlName) & " )"
    
    CopyProperties frm, clearBtnName, "TransparentButton", False
    ctl.Height = frm(searchBoxName).Height
    
    RepositionControlsInRow frm, "3,3,1", lblName & "," & searchBoxName & "," & clearBtnName, x, totalWidth, 0, frm("subform").Top
    
    ''Get the sqlStr of the filter form
    Dim sqlStr
    ''If the field to use has id then get the value from the table of the foreign key
    ''e.g. CardKeywordID will look at tblCardKeywords
    If FieldToUse Like "*ID" Then
        ''Get the ModelID using the PrimaryKey
        Dim SubModelRS As Recordset: Set SubModelRS = GetModelByPrimaryKey(FieldToUse)
        Dim MainField: MainField = SubModelRS.fields("MainField")
        Dim PrimaryKey: PrimaryKey = GetPrimaryKeyFromTable(SubModelRS.fields("ModelID"))
        Dim SubTableName: SubTableName = GetTableNameByPrimaryKey(FieldToUse)
        sqlStr = "SELECT " & PrimaryKey & "," & MainField & " AS MainField From " & SubTableName & " ORDER BY " & MainField
    Else
        sqlStr = "SELECT " & FieldToUse & " As [Value], " & FieldToUse & " AS Label From " & TableName & " GROUP BY " & FieldToUse & " ORDER BY " & FieldToUse
    End If
    
    CreateFilterContForm ctlName, Caption, IsOptionGroup, sqlStr, Model
    
    ''Create the control here, must be a subform type. ControlSource would be "cont" & ctlName
    Set ctl = CreateControl(frm.Name, acSubform, , , , 0, 0, 0, 2000)
    ctl.Name = ctlName
    ctl.SourceObject = "cont" & Model & ctlName
    CopyProperties frm, ctl.Name, "SubformControl", False
    ctl.BorderStyle = 0
    ctl.Height = ControlHeight
    RepositionControlsInRow frm, "1", ctl.Name, x, totalWidth, 0, frm("subform").Top
    
End Sub


Public Function GetCaptionPropertyFromTable(tblName, fldName) As String
    
    Dim rsDef As Object, db As Database
    Set db = CurrentDb
    
    If DoesPropertyExists(CurrentDb.TableDefs, tblName) Then
        Set rsDef = db.TableDefs(tblName)
    Else
        Set rsDef = db.QueryDefs(tblName)
    End If
    
    Dim fld As field
    On Error Resume Next
    Set fld = rsDef.fields(fldName)
    
    ''Get the Caption property
    On Error Resume Next
    GetCaptionPropertyFromTable = fld.Properties("Caption")
    
End Function

Private Sub RenderMainFilter(FilterFieldID, frm As Form, x, y, originalY, totalWidth, ctlHeight, Optional DataEntryMode As Boolean = False, Optional frmName As String = "")
    
    Dim rs As Recordset, modelFieldRS As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE FilterFieldID = " & FilterFieldID)
    Dim startX, ctlName, ctl As control, lblName, ctrlArr As New clsArray
    
    NewRow frm, x, y, originalY, totalWidth
    Set modelFieldRS = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & rs.fields("ModelFieldID"))
    
    ''Boolean Filter
    If modelFieldRS.fields("FieldTypeID") = dbBoolean Then
        
        startX = x
        ctlName = "fltr" & modelFieldRS.fields("ModelField")

        ''Render the optionGroup control
        Set ctl = CreateControl(frm.Name, acOptionGroup, , , , 0, 0, totalWidth)
        ctl.Name = "og" & ctlName
        
        ctl.DefaultValue = 2
        ctl.BorderStyle = 0
        ctl.SpecialEffect = 0
        ctl.HorizontalAnchor = acHorizontalAnchorRight
        
        ''Render the filter label
        Set ctl = CreateControl(frm.Name, acLabel, , , , x, y, totalWidth)
        FilterControlSetCommonProperties ctl, frm
        ctl.Caption = GetFieldCaption(modelFieldRS.fields("VerboseName"), modelFieldRS.fields("ModelField"))
        
        ''Render the ALL
        RenderCheckboxAndLabel frm, ctlName, 2, "ALL", x, y, 0
        ''Render the YES
        RenderCheckboxAndLabel frm, ctlName, -1, "YES", x, y, 0
        ''Render the NO
        RenderCheckboxAndLabel frm, ctlName, 0, "NO", x, y, 0
        
        ctl.Top = y
        ctl.Left = x
        
        ''Repostion the controls
        RepositionControlsInRow frm, "2,2,2,2,2,2", GetBooleanFilterControlNames(ctlName, "ALL,YES,NO"), startX, totalWidth, 50
        GoTo NextFilter:
    End If
    
    
    If rs.fields("IsCheckboxList") Then
        ''Run NewRow each for each checkbox
        ''If FilterCaption is null then use the modelFieldRS.fields("ModelField")
        If isPresent("tblFilterFields", "FilterFieldID = " & FilterFieldID & " AND Not isNull(FilterCaption)") Then
            ctlName = "fltr" & ELookup("tblFilterFields", "FilterFieldID = " & FilterFieldID, "FilterCaption")
        Else
            ctlName = "fltr" & modelFieldRS.fields("ModelField")
        End If
            
        If rs.fields("IsDynamicList") Then
            RenderCheckboxListFilterForm frm, ctlName, x, totalWidth, rs.fields("FilterFieldID")
        Else
            RenderCheckboxListFilterControls frm, ctlName, x, totalWidth, rs.fields("FilterFieldID")
        End If
        
        GoTo NextFilter:
    End If
    
    If rs.fields("IsOptionGroup") Then
        ''Run NewRow each for each checkbox
        ctlName = "fltr" & modelFieldRS.fields("ModelField")
        If rs.fields("IsDynamicList") Then
            RenderCheckboxListFilterForm frm, ctlName, x, totalWidth, rs.fields("FilterFieldID"), True
        Else
            RenderCheckboxListFilterControls frm, ctlName, x, totalWidth, rs.fields("FilterFieldID"), True
        End If
        GoTo NextFilter:
    End If
    
    If rs.fields("IsList") Then
        ''Render the combo box here first
        Dim comboY: comboY = IIf(DataEntryMode, GetMaxY(frm, , , , , InchToTwip(0.25)) + 100, y + ctlHeight + 100)
        Set ctl = CreateControl(frm.Name, acComboBox, , , , x, comboY, totalWidth)
        FilterControlSetCommonProperties ctl, frm
        
        ctlName = "fltr" & modelFieldRS.fields("ModelField")
        ctl.Name = ctlName
        
        ''FilterDataEntryForm frmName, fieldName, FilterControlName
        If DataEntryMode Then
            ctl.AfterUpdate = "=FilterDataEntryForm(" & Esc(frmName) & "," & Esc(modelFieldRS.fields("ModelField")) & "," & Esc(ctlName) & ")"
        End If
        
        If IsNull(modelFieldRS.fields("PossibleValues")) Then
            SetComboBoxSQLForFilter ctl, rs.fields("ModelFieldID")
            
        Else
            ctl.RowSource = Join(Split(modelFieldRS.fields("PossibleValues"), ","), ";")
            ctl.ColumnCount = 1
            ctl.ColumnWidths = "1"
            ctl.RowSourceType = "Value List"
            ctl.LimitToList = -1
            ctl.AllowValueListEdits = 0
        End If
        
        ''Then the label
        Set ctl = CreateControl(frm.Name, acLabel, , ctlName, , x, IIf(DataEntryMode, comboY, y), totalWidth)
        FilterControlSetCommonProperties ctl, frm, DataEntryMode
        lblName = "lbl" & ctlName
        ctl.Name = lblName
        ctl.Caption = GetFieldCaption(modelFieldRS.fields("VerboseName"), modelFieldRS.fields("ModelField"))
        
        If DataEntryMode Then
            Set ctrlArr = New clsArray
            ctrlArr.Add lblName
            ctrlArr.Add ctlName
            RepositionControlsInRow frm, "2,5", ctrlArr.JoinArr(","), x, InchToTwip(4), InchToTwip(0.2), comboY, True, 0
        End If
        GoTo NextFilter:
    End If
    
    
    ''Double Filter
    If modelFieldRS.fields("FieldTypeID") = dbDouble Then
        
        startX = x
        ctlName = "fltr" & modelFieldRS.fields("ModelField")
        
        ''Render the From
        Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, 0)
        ctl.Name = ctlName & "From"
        FilterControlSetCommonProperties ctl, frm
        
        ctl.Format = "Standard"

        ''Render the To
        Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, 0)
        ctl.Name = ctlName & "To"
        FilterControlSetCommonProperties ctl, frm
        ctl.Format = "Standard"
        
        ''Render the label
        Set ctl = CreateControl(frm.Name, acLabel, , , , 0, 0, 0)
        ctl.Name = "lbl" & ctlName & "To"
        ctl.Caption = "TO"
        ctl.TextAlign = 2
        FilterControlSetCommonProperties ctl, frm
        
        RepositionControlsInRow frm, "10,2,10", ctlName & "From,lbl" & ctlName & "To," & ctlName & "To", startX, totalWidth, 80
        
        ''Then the label
        Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "From", , x, y, totalWidth)
        FilterControlSetCommonProperties ctl, frm
            
        ctl.Caption = GetFieldCaption(modelFieldRS.fields("VerboseName"), modelFieldRS.fields("ModelField"))
        GoTo NextFilter:
    End If
    
    
    If rs.fields("IsMonthYear") Then
        
        startX = x
        ctlName = "fltr" & modelFieldRS.fields("ModelField")
        ''Render the Month
        Set ctl = CreateControl(frm.Name, acComboBox, , , , 0, 0, totalWidth)
        ctl.Name = ctlName & "Month"
        FilterControlSetCommonProperties ctl, frm
        
        ctl.ColumnCount = 2
        ctl.ColumnWidths = "0;1"
        ctl.RowSource = "SELECT MonthID, MonthName FROM tblMonths ORDER BY MonthID"
        ''Render the month label
        Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "Month", , x, y, totalWidth)
        ctl.Name = "lbl" & ctlName & "Month"
        ctl.Caption = "Month"
        FilterControlSetCommonProperties ctl, frm
        
        ''Render the year
        Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, totalWidth)
        ctl.Name = ctlName & "Year"
        FilterControlSetCommonProperties ctl, frm
        
         ''Render the year label
        Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "Year", , x, y, totalWidth)
        ctl.Name = "lbl" & ctlName & "Year"
        ctl.Caption = "Year"
        FilterControlSetCommonProperties ctl, frm
        
        RepositionControlsInRow frm, "6,10,4,8", "lbl" & ctlName & "Month" & "," & _
                         ctlName & "Month," & _
                         "lbl" & ctlName & "Year," & _
                         ctlName & "Year", startX, totalWidth, 80
        
        ''Then the label
        Set ctl = CreateControl(frm.Name, acLabel, , , , x, y, totalWidth)
        FilterControlSetCommonProperties ctl, frm
        
        ctl.Caption = GetFieldCaption(modelFieldRS.fields("VerboseName"), modelFieldRS.fields("ModelField"))
        GoTo NextFilter:
    End If
    
    If rs.fields("IsBetween") Then
        
        startX = x
        ctlName = "fltr" & modelFieldRS.fields("ModelField")
        
        ''Render the "From" Control --> ctlName & "From"
        Set ctl = CreateNamedControl(frm, acTextBox, ctlName & "From")
        FilterControlSetCommonProperties ctl, frm
        ctl.Format = "Short Date"
        ctl.AfterUpdate = "=CopyFromToToDate([Form], " & EscapeString(ctlName) & ")"
        
        ''Render the "To" Control
        Set ctl = CreateNamedControl(frm, acTextBox, ctlName & "To")
        FilterControlSetCommonProperties ctl, frm
        ctl.Format = "Short Date"
        ctl.AfterUpdate = "=CopyFromToIfEarlier([Form], " & EscapeString(ctlName) & ")"
        
        ''Render the Label of the to filter --> "lbl" & ctlName & "To"
        Set ctl = CreateNamedControl(frm, acLabel, "lbl" & ctlName & "To")
        ctl.Caption = "TO"
        ctl.TextAlign = 2
        FilterControlSetCommonProperties ctl, frm
        
        ''Render the label of the filter
        Dim parentName: parentName = ctlName & "From": lblName = "lbl" & parentName
        Set ctl = CreateNamedControl(frm, acLabel, lblName, parentName)
        FilterControlSetCommonProperties ctl, frm
        ctl.Caption = GetFieldCaption(modelFieldRS.fields("VerboseName"), modelFieldRS.fields("ModelField"))
        
        ''Reposition the label first
        RepositionControlsInRow frm, "1", lblName, x, totalWidth
        ''Reposition the controls
        Dim ControlNames As New clsArray
        ControlNames.Add ctlName & "From"
        ControlNames.Add "lbl" & ctlName & "To"
        ControlNames.Add ctlName & "To"
        
        RepositionControlsInRow frm, "10,4,10", ControlNames.JoinArr, x, totalWidth, 40
        
        GoTo NextFilter:
    End If
NextFilter:
       
    
End Sub

Public Function GetProportionTotal(proportionArr As clsArray) As Double

    Dim proportion
    For Each proportion In proportionArr.arr
        GetProportionTotal = GetProportionTotal + CDbl(proportion)
    Next proportion
    
End Function


Public Function CreateDSForm(frm2 As Form, Optional DontOpen As Boolean = False)

    Dim frm As Form, rs As Recordset, rsName, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth
    Dim ctl As control
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus
    Dim QueryName, OnFormCreate, SubformName, UserQueryFields, Timestamp, CreatedBy, PrimaryKey, IsKeyVisible
    
    ''Fetch Variables from the Form
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    SubformName = frm2("SubformName")
    UserQueryFields = frm2("UserQueryFields")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    PrimaryKey = frm2("PrimaryKey")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    ''Generate fields from the table of the model
    GenerateFields frm2

    ''Create the form -> get the TableName, if QueryName is present then override
    ''the TableName with it.
    Set frm = CreateForm
    rsName = GetTableName(Model, VerbosePlural)
    If Not IsNull(QueryName) Then rsName = QueryName
    
    frm.recordSource = rsName
    
    ''Set form Caption
    Dim VerboseCaption: VerboseCaption = ELookup("tblSupplementalModels", "ModelID = " & ModelID, "VerboseCaption")
    Dim frmCaption: frmCaption = GetFieldCaption(VerboseName, Model, VerboseCaption)
    frm.Caption = concat(frmCaption, " Datasheet")
    
    ''Activate the headers and footer but reduce their height to 0
    ''(For summing purposes)
    DoCmd.RunCommand acCmdFormHdrFtr
    frm.Section(acHeader).Height = 0
    frm.Section(acFooter).Height = 0
    
    ''Choose the primary key -> Will automatic based on the ModelName
    ''If PrimaryKey is present then will be overriden
    If IsNull(PrimaryKey) Then PrimaryKey = concat(Model, "ID")
    frm.BeforeUpdate = "=SaveFormData2([Form],""" & Model & """)"
    frm.OnLoad = "=SetDefaultUserID([Form])"
    
    ''Set the Form's Properties
    SetFormProperties 5, frm
    
    ''Open the recordset of table or query based on the logic above
    Dim rsObj As Object, db As DAO.Database, fld As DAO.field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    ''Starting x and y position
    x = 400
    y = 600
    
    CurrentCol = 0: isMemo = False
    
    ''Select the fields -> if plain table as is. If useQueryFields is ticked then use
    ''the query fields
    Dim sqlStr As String
    sqlStr = "SELECT * FROM tblModelFields WHERE FieldOrder <> 0 AND ModelID = " & ModelID
    
    If Not UserQueryFields Then
        sqlStr = sqlStr & " AND FieldSource = " & EscapeString(rsName)
    End If
    
    Set rs = ReturnRecordset(sqlStr & " ORDER BY FieldOrder ASC")
    
    Do Until rs.EOF
        
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"), Not IsNull(rs.fields("ParentModelID")))
        fldWidth = 3000
        Set fld = rsObj.fields(fldName)
        
        If (Not IsKeyVisible And fld.Name = PrimaryKey) Or Not IsNull(rs.fields("ControlSource")) Then
            Select Case fld.Type
                Case dbText, dbDate:
                Case Else:
                    GoTo NextField
            End Select
        End If
        
        ''Skip the field if there is an imageType proprety on qryModelFieldProperties
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("imageType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            GoTo NextField
        End If
        
        Select Case fld.Name
            Case "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
        Dim ControlTypeValue
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeValue = acTextBox
        Else
            ControlTypeValue = fld.Properties("DisplayControl")
        End If
        
        ''Generate the control first before the label
        ''Get the width depending on the Number of columns but make sure that the CurrentCol + the columns will not exceed
        ''the FormColumn
        fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
        
        Set ctl = CreateControl(frm.Name, ControlTypeValue, , "", fld.Name, x, y, fldWidth)
        ctl.Name = fld.Name
        
        ''ctl.ColumnWidth = IIf(IsNull(rs.Fields("ColumnWidth")), ctl.ColumnWidth, ctl.ColumnWidth * rs.Fields("ColumnWidth"))
        
        ''Set control property based on ControlTypeValue
        SetControlPropertiesFromTemplate ctl, frm
        
        If Not IsNull(rs.fields("ColumnWidth")) Then
            ctl.ColumnWidth = 2000 + rs.fields("ColumnWidth")
            ctl.Tag = ctl.Tag & "DontAutoWidth"
            Debug.Print ""
        End If
        
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("alwaysHideOnDatasheet") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            ctl.ColumnHidden = True
            ctl.Tag = ctl.Tag & "alwaysHideOnDatasheet"
        End If
        
        Select Case fld.Type
            Case dbMemo:
                ctl.Height = 900
                isMemo = True
            Case dbDouble:
                ctl.Format = "Standard"
            
        End Select
        
        Select Case fld.Type
        
            Case dbDouble, dbInteger:
                ''Create a control at the footer of the form
                Dim footerCtl As control, footerControlCaption
                Set footerCtl = CreateControl(frm.Name, acTextBox, acFooter, "", , 400, 600, 3000)
                SetControlProperties footerCtl
                footerCtl.Name = concat("Sum", rs.fields("ModelField"))
                footerCtl.ControlSource = "=CdblNz(Sum([" & fld.Name & "]))"
                
                If IsNull(rs.fields("VerboseName")) Then
                    footerControlCaption = AddSpaces(rs.fields("ModelField"))
                Else
                    footerControlCaption = rs.fields("VerboseName")
                End If
                
                footerCtl.Properties("DatasheetCaption") = footerControlCaption
                 
        End Select
        
        ''Also set the DataSheetCaption
        If Not IsNull(rs.fields("VerboseName")) Then
            ctl.Properties("DatasheetCaption") = rs.fields("VerboseName")
        Else
            If Not DoesPropertyExists(fld.Properties, "Caption") Then
                ctl.Properties("DatasheetCaption") = AddSpaces(fld.Name)
            Else
                ctl.Properties("DatasheetCaption") = fld.Properties("Caption")
            End If
        End If
    
'        ''Generate the label just above the control
'        Set ctl = CreateControl(frm.Name, acLabel, , fld.Name, fld.Properties("Caption"), x, y - 300)
'        SetControlPropertiesFromTemplate ctl, frm
'        ctl.Width = fldWidth
        
        CurrentCol = CurrentCol + rs.fields("Columns")
        
        If CurrentCol >= FormColumns Then
            
            If x + 3000 + 400 > maxWidth Then
                maxWidth = x + 3000 + 400
            End If
            
            CurrentCol = 0
            x = 400
            If Not isMemo Then
                y = y + 350
            Else
                isMemo = False
                y = y + 350 + 600
            End If


        Else
        
            x = x + (3200 * rs.fields("Columns"))
            
            
        End If
NextField:
        
        rs.MoveNext
    Loop
    
    ''Dont render the timestamp and created by by default
    If isPresent("qryModelProperties", "Property = ""RenderTimestampCreatedBy"" And ModelID = " & ModelID) Then
        ''Create the Timestamp and CreatedBy field (Hidden Fields)
        Set ctl = CreateControl(frm.Name, acTextBox, , "", "Timestamp", 0, 0, 0)
        ctl.Name = "Timestamp"
        SetControlPropertiesFromTemplate ctl, frm
        ctl.ColumnWidth = 2000
        
        Set ctl = CreateControl(frm.Name, acComboBox, , "", "CreatedBy", 0, 0, 0)
        ctl.Name = "CreatedBy"
        ctl.Properties("DatasheetCaption") = "Created By"
        SetControlPropertiesFromTemplate ctl, frm
        
        frm("Timestamp").Locked = True
        frm("CreatedBy").Locked = True
    End If
    
    frm.Width = (FormColumns * 3000) + (FormColumns * 400) - 200

    ''Attach the form validation
    
    ''Buttons
    frm.Section("Detail").Height = y + 800
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, frm, 5
    End If
    
    ''Override Properties Here
    OverrideProperties ModelID, 5, frm
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    frmName = frm.Name
    
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("dsht", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("dsht", Model, "s")
    End If
    
    If Not IsNull(SubformName) Then
        baseFormName = concat("dsht", SubformName)
    End If
    
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not FrmExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    ''Insert the newly created form to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acForm, customFrmName
    
    DoCmd.Rename customFrmName, acForm, frmName
    
    If Not DontOpen Then DoCmd.OpenForm customFrmName, acFormDS
    
    InsertFormInFormForRights customFrmName, Model
    
End Function


Public Function GenerateFields(frm As Form)

    Dim rsName
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, Timestamp, CreatedBy
    
    ModelID = frm("ModelID")
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
    Timestamp = frm("Timestamp")
    CreatedBy = frm("CreatedBy")

    rsName = GetTableName(Model, VerbosePlural)
    
    If Not IsNull(QueryName) Then
        rsName = QueryName
    End If
    
    Dim rsObj As Object, db As DAO.Database, fld As DAO.field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    Dim fieldArr As New clsArray, valueArr As New clsArray, PrimaryKey
    
    PrimaryKey = concat(Model, "ID")
    
    fieldArr.arr = "ModelID,ModelField,FieldTypeID,FieldOrder,ColumnWidth,FieldSource"
    
    Dim ModelField, i As Integer, maxFieldOrder
    maxFieldOrder = ELookup("tblModelFields", "ModelID = " & ModelID & " AND FieldOrder IS NOT NULL", "FieldOrder", "FieldOrder DESC")
    If maxFieldOrder = "" Then
        i = 1
    Else
        i = CInt(maxFieldOrder) + 1
    End If
    
    
    For Each fld In rsObj.fields
            
        ModelField = fld.Name
        If Not isPresent("tblModelFields", "ModelField = " & EscapeString(ModelField) & _
                                          " And ModelID = " & ModelID) Then
            Select Case ModelField
                Case PrimaryKey, "Timestamp", "CreatedBy", "RecordImportID":
                    ''Empty
                Case Else:
                    Set valueArr = New clsArray
                    valueArr.Add ModelID
                    valueArr.Add EscapeString(fld.Name)
                    valueArr.Add fld.Type
                    valueArr.Add i
                    valueArr.Add "Null"
                    valueArr.Add EscapeString(fld.SourceTable)
                    
                    RunSQL "INSERT INTO tblModelFields (" & fieldArr.JoinArr & ") VALUES (" & valueArr.JoinArr & ")"
                    
                    i = i + 1
            End Select
            
        Else
        
            RunSQL "UPDATE tblModelFields SET FieldSource = " & EscapeString(fld.SourceTable) & " WHERE ModelField = " & EscapeString(ModelField) & _
                                          " And ModelID = " & ModelID
            
        End If
        
    Next fld
    
    If IsFormOpen("frmModels") Then
        Forms("frmModels")("subModelFields").Form.Requery
    End If
    
    'DoCmd.OpenForm "frmModels", , , "ModelID = " & ModelID

End Function

Public Function EnumerateSubformFields(frm As Form)
    
    ''Loop at all the models in which the current model is a ParentModelID
    ''Look at each fields and enumerate all the dbInteger and dbDouble FieldTypeIDs
    Dim ModelID
    
    ModelID = frm("ModelID")
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, sqlStr2, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModelFields"
        .fields = "tblModelFields.*, Model, VerbosePlural"
        .AddFilter "ParentModelID = " & ModelID
        .joins.Add GenerateJoinObj("tblModels", "ModelID")
        Set rs = .Recordset
    End With
    
    Do Until rs.EOF
    
        Dim SubformName
        
        If IsNull(rs("VerboseChildName")) Then
            If IsNull(rs.fields("VerbosePlural")) Then
                SubformName = concat("sub", rs.fields("Model"), "s")
            Else
                SubformName = concat("sub", rs.fields("VerbosePlural"))
            End If
        Else
            SubformName = RemoveSpaces(concat("sub", rs("VerboseChildName")))
        End If
        
        Dim ParentModelID
        ParentModelID = rs.fields("ModelID")
        
        ''SELECT All the tblSubformControls of this ModelID
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = "tblSubformControls"
            .AddFilter "ModelID = " & ModelID
            sqlStr2 = .sql
        End With
        
        ''SELECT all the fields from the ParentModelID's ModelID
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = "tblModelFields"
            .AddFilter "tblModelFields.FieldTypeID In (" & dbInteger & "," & dbDouble & ") AND ModelID = " & ParentModelID
            .fields = "tblModelFields.ModelField As controlName, " & EscapeString(SubformName) & " AS SubformName, " & ModelID & " As ModelID" & _
                      ",AddSpaces([ModelField]) AS ControlCaption, FieldTypeID"
            sqlStr = .sql
        End With
        
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = sqlStr
            .AddFilter "SubformControlID IS NULL"
            .fields = "temp2.*"
            .joins.Add GenerateJoinObj(sqlStr2, "controlName,SubformName", "temp", "controlName,SubformName", "LEFT")
            .SourceAlias = "temp2"
            sqlStr = .sql
        End With
        
        'INSERT STATEMENT
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "INSERT"
            .Source = "tblSubformControls"
            .fields = "controlName, SubformName, ModelID, ControlCaption, FieldTypeID"
            .insertSQL = sqlStr
            .InsertFilterField = "controlName, SubformName, ModelID, ControlCaption, FieldTypeID"
            rowsAffected = .Run
        End With
    
        rs.MoveNext
        
    Loop
    
    If DoesPropertyExists(frm, "subSubformControls") Then
        frm("subSubformControls").Form.Requery
    End If
    
End Function

''Command Name: Seed Table
Public Function SeedTable(frm As Object, Optional ModelID = "", Optional UpdateMode As Boolean = False, Optional recordSize = 50, Optional notify As Boolean = True)

    RunCommandSaveRecord

    If isFalse(ModelID) Then
        ModelID = frm("ModelID")
        If ExitIfTrue(isFalse(ModelID), "ModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblModels WHERE ModelID = " & ModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim Model: Model = rs.fields("Model"): If ExitIfTrue(isFalse(Model), "Model is empty..") Then Exit Function
    
    Dim TableName: TableName = GetTableName(Model, , , True)

    Dim i

    Set rs = ReturnRecordset("SELECT * FROM qryModelFields WHERE ModelID = " & ModelID & " AND NOT MockDataTypeID IS NULL ORDER BY ModelFieldID")
    Dim fldArrItem, fldArr As New clsArray
    Dim mockDataTypeArr As New clsArray
    Dim relatedTablesArr As New clsArray
    Dim fieldTypesArr As New clsArray
    Dim possibleValuesArr As New clsArray
    Dim validationStrArr As New clsArray
    Do Until rs.EOF
        Dim ModelField: ModelField = rs.fields("ModelField"): If ExitIfTrue(isFalse(ModelField), "ModelField is empty..") Then Exit Function
        Dim MockDataType: MockDataType = rs.fields("MockDataType"): If ExitIfTrue(isFalse(MockDataType), "MockDataType is empty..") Then Exit Function
        Dim ParentModelID: ParentModelID = rs.fields("ParentModelID")
        Dim FieldTypeID: FieldTypeID = rs.fields("FieldTypeID"): If ExitIfTrue(isFalse(FieldTypeID), "FieldTypeID is empty..") Then Exit Function
        Dim possibleValues: possibleValues = rs.fields("PossibleValues")
        Dim ValidationString: ValidationString = rs.fields("ValidationString")
        
        If IsNull(ValidationString) Then
            validationStrArr.Add ""
        Else
            validationStrArr.Add ValidationString
        End If
        
        If IsNull(possibleValues) Then
            possibleValuesArr.Add ""
        Else
            possibleValuesArr.Add possibleValues
        End If
        
        fieldTypesArr.Add FieldTypeID
        fldArr.Add ModelField
        mockDataTypeArr.Add MockDataType
        If IsNull(ParentModelID) Then
            relatedTablesArr.Add ""
        Else
            Dim UseAsModel: UseAsModel = ELookup("tblSupplementalModels", "ModelID = " & ParentModelID, "UseAsModel")
            relatedTablesArr.Add GetTableNameFromModelID(IIf(isFalse(UseAsModel), ParentModelID, UseAsModel), True)
        End If
        rs.MoveNext
    Loop
    
    Dim fldValArr As New clsArray
    
    If UpdateMode Then
       Set rs = ReturnRecordset("SELECT * FROM " & TableName)
       recordSize = CountRecordset(rs)
    End If
    
    For i = 0 To recordSize - 1
    
        Set fldValArr = New clsArray
        Dim j As Integer: j = 0
        For Each fldArrItem In fldArr.arr
            MockDataType = mockDataTypeArr.items(j)
            Select Case MockDataType
                Case "Code":
                    fldValArr.Add EscapeString(GetRandomCode(), TableName, fldArr.arr(j))
                Case "Related Field":
                    Dim RelatedTable As String: RelatedTable = relatedTablesArr.arr(j)
                    ValidationString = validationStrArr.arr(j)
                    Dim fieldValue
                    
                    If Not ValidationString Like "*required*" Then
                        Randomize ' Initialize random Number generator
                        Dim randomIndex: randomIndex = GetRandomFromRange(0, 1)
                        If randomIndex = 0 Then
                            fieldValue = Null
                        Else
                            fieldValue = GetRandomID(RelatedTable, GetPrimaryKeyFieldFromTable(RelatedTable))
                        End If
                    Else
                        fieldValue = GetRandomID(RelatedTable, GetPrimaryKeyFieldFromTable(RelatedTable))
                    End If
                    
                    fldValArr.Add EscapeString(fieldValue, TableName, fldArrItem)
                Case "Range":
                    If fieldTypesArr.arr(j) = 8 Then ''A date
                        fldValArr.Add EscapeString(CDate(GetRandomFromRange(CLng(#1/1/2024#), CLng(#12/31/2024#))), TableName, fldArrItem)
                    Else
                        fldValArr.Add EscapeString(GetRandomFromRange(0, 1000, True), TableName, fldArrItem)
                    End If
                Case "Possible Values":
                    Dim possibleValuesSplitted As Variant
                    possibleValuesSplitted = Split(possibleValuesArr.arr(j), ",")
                    ' Randomly choose a value from possibleValues
                    Randomize ' Initialize random Number generator
                    randomIndex = Int((UBound(possibleValuesSplitted) - LBound(possibleValuesSplitted) + 1) * Rnd + LBound(possibleValuesSplitted))
                    fldValArr.Add EscapeString(Trim(possibleValuesSplitted(randomIndex)), TableName, fldArrItem)
                Case Else:
                    fldValArr.Add EscapeString(GetRandomID("tblMockData", replace(MockDataType, " ", "")), TableName, fldArrItem)
            End Select
            j = j + 1
        Next fldArrItem
        
        If UpdateMode Then
            Dim setStatementsArr As New clsArray: Set setStatementsArr = New clsArray
            Dim fldItem, k: k = 0
            rs.Edit
            For Each fldItem In fldArr.arr
                rs.fields(fldItem) = UnescapeValue(fldValArr.arr(k))
                k = k + 1
            Next fldItem
            rs.Update
            rs.MoveNext
        Else
            RunSQL "INSERT INTO " & TableName & " (" & fldArr.JoinArr(",") & ") VALUES (" & fldValArr.JoinArr(",") & ")"
        End If
        
    Next i
    
    If notify Then MsgBox Esc(TableName) & " successfully " & IIf(UpdateMode, "updated", "seeded") & "."
        
    
End Function

''Command Name: Update Seed
Public Function UpdateSeed(frm As Object, Optional ModelID = "")

    RunCommandSaveRecord

    If isFalse(ModelID) Then
        ModelID = frm("ModelID")
        If ExitIfTrue(isFalse(ModelID), "ModelID is empty..") Then Exit Function
    End If

    SeedTable frm, ModelID, True
    
End Function
''Command Name: Create Selector Form
Public Function CreateSelectorForm(frm As Object, Optional ModelID = "")

    RunCommandSaveRecord

    If isFalse(ModelID) Then
        ModelID = frm("ModelID")
        If ExitIfTrue(isFalse(ModelID), "ModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblModels WHERE ModelID = " & ModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    CreateContinuousForm frm, True
    
    Dim ContFormName: ContFormName = "cont" & GetFormSuffix(ModelID)
    
    DoCmd.OpenForm ContFormName, acDesign
    Dim contForm As Form: Set contForm = Forms(ContFormName)
    CreateContinuousFormButton contForm, GetStandardControlHeight(contForm), "=" & Esc("Select"), "txtSelect", "cmdSelect"
    contForm.Section(acDetail).Height = 0
    contForm.AllowAdditions = False
    contForm.AllowEdits = False
    contForm.AllowDeletions = False
    contForm.NavigationButtons = False
    
    contForm("txtSelect").InSelection = True
    contForm("cmdSelect").InSelection = True
    
    DoCmd.RunCommand acCmdControlPaddingNone
    DoCmd.RunCommand acCmdRemoveFromLayout
    
    AnchorControls contForm, ModelID
    
    Dim ctl As control
    For Each ctl In contForm.controls
        If ctl.ControlType = acEmptyCell Then
            DeleteControl ContFormName, ctl.Name
        End If
    Next ctl
    
    contForm("txtSelect").Left = 0
    contForm("cmdSelect").Left = 0
    
    contForm("txtSelect").Left = GetMaxX(contForm)
    contForm("txtSelect").Top = 0
    contForm("cmdSelect").Left = contForm("txtSelect").Left
    contForm("cmdSelect").Top = 0
    
    contForm("txtSelect").Width = InchToTwip(1)
    contForm("cmdSelect").Width = InchToTwip(1)
    contForm("cmdSelect").HorizontalAnchor = acHorizontalAnchorRight
    contForm("txtSelect").HorizontalAnchor = acHorizontalAnchorRight
    
    contForm.Section(acDetail).Height = 0
    contForm.Width = 0
    
    DoCmd.Close acForm, ContFormName, acSaveYes
    
    Dim VerbosePluralCaption: VerbosePluralCaption = rs.fields("VerbosePluralCaption")
    Dim mainForm As Form: Set mainForm = CreateForm
    SetFormProperties 4, mainForm
    mainForm.Caption = VerbosePluralCaption
    mainForm.MinMaxButtons = 0
    
    Set ctl = CreateControl(mainForm.Name, acSubform, , , , 0, 0, mainForm.Width, mainForm.Section(acDetail).Height)
    ctl.SourceObject = ContFormName
    ctl.Name = "subform"
    
    Dim origMainFormName: origMainFormName = mainForm.Name
    Dim MainFormName: MainFormName = "main" & GetFormSuffix(ModelID)
    DoCmd.Close acForm, origMainFormName, acSaveYes
    DoCmd.Rename MainFormName, acForm, origMainFormName
    
    DoCmd.OpenForm MainFormName, acDesign
    Set frm = Forms(MainFormName)
    
    Dim OnFormCreate: OnFormCreate = rs.fields("OnFormCreate")
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, frm, 9
    End If
    
End Function

Public Function SelectRecordFromSelector(frm As Object, TargetForm, RecordIDName, DropdownName, AfterUpdateCallback)

    If Not IsFormOpen(TargetForm) Then Exit Function
    
    Dim frm2 As Form: Set frm2 = Forms(TargetForm)
    
    Dim RecordID: RecordID = frm(RecordIDName)
    
    If isFalse(RecordID) Then Exit Function
    
    frm2(DropdownName) = RecordID
    
    Run AfterUpdateCallback, frm2
    
    DoCmd.Close acForm, frm.parent.Name, acSaveNo
    
End Function
''Command Name: Create Single Record Report
Public Function CreateSingleRecordReport(frm As Object, Optional ModelID = "")

    RunCommandSaveRecord

    If isFalse(ModelID) Then
        ModelID = frm("ModelID")
        If ExitIfTrue(isFalse(ModelID), "ModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblModels WHERE ModelID = " & ModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    CreateDEForm frm, True
    
End Function

''Command Name: Create Continuous Record Report
Public Function CreateContinuousRecordReport(frm As Object, Optional ModelID = "")

    RunCommandSaveRecord

    If isFalse(ModelID) Then
        ModelID = frm("ModelID")
        If ExitIfTrue(isFalse(ModelID), "ModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblModels WHERE ModelID = " & ModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    CreateContinuousForm frm, False, True
    
End Function


