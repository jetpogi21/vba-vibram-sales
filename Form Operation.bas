Attribute VB_Name = "Form Operation"
Option Compare Database
Option Explicit


Public Function CloseThisForm(frm As Form)

    DoCmd.Close acForm, frm.Name, acSaveNo
    
End Function

Public Sub CloseAllForms(Optional Exceptions = "")
    
    Dim ExceptionsArr As New clsArray
    If Not isFalse(Exceptions) Then
        ExceptionsArr.arr = Exceptions
    End If
    
    Dim frm As Form
    For Each frm In Forms
        If Not ExceptionsArr.InArray(frm.Name) Then
            DoCmd.Close acForm, frm.Name, acSaveNo
        End If
    Next frm
    
End Sub

Public Sub Test_CloseAllForms()

    DoCmd.OpenForm "frmCustomDashboard"
    DoCmd.OpenForm "mainModels"
    
    CloseAllForms "frmCustomDashboard"
    
    Dim frm As Form: Set frm = GetForm("frmCustomDashboard")
    If frm Is Nothing Then
        ShowError Esc("frmCustomDashboard") & " should be loaded."
        Exit Sub
    End If
    ''Check if frmCustomDashboard is still loaded
    
    ''Check if mainModels is not loaded
    Set frm = GetForm("mainModels")
    If Not frm Is Nothing Then
        ShowError Esc("mainModels") & " should not be loaded."
        Exit Sub
    End If
    
End Sub

Public Function FormReports_ToggleSubformVisiblity(frm As Object)
    
    Dim ctl As control
    Dim isReport As Boolean: isReport = IsObjectAReport(frm)
    For Each ctl In frm.controls
        If ctl.ControlType = acSubform Then
            Dim SubformName: SubformName = ctl.Name
            If isReport Then
                frm(SubformName).Visible = GetSubReportHasData(frm, SubformName)
                Dim bannerCtl: bannerCtl = "banner_" & SubformName
                If DoesPropertyExists(frm, bannerCtl) Then
                    frm(bannerCtl).Visible = Not GetSubReportHasData(frm, SubformName)
                End If
            Else
                Dim rs As Recordset: Set rs = GetSubformRecordsetClone(frm, SubformName)
                frm(SubformName).Visible = CountRecordset(rs) > 0
            End If
            
        End If
    Next ctl
End Function

Public Function GetSubReportHasData(obj As Object, SubformName) As Boolean

On Error GoTo ErrHandler:
    GetSubReportHasData = obj(SubformName).Report.HasData
    Exit Function
ErrHandler:
    If Err.Number = 2455 Then
        Exit Function
    End If
    
End Function

Public Function GetSubformRecordsetClone(obj As Object, SubformName) As Recordset

On Error GoTo ErrHandler:
    Dim rs As Recordset: Set rs = obj(SubformName).Form.RecordsetClone
    Exit Function
ErrHandler:
    If Err.Number = 2467 Then
        Set rs = obj(SubformName).Report.RecordsetClone
        Exit Function
    End If
    
End Function


Public Function GetSubformValue(SubformControl, Optional ValueOnEmpty = Null) As Variant
    On Error GoTo ErrHandler
    GetSubformValue = IIf(IsError(SubformControl), ValueOnEmpty, SubformControl)
    Exit Function
ErrHandler:
  If Err.Number = 2427 Then
    GetSubformValue = ValueOnEmpty
  End If
End Function

Public Sub RequeryForm(frmName, Optional SubformName = "", Optional ControlName = "")
    
    If IsFormOpen(frmName) Then
        Dim frm As Form: Set frm = Forms(frmName)
        
        If Not isFalse(SubformName) Then
            frm(SubformName).Form.Requery
            Exit Sub
        End If
        
        If Not isFalse(ControlName) Then
            frm(ControlName).Requery
            Exit Sub
        End If
        
        frm.Requery
    End If
    
End Sub

Public Function SplitWordsByUppercase(Text) As String
    Dim result As String
    Dim i As Integer
    Dim currentChar As String
    
    result = ""
    
    For i = 1 To Len(Text)
        currentChar = Mid(Text, i, 1)
        
        If Asc(currentChar) >= 65 And Asc(currentChar) <= 90 And i > 1 Then
            result = result & " "
        End If
        
        result = result & currentChar
    Next i
    
    SplitWordsByUppercase = result
End Function

Public Function GetVerboseName(Name)
    
    GetVerboseName = Name
    If Name Like "*ID" Then
        GetVerboseName = Left(GetVerboseName, Len(GetVerboseName) - 2)
    End If
    
    GetVerboseName = SplitWordsByUppercase(GetVerboseName)
    
End Function

Public Sub ReplaceComboBoxProperties(frmName, ControlName, toReplace, replaceWith)

    Dim frm As Form: Set frm = Forms(frmName)
    Dim ctl As ComboBox: Set ctl = frm(ControlName)
    
    ''ControlSource
    ctl.ControlSource = replace(ctl.ControlSource, toReplace, replaceWith)
    ''RowSource
    ctl.RowSource = replace(ctl.RowSource, toReplace, replaceWith)
    ''controlName
    ctl.Name = replace(ctl.Name, toReplace, replaceWith)
    ''DatasheetCaption
    ctl.Properties("DatasheetCaption") = GetVerboseName(replaceWith)
    
End Sub

Public Function AlwaysOpenSwitchboards(Optional frmName = "frmCustomDashboard")
    
    If Not GetIsAdmin And frmName = "frmCustomDashboard" Then
        frmName = "frmNonAdminDashboard"
    End If
    
    If Not IsFormOpen(frmName) Then
        
        DoCmd.OpenForm frmName
    End If
    
End Function

Public Function AlwaysCloseSwitchboards()
    
    If IsFormOpen("frmCustomDashboard") Then
        DoCmd.Close acForm, "frmCustomDashboard", acSaveNo
    End If
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT ParentMenu FROM tblMainMenus WHERE NOT ParentMenu IS NULL GROUP BY ParentMenu")
    
    Do Until rs.EOF
        Dim ParentMenu: ParentMenu = rs.fields("ParentMenu")
        Dim SwitchboardName: SwitchboardName = "frm" & ParentMenu & "Dashboard"
        If IsFormOpen(SwitchboardName) Then
            DoCmd.Close acForm, SwitchboardName, acSaveNo
        End If
        rs.MoveNext
    Loop

End Function

Public Function FindFirst(frm As Object, condition As String) As Boolean
    
    Dim rs As Recordset: Set rs = frm.RecordsetClone
    rs.FindFirst condition
    If Not rs.NoMatch Then
        frm.Bookmark = rs.Bookmark
        FindFirst = True
        Exit Function
    End If
    
End Function

Private Function GetIDValue(tblName, idField, fieldToMatch, valueToMatch)
    
    GetIDValue = ELookup(tblName, fieldToMatch & " = " & EscapeString(valueToMatch), idField)
    
End Function

Public Function SetDefaultValueFromTable(frm As Object, tblName, idField, fieldToMatch, valueToMatch)

    Dim IDValue: IDValue = GetIDValue(tblName, idField, fieldToMatch, valueToMatch)
    If Not isFalse(IDValue) Then frm(idField).DefaultValue = IDValue
    
End Function

Public Function StrictSave(frm As Object, Model, Operation)
    
    frm("isSaved").Caption = -1
    Save2 frm, "Contact", 1
    
End Function

Public Function SaveFormData(frm As Object, tblName As String, PrimaryKey As String, Optional validationSuccessCB As String) As Boolean

    If areDataValid(frm, tblName) Then
    
        If validationSuccessCB <> "" Then
            Run validationSuccessCB, frm
        End If
        
        If Not frm.NewRecord Then
            UpdateFormData frm, tblName, PrimaryKey
        End If
        
    End If
    
End Function

Public Function GetOldValue(tblName, fieldName, PrimaryKey, RecordID As Variant) As Variant
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM " & tblName & " WHERE " & PrimaryKey & " = " & RecordID)
    
    GetOldValue = rs(fieldName)
    
End Function

Private Function UpdateFormData(frm As Object, tblName As String, PrimaryKey As String)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblFormFields WHERE TableName = """ & tblName & """")
    
    Dim ctl As control, RecordID As Variant
    
    Dim fieldName As String, FieldTypeID As Integer, currentValue, oldValue
    Dim updateStatement() As String, i As Integer
    
    Do Until rs.EOF
        fieldName = rs.fields("FieldName")
        FieldTypeID = rs.fields("FieldTypeID")
        If ControlExists(fieldName, frm) Then
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
        rs.MoveNext
    Loop
    
End Function

Private Function InsertFormData(frm As Object, tblName As String)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblFormFields WHERE TableName = """ & tblName & """")
    
    Dim ctl As control
    
    Dim fieldName As String, FieldTypeID As Integer
    Dim fields() As String, fieldValues() As String, i As Integer
    
    Do Until rs.EOF
        fieldName = rs.fields("FieldName")
        FieldTypeID = rs.fields("FieldTypeID")
        If ControlExists(fieldName, frm) Then
            ReDim Preserve fields(i)
            ReDim Preserve fieldValues(i)
            
            fields(i) = fieldName
            fieldValues(i) = ReturnStringBasedOnType(frm(fieldName), FieldTypeID)
            
            i = i + 1
            
        End If
        rs.MoveNext
    Loop
    
    InsertAndLog tblName, fields(), fieldValues()
    
End Function

Public Function CancelEdit(frm As Object, Optional isChild As Boolean = False)
    frm.Undo
    If isChild Then
        DoCmd.Close acForm, frm.parent.Form.Name, acSaveNo
    Else
        DoCmd.Close acForm, frm.Name, acSaveNo
    End If
End Function

Public Function Save(frm As Object, tblName As String, Operation As Integer, Optional isChild As Boolean = False)
    
    ''Operation: 0 is save and new, 1 is save and close
    If areDataValid(frm, tblName) Then
    
'        If validationSuccessCB <> "" Then
'            Run validationSuccessCB, frm
'        End If
        
        Select Case Operation
            Case 0:
                DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
            Case 1:
                If isChild Then
                    DoCmd.Close acForm, frm.parent.Form.Name, acSaveNo
                Else
                    DoCmd.Close acForm, frm.Name, acSaveNo
                End If
        End Select
    End If
    
End Function

Public Function IsFormOpen(frmName) As Boolean
    On Error GoTo Err_Handler:
    
    IsFormOpen = CurrentProject.AllForms(frmName).IsLoaded
    Exit Function
Err_Handler:
    
    IsFormOpen = False
End Function

Public Function IsReportOpen(rptName) As Boolean
    On Error GoTo Err_Handler:
    
    IsReportOpen = CurrentProject.AllReports(rptName).IsLoaded
    Exit Function
Err_Handler:
    
    IsReportOpen = False
End Function

Public Function Save2(frm As Object, Model As String, Operation As Integer, Optional isChild As Boolean = False)

'    Dim BeforeUpdate
'    BeforeUpdate = frm.BeforeUpdate
'
'    frm.BeforeUpdate = ""
    
    ''Operation: 0 is save and new, 1 is save and close
    If areDataValid2(frm, Model) Then
    
'        If validationSuccessCB <> "" Then
'            Run validationSuccessCB, frm
'        End If
        
        Select Case Operation
            Case 0:
                DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
                'frm.BeforeUpdate = BeforeUpdate
            Case 1:
                If isChild Then
                    DoCmd.Close acForm, frm.parent.Form.Name, acSaveYes
                Else
                    DoCmd.Close acForm, frm.Name, acSaveYes
                End If
            Case Else:
            
                RunCommandSaveRecord
                'frm.BeforeUpdate = BeforeUpdate
                
        End Select
        
    Else
    
        'frm.BeforeUpdate = BeforeUpdate
        
    End If
    
End Function

Public Function DefaultFormLoad(frm As Object, PrimaryKey, Optional AutoWidth As Boolean = True)
    
    SetDefaultUserID frm
    
    '''Also hide the subforms if this is a main form
    Dim ctl As control, subCtl As control, LinkChildFields
    
    For Each ctl In frm.controls
        If ctl.ControlType = acSubform Then
            
            If Not ctl.SourceObject Like "Report.*" Then
                ctl.Form.DatasheetAlternateBackColor = RGB(254, 254, 254)
                ''Hide the related field
                LinkChildFields = ctl.LinkChildFields
                If frm(ctl.Name).Form.DefaultView = 2 Then
                    For Each subCtl In frm(ctl.Name).Form.controls
                        
                        If AutoWidth And Not subCtl.Tag Like "*DontAutoWidth*" Then
                            subCtl.ColumnWidth = -2
                        End If
                        
                        subCtl.ColumnHidden = subCtl.Name = LinkChildFields
                        
                        Select Case subCtl.Name
                            Case "Timestamp", "CreatedBy", "VerboseName", "Model":
                            Case Else:
                                SetColumnHidden subCtl, frm
                        End Select
                        
                        If subCtl.Tag Like "*alwaysHideOnDatasheet*" Then
                            subCtl.ColumnHidden = True
                        End If
                        
                    Next subCtl
                End If
            End If
            
        End If
    Next ctl
    
    ''TranslateToArabic frm
   
End Function

Private Sub SetColumnHidden(subCtl As control, frm As Form)
    
    On Error GoTo ErrHandler:
        
        If Not subCtl.Tag Like "*DontAutoWidth*" Then
            subCtl.ColumnHidden = DoesPropertyExists(frm, subCtl.Name)
        End If
        Exit Sub
    
ErrHandler:
    
    If Err.Number = 2101 Then Exit Sub
    If Err.Number = 2427 Then Exit Sub
    ShowError "Error # " & Err.Number & vbCrLf & Err.description

End Sub

Public Function DefaultMainFormLoad(frm As Form)
    
    '''Reveal all the hidden fields from the subform
    Dim ctl As control
    
    If Not frm.subform.SourceObject Like "Report.*" Then
    
        frm.subform.Form.DatasheetAlternateBackColor = RGB(254, 254, 254)
        
        For Each ctl In frm.subform.Form.controls
            
            If Not ctl.Tag Like "*DontAutoWidth*" Then
                ctl.ColumnWidth = -2
            End If
            
            ''Hide the related field
            If ctl.ColumnHidden = True Then
            
                ctl.ColumnHidden = False
                
            End If
            
            If ctl.Tag Like "*alwaysHideOnDatasheet*" Then
                ctl.ColumnHidden = True
            End If
            
        Next ctl
   End If
   
   ''TranslateToArabic frm
   
End Function

Public Function DefaultReportLoad(frm As Object)

    TranslateToArabic frm
   
End Function

Public Function GetSubformOrReport(ctl As control) As Object
    
    Dim subfrm As Object
    Set subfrm = GetSubformForm(ctl)
    If subfrm Is Nothing Then
        Set subfrm = GetSubformReport(ctl)
    End If
    
    If Not subfrm Is Nothing Then
        Set GetSubformOrReport = subfrm
    End If
    
End Function

Public Function GetSubformForm(ctl As control) As Object

On Error GoTo ErrHandler:

    Dim subfrm As Object
    Set subfrm = ctl.Form
    Set GetSubformForm = subfrm
    Exit Function
    
ErrHandler:
If Err.Number = 2467 Then
    Exit Function
End If

End Function

Public Function GetSubformReport(ctl As control) As Object

On Error GoTo ErrHandler:

    Dim subfrm As Object
    Set subfrm = ctl.Report
    Set GetSubformReport = subfrm
    Exit Function
    
ErrHandler:
If Err.Number = 2467 Then
    Exit Function
End If

End Function


Public Function TranslateToArabic(frm As Object)

    If isFalse(g_Language) Then Exit Function
    
    If g_Language <> "Arabic" Then Exit Function

    Dim ctl As control, subCtl As control, subfrm As Object
    
    For Each ctl In frm.controls
        Dim Translation, Caption, ControlSource
        Select Case ctl.ControlType
            Case acCommandButton, acLabel:
                Caption = ctl.Caption
                Translation = GetTranslation(Caption)
                
                If isFalse(Translation) Then GoTo GoToNextCtl
                    
                ctl.Caption = Translation
               
            Case acSubform:
            
On Error GoTo ErrHandler:
                Set subfrm = GetSubformOrReport(ctl)
                
                If subfrm Is Nothing Then GoTo GoToNextCtl
                
                For Each subCtl In subfrm.controls
                    
                    If IsObjectAReport(subfrm) Then
                        TranslateToArabic subfrm
                    Else
                        Caption = GetDatasheetCaption(subCtl)
                        
                        If isFalse(Caption) Then GoTo GoToNextCtl
                    
                        Translation = GetTranslation(Caption)
                        
                        If isFalse(Translation) Then GoTo GoToNextCtl
                            
                        SetDatasheetCaption2 subCtl, Translation
                    End If
                Next subCtl
        End Select
        
        If frm.Name <> "frmCustomDashboard" And frm.Name <> "frmNonAdminDashboard" Then GoTo GoToNextCtl
            
        If Not ctl.Name Like "txt*" Then GoTo GoToNextCtl
           
        ControlSource = ctl.ControlSource
    
        If isFalse(ControlSource) Then GoTo GoToNextCtl
        
        ControlSource = replace(ControlSource, """", "")
        Caption = replace(ControlSource, "=", "")
        
        Translation = GetTranslation(Caption)
                
        If isFalse(Translation) Then GoTo GoToNextCtl
            
        ctl.ControlSource = "=" & Esc(Translation)
        
GoToNextCtl:
    Next ctl
    
    Exit Function
    
ErrHandler:

    If Err.Number = 2467 Then
        GoTo GoToNextCtl
    End If

End Function

Public Function OpenFormFromRecordReport(rpt As Report, fieldName, frmName)
    
    Dim fieldValue
    fieldValue = rpt(fieldName)
    
    If IsNull(fieldValue) Then Exit Function
    
    DoCmd.OpenForm frmName, , , fieldName & "=" & fieldValue
   
End Function


Public Function OpenFormFromRecord(frm As Object, fieldName, frmName)
    
    Dim fieldValue
    fieldValue = frm(fieldName)
    
    If IsNull(fieldValue) Then Exit Function
    
    DoCmd.OpenForm frmName, , , fieldName & "=" & fieldValue
   
End Function

Public Function OpenFormByWhereClause(frmName, whereClause)
    
    If IsNull(whereClause) Then Exit Function
    
    DoCmd.OpenForm frmName, , , whereClause
   
End Function

Public Function CustomMainFormLoad(frm As Form)

    DefaultMainFormLoad frm
    
    ''Disable base on their rights
    CheckUserRights frm

End Function

Public Function CustomFormLoad(frm As Object, ModelID)
    
    DefaultFormLoad frm, ModelID
    CheckUserRights frm
    
End Function

Public Function CustomDatasheetFormLoad(frm As Form)

    SetDefaultUserID frm
    
    ''Disable base on their rights
    CheckUserRights frm

End Function

Private Function CheckUserRights(frm As Form)
    
    ''TABLE: tblUserRights Fields: UserRightID|User|ModelName|CanView|CanEdit|CanAdd|CanDelete|Timestamp|CreatedBy|RecordImportID
    ''TABLE: tblFormForRights Fields: FormForRightsID|ModelName|FormName|FormType|Timestamp|CreatedBy|RecordImportID|ModelCaption
    
    Dim frmName, frmType, ModelName
    frmName = frm.Name
    frmType = GetFormType(frm) ''DataEntry,DataSheet,MainForm
    ModelName = ELookup("tblFormForRights", "FormName = '" & frmName & "'", "ModelName")
    
    ResetFormToDefault frm, frmType
    
    If isPresent("tblUsers", "UserID = " & g_userID & " AND isAdmin") Then Exit Function
    
    ''Check if can be added
    Dim CanAdd, CanEdit, CanDelete
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblUserRights WHERE User = " & g_userID & " AND ModelName = '" & ModelName & "'")
    
    If rs.EOF Then Exit Function
 
    CanAdd = rs.fields("canEdit")
    CanEdit = rs.fields("canEdit")
    CanDelete = rs.fields("canDelete")
    
    If Not CanAdd Then handleCantAdd frm, frmType
    If Not CanEdit Then handleCantEdit frm, frmType
    If Not CanDelete Then handleCantDelete frm, frmType
    
End Function

Private Function handleCantAdd(frm As Object, frmType)

    If frmType = "MainForm" Then
    
        ''Reset to Default
        frm("cmdAdd").Enabled = False
        
    ElseIf frmType = "DataSheet" Then
        
        frm.AllowAdditions = False
        
    ElseIf frmType = "DataEntry" Then
        
        frm("cmdNew").Enabled = False
        frm("cmdSaveClose").Enabled = True
        
    End If
    
End Function

Private Function handleCantEdit(frm As Object, frmType)

    If frmType = "MainForm" Then
    
        ''Reset to Default
        frm("cmdView").Caption = "View"
        
    ElseIf frmType = "DataSheet" Then
        
        frm.AllowEdits = False
        
    ElseIf frmType = "DataEntry" Then
        
        frm("cmdSaveClose").Enabled = False

    End If
    
End Function

Private Function handleCantDelete(frm As Object, frmType)

    If frmType = "MainForm" Then
    
        ''Reset to Default
        frm("cmdDelete").Enabled = False
        
    ElseIf frmType = "DataSheet" Then
        
        frm.AllowDeletions = False
    
    ElseIf frmType = "DataEntry" Then
        
        frm("cmdDelete").Enabled = False

    End If
    
End Function

Private Function ResetFormToDefault(frm As Object, frmType)
    
    On Error Resume Next
    If frmType = "MainForm" Then
    
        ''Reset to Default
        frm("cmdAdd").Enabled = True
        frm("cmdView").Enabled = True
        frm("cmdDelete").Enabled = True
        
    ElseIf frmType = "DataSheet" Then
        
        frm.AllowAdditions = True
        frm.AllowEdits = True
        frm.AllowDeletions = True
        
    ElseIf frmType = "DataEntry" Then
        
        frm("cmdCancel").Enabled = True
        frm("cmdNew").Enabled = True
        frm("cmdSaveClose").Enabled = True
        frm("cmdDelete").Enabled = True
        
    End If

End Function

Private Function GetFormType(frm As Form) As String
    
    Dim frmName
    frmName = frm.Name
    
    GetFormType = "DataEntry"
    
    If frmName Like "main*" Then
    
        GetFormType = "MainForm"
        Exit Function
        
    ElseIf frmName Like "dsht*" Then
            
        GetFormType = "DataSheet"
        Exit Function
            
    End If
    
End Function

Public Function SetDefaultUserID(frm As Form)

    If isFalse(g_userID) Then
        LogIn
    End If
    
    If DoesPropertyExists(frm, "CreatedBy") Then
        On Error Resume Next
        frm("CreatedBy").DefaultValue = "=" & g_userID
    End If
    
End Function

Public Function SetFocusOnForm(frm As Object, ctlName)
    On Error Resume Next
    If ctlName <> "" Then frm(ctlName).SetFocus
End Function

Public Function DeleteRecord(frm As Object, pkName As String, tblName As String, Optional SubformName As String = "")
        
    Dim frm2 As Form
    If SubformName = "" Then Set frm2 = frm Else Set frm2 = frm(SubformName).Form
    
    Dim IsRowForm: IsRowForm = frm2.Name Like "cont*" Or frm2.Name Like "dsht*"
    
    If ExitIfTrue(frm2.NewRecord And Not IsRowForm, "You can't delete an unsaved new record.") Then Exit Function
    
    If IsNull(frm2(pkName)) Then Exit Function
    
    If frm2.NewRecord And IsRowForm Then
        frm2.Undo
        Exit Function
    End If
    
    If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbYes Then
    
        ''frmType 0 => DataEntry | 1 => Datasheet
        Dim RecordID: RecordID = frm2(pkName)
        Dim pkType: pkType = GetFieldType(tblName, pkName)
        Dim pkValue: pkValue = RecordID
        If pkType = 10 Then pkValue = EscapeString(pkValue)
        
        RunSQL "DELETE FROM " & tblName & " WHERE " & pkName & " = " & pkValue
        
        Insert_Delete_Log tblName, "DELETE", RecordID
        
        frm.OnClose = "=RequeryOnClose('" & tblName & "',True)"
         
        If SubformName = "" Then frm.Requery Else frm(SubformName).Requery
        
    End If
    
End Function

Private Function GetFieldType(tblName, fldName)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT " & fldName & " FROM " & tblName)
    GetFieldType = rs.fields(fldName).Type
    
End Function



Public Function OpenFormFromMain(frmName As String, Optional SubformName As String, Optional PrimaryKey As String, Optional frm As Form, Optional DefaultField As String)

    If DefaultField <> "" Then
        If Not DoesPropertyExists(frm, DefaultField) Then
            ShowError "The parent form is empty.."
            Exit Function
        ElseIf IsNull(frm(DefaultField)) Then
            ShowError "The parent form is empty.."
            Exit Function
        End If
    End If
    
On Error GoTo Err_Handler:
    If PrimaryKey = "" Then
        
        ''2501
        DoCmd.OpenForm frmName, , , , acFormAdd
        If DefaultField <> "" Then
            If DoesPropertyExists(Forms(frmName), DefaultField) Then
                Forms(frmName)(DefaultField).DefaultValue = frm(DefaultField)
            End If
        End If
        
    Else
        Dim pkVal: pkVal = frm(SubformName).Form(PrimaryKey)
        Dim pkType: pkType = frm(SubformName).Form.RecordsetClone.fields(PrimaryKey).Type
        
        If IsNull(pkVal) Then
            ShowError "Please select a record from the list.."
            Exit Function
        End If
        
        If pkType <> 10 Then
            DoCmd.OpenForm frmName, , , PrimaryKey & " = " & pkVal
        Else
            DoCmd.OpenForm frmName, , , PrimaryKey & " = " & EscapeString(pkVal)
        End If
        
    End If
    
    Exit Function
    
Err_Handler:

    If Err.Number = 2501 Then
        Exit Function
    Else
        MsgBox Err.description
    End If
    
End Function

Public Function RequeryOnClose(tblName As String, Optional shouldRequery As Boolean)
    
    If shouldRequery Then
        On Error Resume Next
        Dim requeryForms As Variant, requeryFormArray() As String, RequeryForm As Variant, PrimaryKey
        Dim rs As Recordset, frm As Form, rsClone As Recordset
        Set rs = ReturnRecordset("SELECT * FROM tblTables WHERE TableName = '" & tblName & "'")
        requeryForms = rs.fields("RequeryOnClose")
        PrimaryKey = rs.fields("PrimaryKey")
    
        requeryFormArray = Split(requeryForms)
        For Each RequeryForm In requeryFormArray
            Eval "Forms!" & RequeryForm & ".Requery"
            Set frm = ReturnFormObject(RequeryForm)
            'ReturnMainForm(requeryForm).SetFocus
            Set rsClone = frm.RecordsetClone
            rsClone.FindFirst PrimaryKey & " = " & frm(PrimaryKey)
            If Not rsClone.NoMatch Then
                frm.Bookmark = rsClone.Bookmark
            End If
        Next RequeryForm
    End If
    
    
End Function

Public Function ReturnFormObject(RequeryForm As Variant) As Form

    Dim formParts() As String
    Dim formPart As Variant
    
    formParts = Split(RequeryForm, ".")
    Dim obj As Object
    Set obj = Forms
    
    For Each formPart In formParts
        Set obj = obj(formPart)
    Next formPart
    
    Set ReturnFormObject = obj.Form
    
End Function

Public Function ReturnMainForm(RequeryForm As Variant) As Form

    Dim formParts() As String
    Dim formPart As Variant
    
    formParts = Split(RequeryForm, ".")
    Dim obj As Object
    Set obj = Forms
    
    Set ReturnMainForm = obj(formParts(0))
    
End Function

Public Function SelectSubformRecords(frm As Object, Optional mode As Boolean = False)
    
    Dim tblName As String, sqlObj As New clsSQL
    tblName = frm.subform.Form.recordSource
    
    'UPDATE STATEMENT
    With sqlObj
        .SQLType = "UPDATE"
        .Source = tblName
        .SetStatement = "Selected = " & mode
        .Run
    End With
    
    frm.subform.Requery
    
End Function

'Public Function RefreshSubformData(frm As Object, MainFormUtilityID)
'
'    Dim sqlObj As New clsSQL
'    Dim rs As Recordset
'
'    ''FETCH the variables from MainFormUtilities
'    With sqlObj
'        .Source = "tblMainFormUtilities"
'        .AddFilter "MainFormUtilityID = " & MainFormUtilityID
'        Set rs = .Recordset
'    End With
'
'    Dim QueryName, TempTableName, IgnoreFields, AdditionalFields
'    QueryName = rs.Fields("QueryName")
'    TempTableName = rs.Fields("TempTableName")
'    IgnoreFields = rs.Fields("IgnoreFields")
'    AdditionalFields = rs.Fields("AdditionalFields")
'
'    ''Delete the content of the subform
'    ''DELETE STATEMENT
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .SQLType = "DELETE"
'        .Source = TempTableName
'        .Run
'    End With
'
'    ''Insert the query to the content subform
'    ''SELECT STATEMENT
'    Dim fieldNames, sqlStr
'    fieldNames = GenerateFieldNamesString(TempTableName, IgnoreFields) & AdditionalFields
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .Source = QueryName
'        .Fields = fieldNames
'        sqlStr = .SQL
'    End With
'
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .SQLType = "INSERT"
'        .Source = TempTableName
'        .Fields = fieldNames
'        .InsertSQL = sqlStr
'        .Run
'    End With
'
'    frm.subform.Form.Requery
'
'End Function

Public Function DoOpenForm(frmName, Optional whereCondition, Optional addNew As Boolean = True)
    
    If isFalse(whereCondition) Then
         
         If addNew Then
            DoCmd.OpenForm frmName, , , , acFormAdd
        Else
            DoCmd.OpenForm frmName
        End If
    
    Else
         
         DoCmd.OpenForm frmName, , , whereCondition
    
    End If
    
    ''If frmName Like "main*" And Environ("ComputerName") <> "DESKTOP-3G3V8GO" Then
'    If frmName Like "main*" Then
'        Dim frm As Form: Set frm = Forms(frmName)
'        If isPresent("tblMainMenus", "ParentMenu = ""Setup"" AND FormName = " & Esc(frmName)) Then
'            frm.OnClose = "=AlwaysOpenSwitchboards(" & Esc("frmSetupDashboard") & ")"
'        ElseIf isPresent("tblMainMenus", "ParentMenu = ""Report"" AND FormName = " & Esc(frmName)) Then
'            frm.OnClose = "=AlwaysOpenSwitchboards(" & Esc("frmReportDashboard") & ")"
'        Else
'            frm.OnClose = "=AlwaysOpenSwitchboards()"
'        End If
'        AlwaysCloseSwitchboards
'    End If
   
End Function





