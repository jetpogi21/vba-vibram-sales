Attribute VB_Name = "ControlCreationHelper Mod"
Option Compare Database
Option Explicit

Public Function ControlCreationHelperCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Creation of the from and to textbox control start here   ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RenderFromToTextbox(frm As Object, ControlCreationHelperID)
    
    ''This Sub is Called from CreateCustomControl
    ''TABLE: tblControlCreationHelper Fields: ControlCreationHelperID|CustomControlTypeID|Model|PrimaryKey
    ''MainField|FieldToUse|Direction|Timestamp|CreatedBy|RecordImportID|Width|ParentModel|IsDynamicList|FieldCaption
    ''MiddleTable|PossibleValues|NoneValue|PairSize|AsInline
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblControlCreationHelper WHERE ControlCreationHelperID = " & ControlCreationHelperID)
    Dim Direction: Direction = rs.fields("Direction")
    Dim PairSize: PairSize = rs.fields("PairSize")
    Dim Width: Width = rs.fields("Width")
    
    ''Create the form from which the controls will be placed
    Dim frm1 As Form: Set frm1 = CreateForm
    ''Set the properties of this form
    CopyControlTemplateProperties frm1
    
    ''Make sure to copy the necessary properties of the controls
    ''Attach the AfterUpdate function for the controls
    Dim ctl As control
    ''Create the from textbox
    Set ctl = CreateNamedControl(frm1, acTextBox, "txtFrom")
    CopyProperties frm1, "txtFrom", "TextControl"
    ctl.Format = "Short Date"
    ctl.AfterUpdate = "=ValidateFromToTextbox([Form]," & Esc("txtFrom") & ")"
    ''Create the from label
    Set ctl = CreateNamedControl(frm1, acLabel, "lblFrom", "txtFrom")
    CopyProperties frm1, "lblFrom", "LabelControl"
    ctl.Caption = "From"
    ''Create the to textbox
    Set ctl = CreateNamedControl(frm1, acTextBox, "txtTo")
    CopyProperties frm1, "txtTo", "TextControl"
    ctl.Format = "Short Date"
    ctl.AfterUpdate = "=ValidateFromToTextbox([Form]," & Esc("txtTo") & ")"
    ''Create the to label
    Set ctl = CreateNamedControl(frm1, acLabel, "lblTo", "txtTo")
    CopyProperties frm1, "lblTo", "LabelControl"
    ctl.Caption = "To"
    
    ''Reposition based on the direction
    If Direction = "Horizontal" Then
        RepositionControlsInRow frm1, PairSize, "lblFrom,txtFrom,lblTo,txtTo", 100, Width
    Else
       RepositionControlsInRow frm1, PairSize, "lblFrom,txtFrom", 100, Width
       RepositionControlsInRow frm1, PairSize, "lblTo,txtTo", 100, Width
    End If
    
End Sub

Public Function ValidateFromToTextbox(frm As Object, ctlName)

    Dim txtFrom: txtFrom = frm("txtFrom")
    Dim txtTo: txtTo = frm("txtTo")
    
    If txtTo < txtFrom Then
        If ctlName = "txtTo" Then
            frm("txtFrom") = txtTo
        Else
            frm("txtTo") = txtFrom
        End If
    End If
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Creation of the from and to textbox control ends here    ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function SetPrimaryKeyMainFieldAndFieldToUse(frm As Form)

    Dim Model: Model = frm("Model")
    Dim PrimaryKey, MainField, FieldCaption
    
    If Not isFalse(Model) Then
        Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE Model = " & EscapeString(Model))
        If rs.EOF Then Exit Function
        PrimaryKey = GetPrimaryKeyFromTable(rs.fields("ModelID"))
        MainField = rs.fields("MainField")
        
    End If
    
    frm("FieldToUse") = PrimaryKey
    frm("PrimaryKey") = PrimaryKey
    frm("MainField") = MainField
    
    SetFieldCaption frm
    
End Function

Public Function SetFieldCaption(frm As Form)
    
    Dim Model: Model = frm("Model")
    Dim tblName: tblName = GetTableName(Model)
    Dim MainField: MainField = frm("MainField")
    Dim FieldCaption
    
    If Not isFalse(MainField) Then
        FieldCaption = GetCaptionPropertyFromTable(tblName, MainField)
    End If
    
    frm("FieldCaption") = FieldCaption
    
End Function

Public Function CreateBlankForm(frm As Form)
    
    ''Create using CreateForm
    Dim frm1 As Form: Set frm1 = CreateForm
    ''Set the details using the copy control method
    CopyControlTemplateProperties frm1

End Function

Public Function CreateCustomControl(frm As Form)
    
    Dim ControlCreationHelperID: ControlCreationHelperID = frm("ControlCreationHelperID")
    Dim CustomControlTypeID: CustomControlTypeID = frm("CustomControlTypeID")
    Dim IsDynamicList: IsDynamicList = frm("IsDynamicList")
    Dim CustomControlType: CustomControlType = ELookup("tblCustomControlTypes", "CustomControlTypeID = " & CustomControlTypeID, "CustomControlType")
    Dim ParentModel: ParentModel = frm("ParentModel")
    
    If ExitIfTrue(isFalse(CustomControlType), "Please select a custom control type.") Then Exit Function
    
    ''If list is dynamic or a checkbox type then
    If IsDynamicList Or CustomControlType = "Checkbox" Then
        If CustomControlType = "Checkbox" Or CustomControlType = "Option Group" Then
            CreateCBorOGForm ControlCreationHelperID, CustomControlType
        End If
    ElseIf Not IsDynamicList And CustomControlType = "Option Group" Then
        CreateOptionGroup ControlCreationHelperID
    ElseIf CustomControlType = "From-To Textbox" Then
        RenderFromToTextbox frm, ControlCreationHelperID
        Exit Function
    ElseIf CustomControlType = "Listbox Button" Then
        RenderListBoxButton frm, ControlCreationHelperID
        Exit Function
    End If
    
    ''Get the modelID and the form name based on the parent model name
    Dim ModelID, VerbosePlural, frmName
    Dim modelRS As Recordset: Set modelRS = ReturnRecordset("SELECT * FROM tblModels WHERE MOdel = " & Esc(ParentModel))
    
    ''This bit will attach on current event on the forms -> optional on other control types so use proper exit function
    ''so that this part will be skipped.
    If modelRS.EOF Then Exit Function
    ModelID = modelRS.fields("ModelID")
    VerbosePlural = modelRS.fields("VerbosePlural")
    Dim PluralizedModelName: PluralizedModelName = ParentModel & "s"
    If Not isFalse(VerbosePlural) Then
        PluralizedModelName = VerbosePlural
    End If
    
    frmName = "frm" & PluralizedModelName
    
    OpenFormSetDefaults frmName, ModelID
    
End Function

Public Function RenderListBoxButton(frm As Object, ControlCreationHelperID, Optional xPosition, Optional parentFormName)
    
    ''TABLE: tblModels Fields: ModelID|Model|VerboseName|VerbosePlural|MainField|TableWideValidation|FormColumns
    ''SetFocus|IsKeyVisible|QueryName|OnFormCreate|SubformName|UserQueryFields|IsSystemTable|Timestamp|CreatedBy
    ''RecordImportID|PrimaryKey|VerbosePluralCaption
    
    ''TABLE: tblControlCreationHelper Fields: ControlCreationHelperID|CustomControlTypeID|Model|PrimaryKey
    ''MainField|FieldToUse|Direction|Timestamp|CreatedBy|RecordImportID|Width|ParentModel|IsDynamicList|FieldCaption
    ''MiddleTable|PossibleValues|NoneValue|PairSize|AsInline|ControlHeight
    Dim sqlStr: sqlStr = "SELECT * FROM tblControlCreationHelper WHERE ControlCreationHelperID = " & ControlCreationHelperID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim Model: Model = rs.fields("Model")
    Dim ModelID: ModelID = ELookup("tblModels", "Model = " & Esc(Model), "ModelID")
    Dim Width: Width = rs.fields("Width")
    Dim ControlHeight: ControlHeight = rs.fields("ControlHeight")
    
    ''Create the staging form
    If IsNull(parentFormName) Then
        Set frm = CreateForm
    Else
        Set frm = Forms(parentFormName)
    End If
    
    CopyControlTemplateProperties frm
    Dim x: x = 100
    If Not IsNull(xPosition) Then x = xPosition
    
    ''Create the Search box, Create the label, Create the listbox
    Dim ctl As control, lblName, searchName, ListName, clearSearchName
    lblName = "lbl" & Model & "Actions"
    searchName = "txt" & Model & "ActionSearch"
    ListName = "list" & Model & "Actions"
    clearSearchName = "cmdClear" & Model & "ActionSearch"
    
    ''Create the label here
    Set ctl = CreateNamedControl(frm, acLabel, lblName)
    ctl.Caption = "Actions"
    CopyProperties frm, lblName, "FilterLabelControl"
    ''Create the search box
    Set ctl = CreateNamedControl(frm, acTextBox, searchName)
    CopyProperties frm, searchName, "TextControl"
    frm(searchName).OnChange = "=FilterCustomListbox([Form]," & Esc(Model) & ")"
    ''Create a button that will clear the searchbox
    Set ctl = CreateNamedControl(frm, acCommandButton, clearSearchName)
    CopyProperties frm, clearSearchName, "TransparentButton"
    ctl.Caption = "Clear"
    ctl.OnClick = "=ClearCustomListBoxFilter([Form]," & Esc(Model) & ")"
    ''Create the listbox
    Set ctl = CreateNamedControl(frm, acListBox, ListName)
    CopyProperties frm, ListName, "ListboxControl"
    ctl.OnDblClick = "=RunFunctionFromListButton([Form]," & Esc(ListName) & ")"
    ''Set the Rowsource of this listbox
    sqlStr = "SELECT ModelButtonID,ModelButton FROM tblModelButtons WHERE ModelID = " & ModelID & " ORDER BY ModelButtonOrder"
    frm(ListName).RowSource = sqlStr
    frm(ListName).Height = ControlHeight
    ''Reposition the controls
    RepositionControlsInRow frm, "1", lblName, x, Width, , 200
    RepositionControlsInRow frm, "5,1", searchName & "," & clearSearchName, x, Width
    ''Set the height of the clearSearchName to be the same as the searchName
    frm(clearSearchName).Height = frm(searchName).Height
    RepositionControlsInRow frm, "1", ListName, x, Width
    
    ''TABLE: tblModelButtons Fields: ModelButtonID|ModelID|ModelButton|FunctionName|TableWideFunction|Timestamp
    ''CreatedBy|ModelButtonOrder|HideOnMain|HideOnForm|RecordImportID
    
    
End Function

Public Function RunFunctionFromListButton(frm As Object, ListName)

    Dim ModelButtonID: ModelButtonID = frm(ListName)
    
    If IsNull(ModelButtonID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblModelButtons WHERE ModelButtonID = " & ModelButtonID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim TableWideFunction: TableWideFunction = rs.fields("TableWideFunction")
    Dim FunctionName: FunctionName = rs.fields("FunctionName")
    
    If TableWideFunction Then
        Run FunctionName
    Else
        Run FunctionName, frm
    End If
    
End Function

Public Function ClearCustomListBoxFilter(frm As Object, Model)
    
    Dim searchName: searchName = "txt" & Model & "ActionSearch"
    frm(searchName) = ""
    
    FilterCustomListbox frm, Model
End Function

Public Function FilterCustomListbox(frm As Object, Model)

    Dim searchName: searchName = "txt" & Model & "ActionSearch"
    frm(searchName).SetFocus
    Dim searchTxt: searchTxt = frm(searchName).Text
    Dim ListName: ListName = "list" & Model & "Actions"
    
    Dim ModelID: ModelID = ELookup("tblModels", "Model = " & Esc(Model), "ModelID")
    
    Dim filters As New clsArray
    filters.Add "ModelID = " & ModelID
    
    If searchTxt <> "" Then
        filters.Add "(ModelButton Like " & Esc("*" & searchTxt & "*") & " OR FunctionName Like " & Esc("*" & searchTxt & "*") & ")"
    End If
    
    Dim filterStatement: filterStatement = filters.JoinArr(" AND ")
    
    Dim sqlStr: sqlStr = "SELECT ModelButtonID,ModelButton FROM tblModelButtons WHERE " & filterStatement & " ORDER BY ModelButtonOrder"
    frm(ListName).RowSource = sqlStr
    frm(ListName).Requery
    
End Function

Private Sub OpenFormSetDefaults(frmName, ModelID)

    On Error GoTo ErrHandler:
        DoCmd.OpenForm frmName, acDesign, , , , acHidden
        Forms(frmName).OnCurrent = "=frmDefaultOnCurrent([Form], " & ModelID & ")"
        Forms(frmName).AfterUpdate = "=frmDefaultAfterUpdate([Form], " & ModelID & ")"
        Forms(frmName).OnUnload = "=frmDefaultOnUnload([Form], " & ModelID & ")"
        DoCmd.Close acForm, frmName, acSaveYes
    Exit Sub
ErrHandler:
    MsgBox Err.Number & ": " & Err.description
    Exit Sub
    
End Sub

Private Sub CreateCBorOGForm(ControlCreationHelperID, CustomControlType)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblControlCreationHelper WHERE ControlCreationHelperID = " & ControlCreationHelperID)
    Dim Model, PrimaryKey, MainField, FieldToUse, Direction, Width, ParentModel, FieldCaption, MiddleTable, CustomControlTypeID
    Model = rs.fields("Model")
    PrimaryKey = rs.fields("PrimaryKey")
    MainField = rs.fields("MainField")
    FieldToUse = rs.fields("FieldToUse")
    Direction = rs.fields("Direction")
    Width = rs.fields("Width")
    ParentModel = rs.fields("ParentModel")
    FieldCaption = rs.fields("FieldCaption")
    MiddleTable = rs.fields("MiddleTable")
    
    ''ParentModel will be the form or model this model is base from
    ''Create the table for the filter first -> tbl & ctlName
    Dim db As Database: Set db = CurrentDb
    Dim tblName: tblName = "tbl" & ParentModel & FieldToUse
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
    Dim frmName: frmName = "cont" & ParentModel & FieldToUse
    frm.DefaultView = acDefViewContinuous
    frm.recordSource = tblName
    
    ''Create the form controls
    Dim ctlType: ctlType = acCheckBox
    If CustomControlType = "Option Group" Then
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
    
    Dim sqlStr, TableName: TableName = GetTableName(Model)
    sqlStr = "SELECT " & FieldToUse & " As [Value], " & MainField & " AS Label From " & TableName & " ORDER BY " & MainField
    
    ''Add on Load event that will delete all record and requery this record to reflect the existing records
    ''FilterContFormOnLoad(frm As Object, sqlStr As String, tblName As String)
    frm.OnLoad = "=FilterContFormOnLoad([Form]," & Esc(sqlStr) & "," & Esc(tblName) & ")"
    ''If optiongroup uncheck all then just select the existing one.
    ''Attach event to the Selected control
    ''If checkbox then leave as is.
    If CustomControlType = "Option Group" Then
        frm("Selected").AfterUpdate = "=FilterContOptionOnChange([Form]," & Esc(tblName) & "," & Esc(FieldToUse) & ")"
    End If
    
    ''Attach event to the button
    frm("cmdToggleValue").OnClick = "=ToggleFilterCB([Form]," & EscapeString(tblName) & "," & Esc(FieldToUse) & ")"
    
    
    Dim OriginalFormName: OriginalFormName = frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename frmName, acForm, OriginalFormName
    
    ''Create the form that will hold this subform
    CreateStagingForm frmName, FieldToUse, FieldCaption, Width, ParentModel, MiddleTable, tblName, CustomControlType, Model
    
End Sub

Private Function CreateStagingForm(frmName, FieldToUse, FieldCaption, Width, ParentModel, MiddleTable, tblName, CustomControlType, Model)
    
    ''FieldToUse will be the value of the checkboxes as opposed to the label which is the mainfield
    Dim frm As Form: Set frm = CreateForm
    CopyControlTemplateProperties frm
    Dim x: x = 100
    
    Dim ctl As control
'    ''Create the label at the position x and y
'    Dim lblName: lblName = "lbl" & FieldToUse
    Dim SubformName: SubformName = "sub" & FieldToUse
'    Set ctl = CreateNamedControl(frm, acLabel, lblName)
'    ''What is the Caption?
'    ctl.Caption = FieldCaption
'    CopyProperties frm, lblName, "FilterLabelControl", False
    
    ''Render the Textbox that will filter the subform
    Dim searchBoxName: searchBoxName = "txtSearch" & FieldToUse
    Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, 0)
    ctl.Name = searchBoxName
    ctl.Format = "@;""Search " & FieldCaption & """"
    ctl.OnChange = "=FilterFilterSubform([Form], " & Esc(searchBoxName) & ", " & Esc(SubformName) & " )"
    CopyProperties frm, searchBoxName, "TextControl", False
    
    ''Render the add button -> Will add items to the checkbox or optiongroup
    Dim addBtnName: addBtnName = "add" & searchBoxName
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0, 0)
    ctl.Name = addBtnName
    ctl.Caption = "Add"
    ''AddToFilterList(frm As Object, searchBoxName, SubformName, Model)
    ctl.OnClick = "=AddToFilterList([Form]," & Esc(searchBoxName) & "," & Esc(SubformName) & "," & Esc(Model) & ")"
    
    CopyProperties frm, addBtnName, "TransparentButton", False
    ctl.Height = frm(searchBoxName).Height
    
    ''Render the clear button -> Will clear the searchBoxName
    Dim clearBtnName: clearBtnName = "clear" & searchBoxName
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0, 0)
    ctl.Name = clearBtnName
    ctl.Caption = "Clear"
    ctl.OnClick = "=ClearFilterFilterSubform([Form], " & EscapeString(searchBoxName) & ", " & EscapeString(SubformName) & " )"
    
    CopyProperties frm, clearBtnName, "TransparentButton", False
    ctl.Height = frm(searchBoxName).Height
    
    If CustomControlType = "Checkbox" Then
        ''Create subform Caption -> which value is selected
        Dim lblselectedRecords: lblselectedRecords = "lbl" & SubformName
        Set ctl = CreateNamedControl(frm, acLabel, lblselectedRecords)
        ctl.Caption = "Some Caption here"
        CopyProperties frm, lblselectedRecords, "LabelControl", False
        
        ''Add a clear [Caption] button that will clear the checkboxes
        Dim cmdToggle: cmdToggle = "cmdToggle" & SubformName
        Set ctl = CreateNamedControl(frm, acCommandButton, cmdToggle)
        ctl.Caption = "Toggle " & FieldCaption
        CopyProperties frm, cmdToggle, "ButtonControl", False
        ctl.Height = frm(searchBoxName).Height
        ctl.OnClick = "=ToggleSubformCheckboxes([Form], " & EscapeString(SubformName) & " )"
    End If
    
    ''Create the subform
    Set ctl = CreateNamedControl(frm, acSubform, SubformName)
    ctl.SourceObject = frmName
    CopyProperties frm, SubformName, "SubformControl", False
    
    ''Reposition the controls
    ''RepositionControlsInRow frm, "3,3,2,2", lblName & "," & searchBoxName & "," & clearBtnName & "," & addBtnName, x, Width
    RepositionControlsInRow frm, "3,600,600", searchBoxName & "," & clearBtnName & "," & addBtnName, x, Width
    If CustomControlType = "Checkbox" Then
        RepositionControlsInRow frm, "1", lblselectedRecords, x, Width
    End If
    RepositionControlsInRow frm, "1", SubformName, x, Width
    frm(SubformName).Height = 5000
    frm(SubformName).BorderStyle = 0
    If CustomControlType = "Checkbox" Then
        RepositionControlsInRow frm, "1", cmdToggle, x, Width
    End If
    
    DoCmd.OpenForm frm.Name, acNormal
    
    ''Add on the form's on current event
    Dim onCurrentStr: onCurrentStr = "SetSubformCheck frm," & EscapeString(ParentModel) & _
                                    "," & EscapeString(SubformName) & "," & EscapeString(MiddleTable) & "," & _
                                    EscapeString(FieldToUse)
                                    
    Dim onUnloadAndAfterUpdate: onUnloadAndAfterUpdate = "=SetMiddleTableValues([Form]," & Esc(MiddleTable) & "," & Esc(tblName) & "," & Esc(FieldToUse) & _
                                                        "," & Esc(ParentModel) & ")"
                                                        
    MsgBox "Add on Form's onCurrent Event: " & onCurrentStr & vbNewLine & "onUnload and AfterUpdate:" & onUnloadAndAfterUpdate
    CopyToClipboard onCurrentStr & vbNewLine & onUnloadAndAfterUpdate
    
End Function

Public Function AddToFilterList(frm As Object, searchBoxName, SubformName, Model)
    
    Dim val: val = frm(searchBoxName)
    If isFalse(val) Then Exit Function
    ''tblModels, MainField, Model, get the TableName
    Dim sqlStr: sqlStr = "SELECT * FROM tblModels WHERE Model = " & Esc(Model)
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim ModelID: ModelID = rs.fields("ModelID")
    Dim MainField: MainField = rs.fields("MainField")
    Dim tblName: tblName = GetTableNameFromModelID(ModelID)
    
    ''Check if the value exists from the tblName
    val = EscapeString(val, tblName, MainField)
    If Not isPresent(tblName, MainField & " = " & val) Then
        RunSQL "INSERT INTO " & tblName & "(" & MainField & ") VALUES (" & val & ")"
    End If
    
    ''Add to the list
    Dim frm2 As Form: Set frm2 = frm(SubformName).Form
    Dim id: id = ELookup(tblName, MainField & " = " & val, GetPrimaryKeyFromTable(ModelID))
    tblName = frm2.recordSource
    ''FilterLabel, Selected, Value
    Dim fields As New clsArray, vals As New clsArray
    fields.arr = "FilterLabel,[Value]"
    vals.arr = val & "," & id
    RunSQL "INSERT INTO " & tblName & "(" & fields.JoinArr & ") VALUES (" & vals.JoinArr & ")"
    
    frm(SubformName).Form.Requery
    
End Function


Public Function SetFormSingleValue(frm As Object, FieldToUse)

    ''Form would be the subform where the optionboxes reside
    Dim Selected, RowValue
    Selected = frm("Selected")
    RowValue = frm("Value")
On Error Resume Next
    If Selected Then
        frm.parent(FieldToUse) = RowValue
    Else
        frm.parent(FieldToUse) = Null
    End If
    
End Function

Public Function SetMiddleTableValues(frm As Object, MiddleTable, tblName, FieldToUse, ParentModel)

    ''Select the selected ones from the tblName, Values is a string so this must be converted into a CLng
    Dim sqlStr: sqlStr = "SELECT [Value] FROM " & tblName & " WHERE Selected"
    ''Get the fieldType of the FieldToUse
    Dim valItem, valuesArr As New clsArray
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        valuesArr.Add rs.fields(0)
        rs.MoveNext
    Loop
    
    ''Get the pk of the ParentModel
    Dim ParentModelID: ParentModelID = ELookup("tblModels", "Model = " & EscapeString(ParentModel), "ModelID")
    
    ''ModelID is the parentModel as in Cards
    ''Model is the model to use as in CardCardKeyword

    ''Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    ''Get the primary key of the model
    Dim Pk: Pk = GetPrimaryKeyFromTable(ParentModelID)
    Dim pkValue: pkValue = frm(Pk)
    
    If IsNull(pkValue) Then
        Exit Function
    End If
    
    If valuesArr.count > 0 Then
        ''Delete the not existing one first
        RunSQL "Delete from " & MiddleTable & " WHERE " & FieldToUse & " Not In(" & valuesArr.JoinArr(",") & ") And " & Pk & " = " & pkValue
        ''Insert the existing ones
        For Each valItem In valuesArr.arr
            ''Check if the FieldToUse and pkValue is not existing first
            valItem = EscapeString(valItem, MiddleTable, FieldToUse)
            Dim filterStr: filterStr = FieldToUse & " = " & valItem & " AND " & Pk & " = " & pkValue
            If Not isPresent(MiddleTable, filterStr) Then
                RunSQL "INSERT INTO " & MiddleTable & "(" & FieldToUse & "," & Pk & ") VALUES (" & valItem & "," & pkValue & ")"
            End If
        Next valItem
    End If
    
End Function

''=SetSubformCheck([Form],"Card","subCardKeywordID","tblCardCardKeywords","CardKeywordID")
Public Function SetSubformCheck(frm As Object, ParentModel, SubformName, MiddleTable, FieldToUse)
    
    ''Get the pk of the ParentModel
    Dim ParentModelID: ParentModelID = ELookup("tblModels", "Model = " & EscapeString(ParentModel), "ModelID")
    
    ''ModelID is the parentModel as in Cards
    ''Model is the model to use as in CardCardKeyword

    ''Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    ''Get the primary key of the model
    Dim Pk: Pk = GetPrimaryKeyFromTable(ParentModelID)
    Dim pkValue: pkValue = frm(Pk)
    
    If IsNull(pkValue) Then
        Exit Function
    End If
    
    Dim subform As Form: Set subform = frm(SubformName).Form
    Dim recordSource: recordSource = subform.recordSource
    
    ''Run Update of the recordSource based on what belongs to the MiddleTable
    ''Select the records first then do the looping..
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    ''This will be the sql of the data to be selected from the recordsource table
    If isFalse(MiddleTable) Then
        sqlStr = "SELECT Cstr(" & FieldToUse & ") As FieldToUse FROM " & GetTableNameFromModelID(ParentModelID) & " WHERE " & Pk & " = " & pkValue
    Else
        sqlStr = "SELECT Cstr(" & FieldToUse & ") As FieldToUse FROM " & MiddleTable & " WHERE " & Pk & " = " & pkValue
    End If
    RunSQL "UPDATE " & recordSource & " SET SELECTED = False"
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = recordSource
        .SetStatement = "SELECTED = True"
        .joins.Add GenerateJoinObj(sqlStr, "[Value]", "temp", "FieldToUse")
        rowsAffected = .Run
    End With
        
    SetSubformCaption frm, recordSource, "lbl" & SubformName
    
    subform.Requery
    
End Function

Public Function ToggleSubformCheckboxes(frm As Object, SubformName)
    
    ''Get the recordset
    ''Get the total number, count the selected, if selected is more then uncheck everything, vice versa
    Dim subform As Form: Set subform = frm(SubformName).Form
    Dim recordSource: recordSource = subform.recordSource
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM " & recordSource & " WHERE SELECTED")
    Dim rsCount:  rsCount = ECount(recordSource, "ID > 0")
    Dim selectedCount: selectedCount = ECount(recordSource, "SELECTED")
    
    Dim setSelectedAs: setSelectedAs = "True"
    If Divide(selectedCount, rsCount) > 0.5 Then
        setSelectedAs = "False"
    End If
    
    RunSQL "UPDATE " & recordSource & " SET Selected = " & setSelectedAs
    frm(SubformName).Form.Requery
    
End Function

Public Function GetControlPosition(frm, ctlName, query) As Double
    
    Select Case query
        Case "bottom":
            GetControlPosition = frm(ctlName).Top + frm(ctlName).Height
        Case "right":
            GetControlPosition = frm(ctlName).Left + frm(ctlName).Width
        Case "Top":
            GetControlPosition = frm(ctlName).Top
        Case "Left"
            GetControlPosition = frm(ctlName).Left
    End Select
    
End Function

Public Function CreateNamedControl(frm As Object, ControlType As AcControlType, labelName, Optional parent) As control
    
    Dim ctl: Set ctl = CreateControl(frm.Name, ControlType, , parent, , 0, 0, 0, 0)
    ctl.Name = labelName
    Set CreateNamedControl = ctl
    
End Function


