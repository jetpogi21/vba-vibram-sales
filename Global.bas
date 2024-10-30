Attribute VB_Name = "Global"
Option Compare Database
Option Explicit

Public g_userID As Variant
Public g_SiteAccess As Variant
Public g_FrontEndVersion As Variant
Public g_Language As Variant

Public Const AllQuery As String = "SELECT All_Number,[All] from tblAlls Order by All_Number"

Public Function Prompt_Close(Saved As Boolean) As Boolean
    Dim response As Integer
    If Not Saved Then
        response = MsgBox("Any changes on this record will not be saved." & vbCrLf & "Do you want to close this form?", vbCritical + vbYesNo)
        If response = vbNo Then
            Prompt_Close = False
            Exit Function
        End If
    End If
    
    Prompt_Close = True
End Function

Public Function are_data_valid(frm As Form) As Boolean
    On Error GoTo Err_are_data_valid
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("select * from tblFields where FormName = '" & frm.Name & "'")
    
    Dim ControlName As String, ControlCaption As Variant, ValidationRules As Variant, ValidationArray() As String, ValidationRule As Variant
    Dim ctl As control
    
    Do Until rs.EOF
        ControlName = rs.fields("ControlName"): ControlCaption = rs.fields("ControlCaption"): ValidationRules = rs.fields("ValidationRule")
        If ControlExists(ControlName, frm) Then
            Set ctl = frm.controls(ControlName)
            If Not IsNull(ValidationRules) Then
                ValidationArray = Split(ValidationRules)
                For Each ValidationRule In ValidationArray
                    Select Case Trim(ValidationRule)
                    Case "required":
                        If IsNull(ctl) Or ctl = "" Then
                            MsgBox ControlCaption & " is a required field.", vbCritical + vbOKOnly
                            ctl.SetFocus
                            are_data_valid = False
                            Exit Function
                        End If
                    Case "positive":
                        If ctl < 0 Then
                            MsgBox ControlCaption & " must be not be less than 0.", vbCritical + vbOKOnly
                            ctl.SetFocus
                            are_data_valid = False
                            Exit Function
                        End If
                    Case "uniqueBarcode":
                        If Not isBarcodeUnique(ctl.value) Then
                            ctl.SetFocus
                            are_data_valid = False
                            Exit Function
                        End If
                    End Select
                Next ValidationRule
            End If
        End If
        rs.MoveNext
    Loop
    
    are_data_valid = True

Exit_are_data_valid:
    Exit Function
Err_are_data_valid:
    LogError Err.Number, Err.description, "are_data_valid"
    Resume Exit_are_data_valid
End Function

Public Function AllocateDataToFields(frm As Object, RecordID As Variant)

    Dim formProperties As Recordset
    Set formProperties = CurrentDb.OpenRecordset("SELECT * FROM tblMainForms WHERE MainFormName = '" & frm.Name & "'")
    Dim ViewTable, PrimaryKey As String
    ViewTable = formProperties.fields("ViewTable")
    PrimaryKey = formProperties.fields("PrimaryKey")

    Dim row As Recordset
    Dim fields As Recordset

    Set row = CurrentDb.OpenRecordset("SELECT * FROM " & ViewTable & " WHERE " & PrimaryKey & " = " & RecordID)
    Set fields = CurrentDb.OpenRecordset("SELECT * FROM tblFields WHERE FormName = '" & frm.Name & "'")

    Dim ControlName As String
    If Not row.EOF Then
        Do Until fields.EOF
            ControlName = fields.fields("ControlName")
            frm.controls(ControlName) = row.fields(ControlName)
            fields.MoveNext
        Loop
    End If

End Function

Public Function isBarcodeUnique(ctl As String) As Boolean
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM qryProductBarcodes where BarcodeNumber = '" & ctl & "' And Active = -1")
    
    If Not rs.EOF Then
        MsgBox "Barcode " & ctl & " is already associated to SKU " & rs.fields("SKU"), vbCritical + vbOKOnly
        isBarcodeUnique = False
        Exit Function
    End If
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblProducts where SKU = '" & ctl & "' And Active = -1")
    If Not rs.EOF Then
        MsgBox "The barcode entered: " & ctl & " is an exact match for SKU and can't be added as it is not unique.", vbCritical + vbOKOnly
        isBarcodeUnique = False
        Exit Function
    End If
    
    isBarcodeUnique = True
End Function

Public Function cancel_and_close(frm As Form)
    frm.Undo
    DoCmd.Close acForm, frm.Name, acSaveNo
End Function

Public Function save_record(frm As Object, Optional saveType As Integer)
    If are_data_valid(frm) Then
        RunCommandSaveRecord
        If saveType = 0 Then
            DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
        ElseIf saveType = 1 Then
            DoCmd.Close acForm, frm.Name, acSaveNo
        End If
    End If
End Function

Public Function delete_record(TableName As String, record_id As Variant, field_name As String, formToRequery As Form)
    If IsNull(record_id) Then
        MsgBox "Record not found.", vbCritical + vbOKOnly, "Error: Record"
        Exit Function
    End If
    
    If MsgBox("Are you sure you want to make this record inactive?", vbYesNo, "Delete Prompt") = vbNo Then
        Exit Function
    End If
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE " & TableName & " SET Active = 0 WHERE " & field_name & " = " & record_id
    DoCmd.SetWarnings True
    
    Insert_Delete_Log TableName, "DELETE", record_id
    
    If Not formToRequery Is Nothing Then
        formToRequery.Requery
    End If
    
End Function

Public Function OpenForm(frmName As String, openArgs As Variant)
    If CurrentProject.AllForms(frmName).IsLoaded Then
        DoCmd.Close acForm, frmName, acSaveNo
    End If
    
    DoCmd.OpenForm frmName, , , , , , openArgs
    
End Function

Public Function open_form(frmName As String, Optional record_id_field As String, Optional frm As Form)

On Error GoTo Err_Handler:
    
    ''IF form is loaded then close the form first
    If CurrentProject.AllForms(frmName).IsLoaded Then
        DoCmd.Close acForm, frmName, acSaveNo
    End If

    If record_id_field = "" Then
    
        DoCmd.OpenForm frmName, , , , acFormAdd
        
    Else
        
        Dim record_id
        record_id = frm.controls(record_id_field)
        
        If isFalse(record_id) Then
            MsgBox "Record not found.", vbCritical + vbOKOnly, "Error: Record"
        Else
            DoCmd.OpenForm frmName, , , , , , record_id
        End If
        
    End If

    Exit Function
Err_Handler:

    If Err.Number = 2427 Then
        ShowError "There is no record selected.."
        Exit Function
    End If
    
End Function

Public Function Form_Update_Log(PrimaryKeyName As String, TableName As String, frm As Form)
    Dim RecordID As Variant
    If frm.NewRecord Then
        RecordID = DMax(PrimaryKeyName, TableName)
        If IsNull(RecordID) Then
            RecordID = 1
        Else
            RecordID = RecordID + 1
        End If
        
        Insert_Delete_Log TableName, "ADD", RecordID
    Else
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset(TableName)
        Dim fld As field
        Dim oldValue As String
        Dim newValue As String
        RecordID = frm.controls(PrimaryKeyName)
        For Each fld In rs.fields
            If ControlExists(fld.Name, frm) Then
                If frm.controls(fld.Name).oldValue <> frm.controls(fld.Name) Then
                    oldValue = frm.controls(fld.Name).oldValue
                    newValue = frm.controls(fld.Name)
                    Update_Log TableName, oldValue, newValue, RecordID, fld.Name
                End If
            End If
        Next fld
    End If
End Function

Public Function Insert_Delete_Log(TableName As String, EventName As String, RecordID As Variant)

    Dim UserID As Integer
    If isFalse(g_userID) Then
        UserID = 0
    Else
        UserID = g_userID
    End If

    Dim computerName As String: computerName = Environ$("computername")
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO tblLogs (UserID,EventName,ComputerName,TableName,RecordID) VALUES (" & UserID & ",'" & EventName & "','" & computerName & "'" & _
                 ",'" & TableName & "','" & RecordID & "')"
    DoCmd.SetWarnings True
    
End Function

Public Function Update_Log(TableName As String, oldValue As Variant, newValue As Variant, RecordID As Variant, fieldName As String)
    
    If TableName = "tblClipboardForms" Then Exit Function
    Dim UserID As Integer
    If isFalse(g_userID) Then
        UserID = 0
    Else
        UserID = g_userID
    End If
    
    If Not IsNull(oldValue) Then
        oldValue = replace(oldValue, Chr(34), Chr(34) & Chr(34))
    End If
    
    If Not IsNull(newValue) Then
        newValue = replace(newValue, Chr(34), Chr(34) & Chr(34))
    End If
    
    Dim computerName As String: computerName = Environ$("computername")
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO tblLogs (UserID,EventName,ComputerName,TableName,RecordID,OldValue,NewValue,FieldName) VALUES (" & UserID & ",'UPDATE','" & computerName & "'" & _
                 ",'" & TableName & "','" & RecordID & "'," & EscapeString(oldValue) & "," & EscapeString(newValue) & ",'" & fieldName & "')"
                 
    DoCmd.SetWarnings True
End Function

Function ControlExists(ControlName, FormCheck As Form) As Boolean
    Dim strTest As String
    On Error Resume Next
    strTest = FormCheck(ControlName).Name
    ControlExists = (Err.Number = 0)
End Function

Public Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element As Variant
    On Error GoTo IsInArrayError:                'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function

Public Function Insert_Data(frm As Form)

    Dim sqlStr As String
    Dim fields() As String
    Dim fieldValues() As String
    Dim formProperties As Recordset
    Dim ctl As control
    Set formProperties = CurrentDb.OpenRecordset("SELECT * FROM tblMainForms WHERE MainFormName = '" & frm.Name & "'")
    Dim TableName As String
    TableName = formProperties.fields("TableName")
    Dim FieldsRS As Recordset
    Set FieldsRS = CurrentDb.OpenRecordset("SELECT * FROM tblFields WHERE FormName = '" & frm.Name & "'")
    Dim i As Integer
    Dim RecordID As Variant
    ''Loop the Fields
    Do Until FieldsRS.EOF
        
        ReDim Preserve fields(i)
        ReDim Preserve fieldValues(i)
    
        fields(i) = FieldsRS.fields("ControlName")
        Set ctl = frm.controls(fields(i))
        If IsNull(ctl) Then
            fieldValues(i) = "Null"
        Else
            Select Case FieldsRS.fields("ControlType")
            Case "String":
                fieldValues(i) = """" & ctl & """"
            Case "Integer":
                fieldValues(i) = ctl
            Case "Date":
                fieldValues(i) = "#" & SQLDate(ctl) & "#"
            End Select
        End If
        i = i + 1
        
        FieldsRS.MoveNext
    Loop
    
    sqlStr = "INSERT INTO " & TableName & " (" & Join(fields, ",") & ") VALUES (" & Join(fieldValues, ",") & ")"
    CurrentDb.Execute sqlStr
    RecordID = CurrentDb.OpenRecordset("SELECT @@IDENTITY")(0)
    frm.Undo
    Insert_Delete_Log TableName, "ADD", RecordID

End Function

Public Function Set_Recordsource(frm As Form)

    Dim formProperties As Recordset
    Set formProperties = CurrentDb.OpenRecordset("SELECT * FROM tblMainForms WHERE MainFormName = '" & frm.Name & "'")
    
    Dim ViewTable, TableName, PrimaryKey As String
    ViewTable = formProperties.fields("ViewTable")
    TableName = formProperties.fields("TableName")
    PrimaryKey = formProperties.fields("PrimaryKey")

    Dim logSQL As String
    logSQL = "SELECT UserName,DateTime,RecordID FROM qryLogs WHERE EventName = 'ADD' And TableName = '" & TableName & "'"
    
    Dim sql As String
    sql = "SELECT " & ViewTable & ".*, qryLogs.UserName As UserLog, qryLogs.DateTime As LoggedDate FROM " & ViewTable & _
          " LEFT JOIN (" & logSQL & ") qryLogs ON " & ViewTable & "." & PrimaryKey & " = qryLogs.RecordID WHERE Active = -1"
    
    ''Set the record source of the datasheet
    frm.recordSource = sql
    
    SetDatasheetCaption frm

End Function

Public Function SetDatasheetCaption(frm As Form)
    
    ''Captions
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * from tblFields where FormName = '" & frm.Name & "'")
    Dim ctl As control
    Do Until rs.EOF
        If ControlExists(rs.fields("ControlName"), frm) Then
            Set ctl = frm.controls(rs.fields("ControlName"))
            ctl.Properties("DatasheetCaption") = rs.fields("ControlCaption")
            ctl.Locked = True
            If Not isFalse(rs.fields("FieldWidth")) Then
                ctl.ColumnWidth = rs.fields("FieldWidth")
            End If
        End If
        rs.MoveNext
    Loop
    
End Function

Public Function Update_Data(frm As Object, RecordID As Variant)

    ''Check for field changes
    Dim formProperties As Recordset
    Set formProperties = CurrentDb.OpenRecordset("SELECT * FROM tblMainForms WHERE MainFormName = '" & frm.Name & "'")
    Dim ViewTable, PrimaryKey, TableName As String
    ViewTable = formProperties("ViewTable")
    PrimaryKey = formProperties("PrimaryKey")
    TableName = formProperties("TableName")
    Dim row As Recordset
    Dim fields As Recordset
    Set row = CurrentDb.OpenRecordset("SELECT * FROM " & ViewTable & " WHERE " & PrimaryKey & " = " & RecordID)
    Set fields = CurrentDb.OpenRecordset("SELECT * FROM tblFields WHERE FormName = '" & frm.Name & "'")
    Dim ControlName As String                    ''Current ControlName
    Dim oldValue As Variant                      ''Old Value -> To be compared from current value
    Dim updateArray() As String                  ''Array of Update Statements
    Dim fieldValue As String                     ''New Value
    Dim ctl As control                           ''Current Control
    Dim i As Integer
    Dim sqlStr As String
    If Not row.EOF Then
        Do Until fields.EOF
            ControlName = fields.fields("ControlName")
            Set ctl = frm.controls(ControlName): oldValue = row.fields(ControlName)
            If ctl <> oldValue Or (Not IsNull(ctl) And IsNull(oldValue)) Then
                Select Case fields.fields("ControlType")
                Case "String":
                    fieldValue = """" & ctl & """"
                Case "Integer":
                    fieldValue = ctl
                Case "Date":
                    fieldValue = "#" & SQLDate(ctl) & "#"
                End Select
                ReDim Preserve updateArray(i)
                updateArray(i) = ControlName & " = " & fieldValue
                i = i + 1
                Update_Log TableName, oldValue, ctl, RecordID, ControlName
            ElseIf IsNull(ctl) And Not IsNull(oldValue) Then
                fieldValue = "Null"
                ReDim Preserve updateArray(i)
                updateArray(i) = ControlName & " = " & fieldValue
                i = i + 1
                Update_Log TableName, oldValue, ctl, RecordID, ControlName
            End If
            fields.MoveNext
        Loop
    End If
    
    If Join(updateArray, ",") <> "" Then
        sqlStr = "UPDATE " & TableName & " SET " & Join(updateArray, ",") & " WHERE " & PrimaryKey & " = " & RecordID
        CurrentDb.Execute sqlStr
    End If
    
End Function

Public Function Reset_Form(frm As Form)
    
    Dim fields As Recordset
    Set fields = CurrentDb.OpenRecordset("SELECT * FROM tblFields WHERE FormName = '" & frm.Name & "'")
    Dim ctl As control
    Do Until fields.EOF
        Set ctl = frm.controls(fields.fields("ControlName"))
        ctl = fields.fields("DefaultValue")
        fields.MoveNext
    Loop

End Function

Public Function Attach_Events_To_Form(frm As Form)

    ''Set Form Caption
    Dim formProperties As Recordset
    Set formProperties = CurrentDb.OpenRecordset("SELECT * FROM tblMainForms WHERE MainFormName = '" & frm.Name & "'")
    Dim FormToOpen, PrimaryKey, Caption, TableName As String
    FormToOpen = formProperties.fields("FormToOpen")
    PrimaryKey = formProperties.fields("PrimaryKey")
    Caption = formProperties.fields("Caption")
    TableName = formProperties.fields("TableName")

    Dim ctl As control

    If Not formProperties.EOF Then

        frm.Caption = Caption
        Set ctl = frm.controls("cmdAdd_New")
        ctl.OnClick = "=open_form(""" & FormToOpen & """)"
        Set ctl = frm.controls("cmdView")
        ctl.OnClick = "=open_form(""" & FormToOpen & """,""" & PrimaryKey & """,[Form].[subform].[Form])"
        Set ctl = frm.controls("cmdDelete")
        ctl.OnClick = "=delete_record(""" & TableName & """,[Form].[subform].[Form].[" & PrimaryKey & "],""" & PrimaryKey & """,[Form].[subform].[Form])"
        
        '        If Not isFalse(formProperties.Fields("FormWidth")) Then
        '            frm.InsideWidth = formProperties.Fields("FormWidth")
        '        End If
        '
        '        If Not isFalse(formProperties.Fields("FormHeight")) Then
        '            frm.Detail.Height = formProperties.Fields("FormHeight")
        '        End If
        
    End If

End Function

Public Function is_an_exception(OrderID As Variant) As Boolean
    If IsNull(OrderID) Then
        is_an_exception = False
        Exit Function
    End If
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblOrderComments WHERE OrderID = " & OrderID)
    
    If rs.EOF Then
        is_an_exception = False
        Exit Function
    End If
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblOrderComments WHERE OrderID = " & OrderID & " And Exception = -1 And ExceptionCleared = 0")
    
    If Not rs.EOF Then
        is_an_exception = True
        Exit Function
    End If
    
    is_an_exception = False
    
End Function

Public Function has_barcode(ProductID As Variant) As Boolean
    If IsNull(ProductID) Then
        has_barcode = False
        Exit Function
    End If

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblProductBarcodes WHERE ProductID = " & ProductID & " And Active = -1")

    If rs.EOF Then
        has_barcode = False
    Else
        has_barcode = True
    End If
    
End Function

Public Function backordered_qty(OrderID As Variant) As Integer
    If IsNull(OrderID) Then
        backordered_qty = 0
        Exit Function
    End If
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblOrdersProducts WHERE OrderID = " & OrderID & " And QTYBackordered > 0")
    
    If rs.EOF Then
        backordered_qty = 0
        Exit Function
    End If
    
    Dim i As Integer
    Do Until rs.EOF
        i = i + rs.fields("QTYBackordered")
        rs.MoveNext
    Loop
    
    backordered_qty = i
    
End Function

Public Function SQLDate(varDate As Variant, Optional isPostgreSQL As Boolean = False) As String
    'Purpose:    Return a delimited string in the date format used natively by JET SQL.
    'Argument:   A date/time value.
    'Note:       Returns just the date format if the argument has no time component,
    '                or a date/time format if it does.
    'Author:     Allen Browne. allen@allenbrowne.com, June 2006.
    If IsDate(varDate) Then
        If DateValue(varDate) = varDate Then
            SQLDate = Format$(varDate, IIf(isPostgreSQL, "yyyy-mm-dd", "yyyy/mm/dd"))
        Else
            SQLDate = Format$(varDate, IIf(isPostgreSQL, "yyyy-mm-dd hh:nn:ss", "yyyy/mm/dd hh:nn:ss"))
        End If
    End If
    
    SQLDate = replace(SQLDate, ". ", "-")
    SQLDate = replace(SQLDate, ".", "-")
    
End Function

Public Function TruncateText(str As String, size As Integer) As String
    Dim truncatedStr As String
    If Len(str) > size Then
        truncatedStr = Left(str, size - 3) & "..."
        TruncateText = truncatedStr
        Exit Function
    End If
    
    TruncateText = str
    
End Function

Public Function SetCaptionAndWidth(TableName As String, frm As Form)
    ''Captions
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * from tblTableFields where TableName = """ & TableName & """")
    Dim ctl As control
    Do Until rs.EOF
        If ControlExists(rs.fields("FieldName"), frm) Then
            Set ctl = frm.controls(rs.fields("FieldName"))
            ctl.Properties("DatasheetCaption") = rs.fields("FieldCaption")
            If Not isFalse(rs.fields("ColumnWidth")) Then
                ctl.ColumnWidth = rs.fields("ColumnWidth")
            End If
        End If
        rs.MoveNext
    Loop
End Function

Public Function InsertAndLog(tblName As String, fields() As String, fieldValues() As String) As Variant

    RunSQL "INSERT INTO " & tblName & " (" & Join(fields, ",") & ") VALUES (" & Join(fieldValues, ",") & ")"
    Dim RecordID
    RecordID = CurrentDb.OpenRecordset("SELECT @@IDENTITY")(0)
    Insert_Delete_Log tblName, "ADD", RecordID
    
    InsertAndLog = RecordID
    
End Function

Public Function GetFirstDayOfWeek(nYear As Long, nWeek As Integer) As Date

    Dim nTrimester As Integer, _
    wd As Integer, _
    StartDate As Date, _
    inputDate As Date
    
    inputDate = DateSerial(nYear, 1, 1)
    inputDate = DateAdd("ww", nWeek - 1, inputDate)
    wd = weekday(inputDate)
    
    GetFirstDayOfWeek = DateAdd("d", 1 - wd, inputDate)
    
End Function

Public Function GetFirstDayOfNextMonth(vDate As Variant) As Date
    
    Dim firstThis
    firstThis = DateSerial(Year(vDate), Month(vDate), 1)
    GetFirstDayOfNextMonth = DateAdd("m", 1, firstThis)
    
End Function

Public Function GetGlobalSetting(ByVal GlobalSettingName As String) As Variant
    GetGlobalSetting = ELookup("tblGlobalSettings", "GlobalSetting = '" & GlobalSettingName & "'", "GlobalSettingValue")
End Function

Public Function GetSecureSetting(ByVal SecureSettingName As String) As Variant
    GetSecureSetting = ELookup("tblSecureSettings", "SecureSettingName = '" & SecureSettingName & "' AND WebEndpoint = " & GetGlobalSetting("WebEndpoint") & "", "SecureSettingValue")
End Function

Function LogError(ByVal lngErrNumber As Long, ByVal strErrDescription As String, _
                  strCallingProc As String, Optional vParameters, Optional bShowUser As Boolean = True) As String
    On Error GoTo Err_LogError
    ' Purpose: Generic error handler.
    ' Logs errors to table "tLogError".
    ' Arguments: lngErrNumber - value of Err.Number
    ' strErrDescription - value of Err.Description
    ' strCallingProc - name of sub|function that generated the error.
    ' vParameters - optional string: List of parameters to record.
    ' bShowUser - optional boolean: If False, suppresses display.
    ' Author: Allen Browne, allen@allenbrowne.com

    Dim strMsg As String                         ' String for display in MsgBox
    Dim rst As DAO.Recordset                     ' The tLogError table

    Select Case lngErrNumber
    Case 0
        Debug.Print strCallingProc & " called error 0."
    Case 2501                                    ' Cancelled
        'Do nothing.
    Case 3314, 2101, 2115                        ' Can't save.
        If bShowUser Then
            strMsg = "Record cannot be saved at this time." & vbCrLf & _
                     "Complete the entry, or press <Esc> to undo."
            MsgBox strMsg, vbExclamation, strCallingProc
        End If
    Case Else
        If bShowUser Then
            strMsg = "Error " & lngErrNumber & ": " & strErrDescription
            MsgBox strMsg, vbExclamation, strCallingProc
        End If
        Set rst = CurrentDb.OpenRecordset("tblErrorLog", , dbAppendOnly)
        rst.addNew
        rst![ErrNumber] = lngErrNumber
        rst![ErrDescription] = Left$(strErrDescription, 255)
        rst![ErrDate] = Now()
        rst![CallingProc] = strCallingProc
        rst![UserName] = CurrentUser() & " - " & g_userID
        rst![ShowUser] = bShowUser
        If Not IsMissing(vParameters) Then
            rst![Parameters] = Left(vParameters, 255)
        End If
        rst.Update
        rst.Close
        LogError = strMsg
    End Select
    
Exit_LogError:
    Set rst = Nothing
    Exit Function

Err_LogError:
    strMsg = "An unexpected situation arose in your program." & vbCrLf & _
             "Please write down the following details:" & vbCrLf & vbCrLf & _
             "Calling Proc: " & strCallingProc & vbCrLf & _
             "Error number " & lngErrNumber & vbCrLf & strErrDescription & vbCrLf & vbCrLf & _
             "Unable to record because Error " & Err.Number & vbCrLf & Err.description
    LogError = strMsg
    MsgBox strMsg, vbCritical, "LogError()"
    Resume Exit_LogError
End Function




