Attribute VB_Name = "Validation"
Option Compare Database
Option Explicit

Public Function areDataValid(frm As Object, Optional tblName As String) As Boolean
    
    ''Fetch the tblFormFields from the specific form
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("select * from tblFormFields where TableName = '" & tblName & "'")
    
    Dim ControlName As String, ControlCaption As Variant, ValidationRules As Variant, ValidationArray() As String, ValidationRule As Variant
    Dim ctl As control
    
    Do Until rs.EOF
        ControlName = rs.fields("FieldName"): ControlCaption = rs.fields("FieldCaption"): ValidationRules = rs.fields("ValidationString")
        If ControlExists(ControlName, frm) Then
            Set ctl = frm.controls(ControlName)
            If Not IsNull(ValidationRules) Then
                ValidationArray = Split(ValidationRules)
                For Each ValidationRule In ValidationArray
                    Select Case Trim(ValidationRule)
                        Case "required":
                            If IsNull(ctl) Or ctl = "" Then
                                MsgBox ControlCaption & " is a required field.", vbCritical + vbOKOnly
                                If ControlExists(ctl.Name, frm) And ctl.ColumnHidden = False Then
                                    ctl.SetFocus
                                End If
                                areDataValid = False
                                DoCmd.CancelEvent
                                Exit Function
                            End If
                        Case "positive":
                            If ctl < 0 Then
                                MsgBox ControlCaption & " must be not be less than 0.", vbCritical + vbOKOnly
                                If ControlExists(ctl.Name, frm) And ctl.ColumnHidden = False Then
                                    ctl.SetFocus
                                End If
                                areDataValid = False
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
    Set rs = CurrentDb.OpenRecordset("select * from tblFormFields where TableName = '" & tblName & "' And ValidationString Like '*unique*'")
    Dim i As Integer
    Dim filterStr() As String
    Dim fieldCaptions() As String
    Dim fieldValue As String
    If Not rs.EOF Then
        Do Until rs.EOF
            fieldValue = frm(rs.fields("FieldName"))
            ReDim Preserve filterStr(i)
            Select Case rs.fields("FieldTypeID")
                Case 10:
                    fieldValue = "'" & fieldValue & "'"
            End Select
            filterStr(i) = rs.fields("FieldName") & " = " & fieldValue
            ReDim Preserve fieldCaptions(i)
            fieldCaptions(i) = rs.fields("FieldCaption")
            i = i + 1
            rs.MoveNext
        Loop
        
        Dim filterStmt As String
        filterStmt = Join(filterStr, " And ")
        
        ''If not a new record then disregard this record from the filter
        If Not frm.NewRecord Then
            Dim PrimaryKey As String
            PrimaryKey = ELookup("tblTables", "TableName = '" & tblName & "'", "PrimaryKey")
            filterStmt = filterStmt & " And " & PrimaryKey & " <> " & frm(PrimaryKey)
        End If
        
        Dim errorMsg As String
        errorMsg = Join(fieldCaptions, " | ") & " is already present from the record list"
        
        Dim ViewName As String
        ViewName = ELookup("tblTables", "TableName = '" & tblName & "'", "ViewName")
        
        If isPresent(ViewName, filterStmt) Then
            ShowError errorMsg
            DoCmd.CancelEvent
            areDataValid = False
            Exit Function
        End If
        
    End If
    
    ''Look for AdditionalValidation from tblTables
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblTables WHERE TableName = '" & tblName & "'")
    If Not rs.EOF Then
        If Not IsNull(rs.fields("AdditionalValidation")) Then
            ValidationArray = Split(rs.fields("AdditionalValidation"))
            For Each ValidationRule In ValidationArray
                If Not Application.Run(ValidationRule, frm) Then
                    areDataValid = False
                    DoCmd.CancelEvent
                    Exit Function
                End If
            Next ValidationRule
        End If
    End If
    
    If frm.NewRecord Then
        frm.OnClose = "=RequeryOnClose('" & tblName & "',True)"
    End If
    
    areDataValid = True
    
End Function

Public Function ImportFileValidation(frm As Form) As Boolean
    
    Dim fields(1) As String, fieldValues(1) As String, i As Integer, FieldType As Integer, filterStmt(1) As String
    fields(0) = "ImportPlatformID"
    fields(1) = "ImportFileOrder"
    
    For i = 0 To UBound(fields)
        FieldType = CInt(ELookup("tblFormFields", "TableName = 'tblImportFiles' And FieldName = '" & fields(i) & "'", "FieldTypeID"))
        fieldValues(i) = ReturnStringBasedOnType(frm(fields(i)), FieldType)
        filterStmt(i) = fields(i) & " = " & fieldValues(i)
    Next i
    
    Dim filterStr As String
    filterStr = Join(filterStmt, " And ")
    
    If Not IsNull(frm("ImportFileID")) Then
        filterStr = filterStr & " And ImportFileID <> " & frm("ImportFileID")
    End If
    
    If isPresent("tblImportFiles", filterStr) Then
        ShowError "Platform | File Order is already existing.."
        ImportFileValidation = False
    Else
        ImportFileValidation = True
    End If
    
End Function

Public Function BarcodeValidation(frm As Form) As Boolean
    
    Dim filterStr, BarcodeNumber
    BarcodeNumber = frm("BarcodeNumber")
    filterStr = "SKU = '" & BarcodeNumber & "'"
    
    If isPresent("tblProducts", filterStr) Then
        ShowError BarcodeNumber & " is already present as an SKU"
        BarcodeValidation = False
    Else
        BarcodeValidation = True
    End If
    
End Function

'Public Function UserUserGroupValidation(frm As Form) As Boolean
'
'    Dim UserGroupID, UserID
'    UserGroupID = frm("UserGroupID")
'    UserID = frm("UserID")
'
'    If UserGroupID = 1 Then
'        ''Lookup if there's already an exisiting access account
'        If isPresent("tblUserUserGroups", "UserID = " & UserID) Then
'            ShowError "This user already has an access group." & vbCrLf & _
'                "Please delete any esxisting group first to continue using the administrator group.."
'            UserUserGroupValidation = False
'            Exit Function
'        End If
'    Else
'        ''Lookup if there's already an admin account
'        If isPresent("tblUserUserGroups", "UserID = " & UserID & " And UserGroupID = 1") Then
'            ShowError "This user is already an administrator so there's no need to add more access.."
'            UserUserGroupValidation = False
'            Exit Function
'        End If
'
'    End If
'
'    UserUserGroupValidation = True
'
'End Function

Public Function ValidateExceptionResolution(frm As Form) As Boolean
    
    If frm.ExceptionCleared = -1 Then
        If isFalse(frm.ExceptionResolution) Then
            MsgBox "If exception cleared, resolution is a required field..", vbOKOnly + vbCritical
            ValidateExceptionResolution = False
            Exit Function
        End If
        'RunSQL "UPDATE tblOrders SET tblOrders.Exception = " & False & " WHERE OrderID = " & Me.OrderID.Value
    End If
    
    frm.ExceptionClearedDateTime = Now()
    frm.ExceptionClearedUser = g_userID
    
    ValidateExceptionResolution = True
    
End Function
