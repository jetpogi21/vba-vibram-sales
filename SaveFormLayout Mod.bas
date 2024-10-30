Attribute VB_Name = "SaveFormLayout Mod"
Option Compare Database
Option Explicit

Public Function SaveFormLayout(frm As Form)

    If Not areDataValid2(frm, "SaveFormLayout") Then Exit Function
    
    Dim FormName
    FormName = frm("FormName")
           
    RunSQL "DELETE FROM tblForms WHERE FormName = " & Esc(FormName)
    Dim propertyArr As New clsArray
    propertyArr.arr = "Top,Left,Height,Width"
    
    ''Open the form in design view
    Dim FormID
    FormID = ELookup("tblForms", "FormName = " & EscapeString("FormName"), "FormID")
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    If FormID = "" Then
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "INSERT"
            .Source = "tblForms"
            .fields = "FormName"
            .insertValues = EscapeString(FormName)
            rowsAffected = .Run
            FormID = .LastInsertID
        End With
    End If
    
    ''Delete all the data from tblFormControls using FormID
    RunSQL "DELETE FROM tblFormControls WHERE FormID = " & FormID
    
    Dim ctl As control, frm2 As Object
    Set frm2 = GetForm(FormName, True, True)
    
    If frm2 Is Nothing Then
        Set frm2 = GetReport(FormName, True, True)
    End If
    
    Dim property
    For Each ctl In frm2.controls
        For Each property In propertyArr.arr
            Set sqlObj = New clsSQL
            With sqlObj
                .SQLType = "INSERT"
                .Source = "tblFormControls"
                .fields = "FormControlName,FormID,FormControlProperty,FormControlProperyValue"
                .insertValues = EscapeString(ctl.Name) & "," & _
                                FormID & "," & _
                                EscapeString(property) & "," & _
                                EscapeString(ctl.Properties(property))
                rowsAffected = .Run
            End With
        Next property
    Next ctl
    
    DoCmd.Close IIf(IsObjectAReport(frm2), acReport, acForm), FormName
    MsgBox "Layout Saved..."
    
End Function

Public Function LoadSavedFormLayout(FormName, Optional ShouldCloseForm As Boolean = True)

    ''This will load the form layout saved
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryFormControls WHERE FormName = " & EscapeString(FormName))
    
    If rs.EOF Then Exit Function
    
    Dim frm As Object
    Set frm = GetForm(FormName, True, True)
    
    If frm Is Nothing Then
        Set frm = GetReport(FormName, True, True)
    End If
    
    If frm Is Nothing Then Exit Function
    
    Dim FormControlID, FormControlName, FormControlProperty, FormControlProperyValue
    
    Do Until rs.EOF
        
        FormControlName = rs.fields("FormControlName")
        FormControlProperty = rs.fields("FormControlProperty")
        FormControlProperyValue = rs.fields("FormControlProperyValue")
        
        frm(FormControlName).Properties(FormControlProperty) = FormControlProperyValue
        rs.MoveNext
        
    Loop
    
    If ShouldCloseForm Then
        frm.Width = GetMaxX(frm) + 400
        frm.Section("Detail").Height = GetMaxY(frm) + 400
        DoCmd.Close acForm, FormName, acSaveYes
    Else
        frm.Width = 0
    End If
    

End Function


