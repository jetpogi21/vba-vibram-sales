Attribute VB_Name = "UnboundForm Mod"
Option Compare Database
Option Explicit

Public Function UnboundFormCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function UnboundFormAddRecord(frm As Object, tblName)

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM " & tblName)
  
    ''Dim pkName: pkName = db.TableDefs(tblName).Fields(0)
    
    Dim ctl As control, fldArr As New clsArray, fldValArr As New clsArray
    
    Dim fldName, fldVal
    For Each ctl In frm.controls
        fldName = ctl.Name
        
        If ctl.Tag = "pk" Then
            If Not IsNull(ctl) Then
                UnboundFormUpdateRecord frm, tblName, fldName
                UnboundFormAddRecord = ctl
                Exit Function
            End If
            GoTo Next_ctl
        End If
        
        If DoesPropertyExists(rs.fields, fldName) Then
            fldArr.Add "[" & fldName & "]"
            fldVal = frm(fldName)
            fldValArr.Add EscapeString(fldVal, tblName, fldName)
        End If
        
Next_ctl:

    Next ctl
    
    RunSQL "INSERT INTO " & tblName & " (" & fldArr.JoinArr(",") & ") VALUES (" & fldValArr.JoinArr(",") & ")"
    UnboundFormAddRecord = CurrentDb.OpenRecordset("SELECT @@IDENTITY")(0)
    
End Function

Public Function UnboundFormUpdateRecord(frm, tblName, pkName)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM " & tblName)
    Dim pkVal: pkVal = frm(pkName)
  
    ''Dim pkName: pkName = db.TableDefs(tblName).Fields(0)
    
    Dim ctl As control, updateArr As New clsArray
    
    Dim fldName, fldVal
    For Each ctl In frm.controls
        fldName = ctl.Name
        
        If DoesPropertyExists(rs.fields, fldName) And fldName <> pkName Then
            fldVal = frm(fldName)
            updateArr.Add fldName & " = " & EscapeString(fldVal, tblName, fldName)
        End If
    
    Next ctl
    
    RunSQL "UPDATE " & tblName & " SET " & updateArr.JoinArr(",") & " WHERE " & pkName & " = " & pkVal
    UnboundFormUpdateRecord = pkVal
    
End Function

Public Function OpenUnboundForm(frm As Object, frmName, tblName, pkName, Optional RunOnOpen = "")

    Dim pkValue: pkValue = frm(pkName)
    If ExitIfTrue(isFalse(pkName), "Please select a valid record.") Then Exit Function
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM " & tblName & " WHERE " & pkName & " = " & pkValue)
    
    DoCmd.OpenForm frmName
    Set frm = Forms(frmName)
    
    Dim fld As field, fldName
    For Each fld In rs.fields
        fldName = fld.Name
        
        If DoesPropertyExists(frm, fld.Name) Then
            frm(fldName) = rs.fields(fldName)
            
        End If
        
    Next fld
    
    If RunOnOpen <> "" Then Run RunOnOpen, frm
    
End Function

Public Function OpenUnboundFormFromReport(rpt As Report, frmName, tblName, pkName, Optional RunOnOpen = "")

    Dim pkValue: pkValue = rpt(pkName)
    If ExitIfTrue(isFalse(pkName), "Please select a valid record.") Then Exit Function
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM " & tblName & " WHERE " & pkName & " = " & pkValue)
    
    DoCmd.OpenForm frmName
    Dim frm As Form: Set frm = Forms(frmName)
    
    Dim fld As field, fldName
    For Each fld In rs.fields
        fldName = fld.Name
        
        If DoesPropertyExists(frm, fld.Name) Then
            frm(fldName) = rs.fields(fldName)
            
        End If
        
    Next fld
    
    If RunOnOpen <> "" Then Run RunOnOpen, frm
    
End Function

