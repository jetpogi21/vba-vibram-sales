Attribute VB_Name = "RethemeForm Mod"
Option Compare Database
Option Explicit

Public Function RethemeFormCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function RethemeBulkForms()
    
    Dim frmStrs As New clsArray: frmStrs.arr = "Companies,InvestmentArea,CompanyIAs,InternalLead,Oppurtunities,CompanyIAChoices,ContactNotes,Pipelines"
    Dim frmStr
    
    For Each frmStr In frmStrs.arr
        RethemeForm "main" & frmStr
        RethemeForm "frm" & frmStr
    Next frmStr
    
End Function

Private Function TryToOpenForm(frmName) As Boolean
    
    On Error GoTo Err_Handler:
    
    DoCmd.OpenForm frmName, acDesign
    
    TryToOpenForm = True
  
Err_Handler:

    Exit Function
    
End Function

Public Function RethemeForm(frmName)

    If Not TryToOpenForm(frmName) Then Exit Function
    
    ''Dark Blue
    Dim bgColor: bgColor = 9917743
    
    Dim frm As Form: Set frm = Forms(frmName)
    
    frm.Section(acDetail).BackColor = bgColor
    
    Dim ctl As control
    
    For Each ctl In frm.controls
        
        Select Case ctl.ControlType
            
            Case acCommandButton, acTabCtl:
                
                FormatControls ctl
                
        End Select
        
    Next ctl
    
    DoCmd.Close acForm, frmName, acSaveYes
    
End Function

Private Function FormatControls(ctl As control)

    Dim ctlType: ctlType = ctl.ControlType
    ''TABLE: tblControlProps Fields: ControlPropID|ControlPropValue|ControlTypeID|ControlProp|Timestamp|CreatedBy
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM qryControlTypes WHERE ControlTypeValue = " & ctlType)
    Dim ControlPropID, ControlPropValue, ControlTypeID, ControlProp
    
    Do Until rs.EOF
        
        ControlPropID = rs.fields("ControlPropID")
        ControlPropValue = rs.fields("ControlPropValue")
        ControlTypeID = rs.fields("ControlTypeID")
        ControlProp = rs.fields("ControlProp")
        
        ctl.Properties(ControlPropValue) = ControlProp
        
        rs.MoveNext
    Loop
    
End Function

