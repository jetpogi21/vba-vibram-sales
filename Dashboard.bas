Attribute VB_Name = "Dashboard"
Option Compare Database
Option Explicit

Public systemAutomationUserID
Public secondsSinceLastRun
Public secondsSinceUpdateChecked

Public Function FilterMenu(frm As Form)

    If isFalse(g_userID) Then
        Exit Function
    End If
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("select * from qryUserUserGroups where UserID = " & g_userID & " And UserGroup = 'Administrator'")

    If Not rs.EOF Then
        Dim ctl As control
        For Each ctl In frm.controls
            If ctl.ControlType = acCommandButton Then
                ctl.Visible = True
            End If
        Next ctl
    Else
        Dim sqlStr As String
        frm.recordSource = "select * from tblMainMenus where MainMenuID IN(select MainMenuID from qryUserGroupButtons where UserID = " & g_userID & " GROUP BY MainMenuID)"
    End If
    
End Function

Private Function NewVersionRequired() As Boolean

    Dim currGlobalVersion As Variant

    currGlobalVersion = GetGlobalSetting("Application_FrontEndVersion")

    If g_FrontEndVersion = currGlobalVersion Then
        NewVersionRequired = False
    Else
        NewVersionRequired = True
    End If
    
End Function

'Public Function DashboardLoad(frm As Form)
'
'    If isFalse(g_UserID) Then Exit Function
'
'    Dim UserID, sqlObj As New clsSQL
'    UserID = g_UserID
'
'    Dim rs As Recordset, rs2 As Recordset
'    ''SELECT STATEMENT
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .Source = "tblUsers"
'        .AddFilter "UserID = " & g_UserID
'        Set rs = .Recordset
'    End With
'
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .Source = "qryUserUserGroups"
'        .AddFilter "UserID = " & g_UserID
'        Set rs2 = .Recordset
'    End With
'
'    Dim lblString As String
'    If Not rs.EOF Then
'        lblString = rs.Fields("UserName") & " | "
'        Do Until rs2.EOF
'            lblString = lblString & rs2.Fields("UserGroup") & " | "
'            rs2.MoveNext
'        Loop
'    End If
'
'    frm.lblLoginInfo.Caption = lblString & " Active Since: " & Now()
'
'End Function

Public Function DashboardOnOpen()
    If isFalse(g_userID) Then
        DoCmd.OpenForm "frmLogin"
        DoCmd.CancelEvent
    End If
End Function

Public Function DashboardOnUnload()
    If MsgBox("Are you sure you want to log off?", vbCritical + vbYesNo, "Log-off Prompt") = vbYes Then
    
        Dim frm As Form
        For Each frm In Application.Forms
            If frm.Name <> "frmDashboard" And CurrentProject.AllForms(frm.Name).IsLoaded Then
                DoCmd.Close acForm, frm.Name, acSaveNo
            End If
        Next frm
        
        g_userID = Null
        DoCmd.OpenForm "frmLogin"
    Else
        DoCmd.CancelEvent
    End If
End Function
