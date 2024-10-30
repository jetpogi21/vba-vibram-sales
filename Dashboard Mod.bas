Attribute VB_Name = "Dashboard Mod"
Option Compare Database
Option Explicit
Public IsImported As Boolean

Public Function ChangeProjectAndClientPathsBasedOnComputerName()
    
    ''C:\Users\Jet\OneDrive\Desktop\Web Development\vibram-sales\
    ''C:\Users\Clarisse\Desktop\Web Development\vibram-sales
    ''Replace the Jet\OneDrive with Clarisse
    
    Dim ReplaceString
    
    If Environ("ComputerName") = "DESKTOP-UOL72RE" Then
        ReplaceString = "Replace([ProjectPath],""Jet\OneDrive"",""Clarisse"")"
    Else
        ReplaceString = "Replace([ProjectPath],""Clarisse"",""Jet\OneDrive"")"
    End If
    
    RunSQL "UPDATE tblBackendProjects SET ProjectPath = " & ReplaceString & ",ClientPath = " & ReplaceString
    
    ReplaceString = replace(ReplaceString, "ProjectPath", "FilePath")
    RunSQL "UPDATE tblBackendProjectFiles SET FilePath = " & ReplaceString
    RunSQL "UPDATE tblSeqModelFiles SET FilePath = " & ReplaceString
    
End Function

Public Function QuitApp()
    DoCmd.Quit
End Function

Public Sub frmCustomDashboard_SelectAllButtons()

    DoCmd.OpenForm "frmCustomDashboard", acDesign
    Dim frm As Form: Set frm = Forms("frmCustomDashboard")
    
    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.ControlType = acRectangle Or ctl.ControlType = acTextBox Then
            ctl.InSelection = True
        End If
    Next ctl
End Sub

Public Function LogOutUser()
    If MsgBox("Are you sure you want to log off?", vbCritical + vbYesNo, "Log-off Prompt") = vbYes Then
    
        Dim frm As Form
        For Each frm In Application.Forms
            If frm.Name <> "frmCustomDashboard" And CurrentProject.AllForms(frm.Name).IsLoaded Then
                DoCmd.Close acForm, frm.Name, acSaveNo
            End If
        Next frm
        
        g_userID = Null
        
        ''BackupBackendFile
        DoCmd.Quit
        ''DoCmd.OpenForm "frmLogin"
        ''DoCmd.Close acForm, "frmCustomDashboard", acSaveNo
    Else
        DoCmd.CancelEvent
    End If
End Function

Public Sub SetUp_frmCustomDashboard(Optional frmName = "frmCustomDashboard", Optional PreventReposition As Boolean = False)
    
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    
    frm.Moveable = True
    frm.BorderStyle = 1
    ''Center the controls
    Dim frmWidth: frmWidth = frm.Width
    
    ''Set the widths of all labels to the frmWidth
    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.ControlType = acLabel Then
            ctl.Width = frmWidth
            ctl.Left = 0
            ctl.BackStyle = 0
        End If
    Next ctl
    
    ''Set the names of the cmdButtons to sequential cmdi
    Dim j: j = 211
    Dim i: i = 13
    For j = 249 To 259
        
        frm("Command" & j).Name = "cmd" & i
        i = i + 1
        
    Next j
    
    'Next are the buttons. center them
    If Not PreventReposition Then
        For Each ctl In frm.controls
            If ctl.ControlType = acCommandButton Or ctl.ControlType = acImage Then
                ctl.Left = GetLeftPosition(frmWidth, ctl.Width)
            End If
        Next ctl
    End If
    
    Dim filterStr: filterStr = "ParentMenu IS NULL"
    If frmName = "frmSetupDashboard" Then
        filterStr = "ParentMenu = ""Setup"""
    End If
    
    If frmName = "frmReportDashboard" Then
        filterStr = "ParentMenu = ""Report"""
    End If
    
    If frmName = "frmNonAdminDashboard" Then
        filterStr = "ParentMenu = ""NonAdmin"""
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblMainMenus WHERE " & filterStr & " ORDER BY MenuOrder"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    i = 0
    Do Until rs.EOF
        Dim MenuCaption: MenuCaption = rs.fields("MenuCaption")
        Dim Translation: Translation = rs.fields("Translation")
        Dim Icon: Icon = rs.fields("Icon")
        Dim ButtonName: ButtonName = "cmd" & i
        Set ctl = frm(ButtonName)
        ctl.BackStyle = 0
        ctl.Caption = ""
        ''Create the image control for the icon
        Dim imageCtlName: imageCtlName = "img" & ctl.Name
        Dim imageCtl As image
        If DoesPropertyExists(frm, imageCtlName) Then
            Set imageCtl = frm(imageCtlName)
        Else
            Set imageCtl = CreateControl(frm.Name, acImage)
            imageCtl.Name = imageCtlName
        End If
        ''Set the position and dimensions
        imageCtl.Top = ctl.Top + InchToTwip(0.02)
        imageCtl.Left = ctl.Left + InchToTwip(0.09)
        imageCtl.Width = InchToTwip(0.2)
        imageCtl.Height = InchToTwip(0.23)
        imageCtl.Picture = Coalesce(Icon, "")
            ''Set the position color and size
        ''Create the textbox for the actual menu text
        Dim textboxCtlName: textboxCtlName = "txt" & ctl.Name
        Dim textboxCtl As TextBox
        If DoesPropertyExists(frm, textboxCtlName) Then
            Set textboxCtl = frm(textboxCtlName)
        Else
            Set textboxCtl = CreateControl(frm.Name, acTextBox)
            textboxCtl.Name = textboxCtlName
        End If
        ''Set the position and dimensions
        textboxCtl.BackStyle = 1
        textboxCtl.BackColor = 12928318
        textboxCtl.BorderStyle = 0
        
        Dim ButtonCaption: ButtonCaption = Coalesce(Translation, MenuCaption)
        textboxCtl.ControlSource = "=" & Esc(ButtonCaption)
        
        Dim TemplateControlName: TemplateControlName = "ReverseTextControl"
        If frmName = "frmReportDashboard" Then
            TemplateControlName = "ReverseTextControl1"
        End If
        
        If frmName = "frmSetupDashboard" Then
            TemplateControlName = "ReverseTextControl2"
        End If
        
        If ButtonCaption = "Switchboard" Then
            TemplateControlName = "ReverseTextControl3"
        End If
        
        CopyProperties frm, textboxCtlName, TemplateControlName
        
        textboxCtl.BorderColor = textboxCtl.BackColor
        ctl.BorderColor = textboxCtl.BackColor
        textboxCtl.TextAlign = 1
        textboxCtl.Top = ctl.Top
        textboxCtl.Left = ctl.Left
        textboxCtl.Width = ctl.Width
        textboxCtl.Height = ctl.Height
        textboxCtl.TabStop = False
        textboxCtl.TopMargin = 75
        textboxCtl.LeftMargin = 600
        CenterVertically frm.Name, ButtonName, textboxCtlName
        textboxCtl.Enabled = False
        textboxCtl.Locked = True
        
        textboxCtl.InSelection = True
        DoCmd.RunCommand acCmdSendToBack
        
        i = i + 1
        rs.MoveNext
    Loop
    
    ''Remove extra widths
    frm.Width = 0
    
End Sub

Public Sub SyncSetUpFormsWithControlTemplate()

    ''except the frmCustomDashboard
    Dim rs As Recordset: Set rs = ReturnRecordset("Select * from tblMainMenus where ParentMenu = ""Setup"" ANd FormName <> ""frmCustomDashboard""" & _
        " ORDER BY MenuOrder")
    Dim frm As Form
    Do Until rs.EOF
        Dim FormName: FormName = rs.fields("FormName")
        SyncControlTemplateOfForm FormName, True
        FormName = replace(FormName, "main", "frm")
        SyncControlTemplateOfForm FormName, True
        rs.MoveNext
    Loop
    
End Sub

Private Sub frmCustomDashboard_HideUnusedButtons(frm As Form)
    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.ControlType = acCommandButton Then
            If isFalse(ctl.Caption) Then
                ctl.Visible = False
            End If
        End If
    Next ctl
End Sub

Public Function frmCustomdashboard_SyncMenu(frm As Object, Optional ParentMenu = "")
    
    Dim filterStr: filterStr = IIf(isFalse(ParentMenu), "ParentMenu IS NULL", "ParentMenu = " & Esc(ParentMenu))
    Dim rs As Recordset: Set rs = ReturnRecordset("Select * from tblMainMenus WHERE " & filterStr & " ORDER BY MenuOrder")
    Dim i, ctl As CommandButton: i = 0
    Do Until rs.EOF
        Dim MenuCaption: MenuCaption = rs.fields("MenuCaption")
        Dim Translation: Translation = rs.fields("Translation")
        Dim FormName: FormName = rs.fields("FormName")
        Set ctl = frm("cmd" & i)
        ctl.Caption = IIf(Not isFalse(Translation), Translation, MenuCaption)
        ctl.OnClick = "=DoOpenForm(" & EscapeString(FormName) & ",Null,False)"
        i = i + 1
        rs.MoveNext
    Loop
    
    frmCustomDashboard_HideUnusedButtons frm
    
End Function

Public Function DashboardLoad(frm As Object, Optional ParentMenu = "")
    
    ''IsExpired DateSerial(2024, 7, 20)
    
    If isFalse(g_userID) Then
    
        ''If Environ("computername") <> "LAPTOP-4EL19IO4" Then Exit Function
        g_userID = 1
        ''frm.lblLoginInfo.Caption = ""
        
    End If
    
    ''If Environ("computername") <> "LAPTOP-4EL19IO4" Then RunOneTimeFixes
    ''RunOneTimeFixes
    
'    ''TABLE: qryInternalLeads Fields: InternalLeadID|InternalLeadName|Timestamp|CreatedBy|RecordImportID|BusinessUnit|InternalLeadFullName
'    Dim rs As Recordset
'    Set rs = ReturnRecordset("Select InternalLeadFullName FROM qryInternalLeads WHERE InternalLeadID = " & g_userID)
'    If rs.EOF Then
'        frm.lblLoginInfo.Caption = ""
'        Exit Function
'    End If
    
'    frm.lblLoginInfo.Caption = "Logged in as: " & rs.fields("InternalLeadFullName")
    ''frm.lblLoginInfo.Caption = lblString & " Active Since: " & Now()
    
    ChangeProjectAndClientPathsBasedOnComputerName
    FilterDashboardMenu frm
    ''frmCustomdashboard_SyncMenu frm, ParentMenu
    ''CloseAllOtherSwitchboards ParentMenu
    ''TranslateToArabic frm
    'NoHasWriteToFilePrompt = False
    
    ''If Environ("computername") <> "LAPTOP-4EL19IO4" Then Shell ("OUTLOOK")
    
    ''FilterDashboardReports frm

End Function


Private Sub CloseAllOtherSwitchboards(Optional ParentMenu = "")
    
    Dim frm As Form
    If ParentMenu <> "" Then
        Set frm = GetForm("frmCustomDashboard")
        DoCmd.Close acForm, frm.Name, acSaveNo
    End If
    
    Dim filterStr: filterStr = "NOT ParentMenu IS NULL"
    If ParentMenu <> "" Then
        filterStr = filterStr & " AND ParentMenu <> " & Esc(ParentMenu)
    End If
    
    Dim sqlStr: sqlStr = "SELECT ParentMenu FROM tblMainMenus WHERE " & filterStr & " GROUP BY ParentMenu ORDER BY ParentMenu"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        ParentMenu = rs.fields("ParentMenu")
        Set frm = GetForm("frm" & ParentMenu & "Dashboard")
        If Not frm Is Nothing Then
            DoCmd.Close acForm, frm.Name, acSaveNo
        End If
        rs.MoveNext
    Loop
    
End Sub

Public Function IsExpired(targetDate As Date) As Boolean
    ' Get the current date
    Dim currentDate As Date
    currentDate = Date

    ' Compare the current date to the target date
    If currentDate > targetDate Then
        IsExpired = True
        DoCmd.Quit
    End If
    
End Function

Private Sub UpdateEntitiesIsSeller()

    RunSQL "UPDATE tblEntities SET IsSeller = -1 WHERE EntityCategoryID = 2"
    
End Sub

Private Sub DeleteTable(LinkedTableName)

On Error GoTo ErrHandler:

    DoCmd.DeleteObject acTable, LinkedTableName
    
ErrHandler:
    Exit Sub
    
End Sub

Public Function LinkTheTables()
    
    Dim ProjectPath, filePath
    
    ProjectPath = CurrentProject.path
    
    If Environ("computername") <> "LAPTOP-4EL19IO4" Then
        ProjectPath = "Z:\MY PANDA APP"
        If Not DirectoryExists(ProjectPath) Then
            ProjectPath = "\\TRUENAS\database\MY PANDA APP"
            If Not DirectoryExists(ProjectPath) Then
                MsgBox "The database tables can't be linked to the backend file. The app will exit.", vbCritical
                DoCmd.Quit
                Exit Function
            End If
        End If
    End If
    filePath = ProjectPath & "\PTS Backend.accdb"
    
    ''AlterBackendTable filePath
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblLinkedTables ORDER BY Timestamp ASC")
    
    Do Until rs.EOF
        
        DeleteTable rs.fields("LinkedTableName")
            
        DoCmd.TransferDatabase TransferType:=acLink, _
            DatabaseType:="Microsoft Access", _
            DatabaseName:=filePath, _
            ObjectType:=acTable, _
            Source:=rs.fields("LinkedTableName"), _
            Destination:=rs.fields("LinkedTableName")
        
        rs.MoveNext
        
    Loop
    
    IsImported = True
    
End Function

Private Sub FilterDashboardMenu(frm As Form)

    Dim rs As Recordset
    
    If isPresent("qryUserUserGroups", "UserID = " & g_userID & " And UserGroup = " & EscapeString("Administrator")) Then
        Set rs = ReturnRecordset("SELECT * FROM tblMainMenus ORDER BY MenuOrder")
    Else
        Set rs = ReturnRecordset("select * from tblMainMenus where MainMenuID IN(select MainMenuID from qryUserGroupButtons where UserID = " & g_userID & " GROUP BY MainMenuID) ORDER BY MenuOrder")
    End If
    
    Dim i
    For i = 0 To 17
    
        If rs.EOF Then
            frm("cmd" & i).Visible = False
        Else
            frm("cmd" & i).Visible = True
            frm("cmd" & i).Transparent = False
            frm("cmd" & i).Caption = rs.fields("MenuCaption")
            frm("cmd" & i).OnClick = "=DoOpenForm(" & EscapeString(rs.fields("FormName")) & ",Null,False)"
            
            ''Override the "Web Search" Menu Caption
            If frm("cmd" & i).Caption = "Web Search" Then
                frm("cmd" & i).OnClick = "=OpenGoogle()"
            End If
            
            rs.MoveNext
        End If
    
    Next i

End Sub

Private Sub FilterDashboardReports(frm As Form)

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("Select * From tblCustomReports ORDER BY ReportName ASC")

    Dim i
    For i = 0 To 9
    
        If rs.EOF Then
            frm("rpt" & i).Visible = False
        Else
            frm("rpt" & i).Visible = True
            frm("rpt" & i).Caption = rs.fields("ReportName")
            frm("rpt" & i).OnClick = "=OpenCustomReportForm(" & EscapeString(rs.fields("FilterFormName")) & ")"
            rs.MoveNext
        End If
    
    Next i

End Sub


'Public Sub CreateCustomDashboard()
'
'    ''Initialize x and y axis
'    x = 200: y = 200
'    Set frm = CreateForm
'    SetFormProperties
'    InsertLogo
'    InsertDashboardHeader
'    InsertDashboardSubHeader
'    InsertDashboardMenu
'
'End Sub
'
'
'Private Sub SetFormProperties()
'
'    With frm
'        .RecordSelectors = 0
'        .CloseButton = 0
'        .NavigationButtons = 0
'        .ScrollBars = 0
'        .Caption = "Main Menu"
'        .Picture = "BackgroundCropped"
'        .PictureSizeMode = 1
'    End With
'
'End Sub
'
'Private Sub InsertLogo()
'
'    Dim ctl As Control
'    Set ctl = CreateControl(frm.Name, acImage, , , , x, y, 5000, 2500)
'    ctl.Name = "imgLogo"
'    ctl.Picture = "Logo"
'
'End Sub
'
'
'Private Sub InsertDashboardHeader()
'
'    x = frm("imgLogo").Left + frm("imgLogo").Width + 200
'    Dim ctl As Control
'    Set ctl = CreateControl(frm.Name, acLabel, , , , x, y, 9000, 660)
'    ctl.Name = "lblHeader"
'    ctl.Caption = "EPC INVENTORY AND ASSET TRACKING"
'    ctl.FontName = "Segoe UI Black"
'    ctl.ForeColor = RGB(254, 254, 254)
'    ctl.FontSize = 22
'
'End Sub
'
'Private Sub InsertDashboardSubHeader()
'
'    x = frm("imgLogo").Left + frm("imgLogo").Width + 200
'    y = frm("lblHeader").Top + frm("lblHeader").height
'    Dim ctl As Control
'    Set ctl = CreateControl(frm.Name, acLabel, , , , x, y, frm("lblHeader").Width, 450)
'    ctl.Caption = "Welcome. To begin, select an option below."
'    ctl.FontName = "Segoe UI Black"
'    ctl.ForeColor = RGB(254, 254, 254)
'    ctl.FontSize = 14
'
'End Sub
'
'Private Sub InsertDashboardMenu()
'
'    Dim proportionArr As New clsArray, controlArr As New clsArray, proportionTotal, totalWidth, colSpaceWidth, i, proportion
'
'    y = frm("imgLogo").Top + frm("imgLogo").height + 500
'    x = frm("imgLogo").Left
'
'    colSpaceWidth = 200
'    totalWidth = 7000
'
'    Dim ctl As Control
'    For i = 0 To 11
'         Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0) ''Button Portion
'         ctl.Name = "cmd" & i
'         SetControlPropertiesFromTemplate ctl, frm
'    Next i
'
'    ''Render the Filter buttons
'    ''Filter and Clear
'    proportionArr.Arr = "4,4,4"
'    controlArr.Arr = "cmd0,cmd1,cmd2"
'    proportionTotal = GetProportionTotal(proportionArr)
'
'    For i = 0 To proportionArr.Count - 1
'
'        proportion = CDbl(proportionArr.Arr(i)) / proportionTotal
'        frm(controlArr.Arr(i)).Left = x
'        frm(controlArr.Arr(i)).Top = y
'        frm(controlArr.Arr(i)).Width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
'        frm(controlArr.Arr(i)).HorizontalAnchor = acHorizontalAnchorLeft
'
'        x = x + (colSpaceWidth * 2) + frm(controlArr.Arr(i)).Width
'
'    Next i
'
'    y = y + frm("cmd0").height + 200
'    x = frm("imgLogo").Left
'
'    controlArr.Arr = "cmd3,cmd4,cmd5"
'
'    For i = 0 To proportionArr.Count - 1
'
'        proportion = CDbl(proportionArr.Arr(i)) / proportionTotal
'        frm(controlArr.Arr(i)).Left = x
'        frm(controlArr.Arr(i)).Top = y
'        frm(controlArr.Arr(i)).Width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
'        frm(controlArr.Arr(i)).HorizontalAnchor = acHorizontalAnchorLeft
'
'        x = x + (colSpaceWidth * 2) + frm(controlArr.Arr(i)).Width
'
'    Next i
'
'    y = y + frm("cmd3").height + 200
'    x = frm("imgLogo").Left
'
'    controlArr.Arr = "cmd6,cmd7,cmd8"
'
'    For i = 0 To proportionArr.Count - 1
'
'        proportion = CDbl(proportionArr.Arr(i)) / proportionTotal
'        frm(controlArr.Arr(i)).Left = x
'        frm(controlArr.Arr(i)).Top = y
'        frm(controlArr.Arr(i)).Width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
'        frm(controlArr.Arr(i)).HorizontalAnchor = acHorizontalAnchorLeft
'
'        x = x + (colSpaceWidth * 2) + frm(controlArr.Arr(i)).Width
'
'    Next i
'
'    y = y + frm("cmd6").height + 200
'    x = frm("imgLogo").Left
'
'    controlArr.Arr = "cmd9,cmd10,cmd11"
'
'    For i = 0 To proportionArr.Count - 1
'
'        proportion = CDbl(proportionArr.Arr(i)) / proportionTotal
'        frm(controlArr.Arr(i)).Left = x
'        frm(controlArr.Arr(i)).Top = y
'        frm(controlArr.Arr(i)).Width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
'        frm(controlArr.Arr(i)).HorizontalAnchor = acHorizontalAnchorLeft
'
'        x = x + (colSpaceWidth * 2) + frm(controlArr.Arr(i)).Width
'
'    Next i
'
'
'End Sub








    


