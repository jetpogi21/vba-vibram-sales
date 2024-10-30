VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'***************************************************************************************
' Module    : clsRidesPBar
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Copyright : Please note that U.O.S. all the content herein considered to be
'             intellectual property (copyrighted material).
'             It may not be copied, reused or modified in any way without prior
'             authorization from its author(s).
'***************************************************************************************
Private Const sModName = "clsRidesPBar"    'Application.VBE.ActiveCodePane.CodeModule
Dim frmProgress               As Form_frm_ProgressBar

Public Enum pTextAlign
    Align_General = 0      '(Default) The text aligns to the left; numbers and dates align to the right.
    Align_Left = 1     'The text, numbers, and dates align to the left.
    Align_Center = 2   'The text, numbers, and dates are centered.
    Align_Right = 3    'The text, numbers, and dates align to the right.
    Align_Distribute = 4   'The text, numbers, and dates are evenly distributed.
End Enum

Public Enum pFontWeight
    Weight_Thin = 100
    Weight_ExtraLight = 200
    Weight_Light = 300
    Weight_Normal = 400
    Weight_Medium = 500
    Weight_SemiBold = 600
    Weight_Bold = 700
    Weight_ExtraBold = 800
    Weight_Heavy = 900
End Enum

Public Function ProgressBar_Show()
    Set frmProgress = New Form_frm_ProgressBar

    frmProgress.Visible = True
    frmProgress.Section(acFooter).Visible = False
    DoEvents
End Function

Public Function ProgressBar_Hide()
    frmProgress.Visible = False
    Set frmProgress = Nothing
End Function

Public Function ProgressBar_Caption(sInput As String)
    frmProgress.Caption = sInput
End Function

Public Function ProgressBar_Color(lColor As Long)
    frmProgress.lbl_Progress.BackColor = lColor
    DoEvents
End Function

Public Function ProgressBar_Message1(sInput As String)
    frmProgress.lbl_Message1.Caption = sInput
    DoEvents
End Function

Public Function ProgressBar_Message1_Align(lTextAlign As pTextAlign)
    frmProgress.lbl_Message1.TextAlign = lTextAlign
End Function

Public Function ProgressBar_Message1_FontParam(Optional lTextAlign As pTextAlign = -1, Optional sFontName As String, Optional lFontWeight As pFontWeight, Optional iFontSize As Integer, Optional lFontColor As Long)
'Font Weight => https://learn.microsoft.com/en-us/office/vba/api/access.label.fontweight?f1url=%3FappId%3DDev11IDEF1%26l%3Den-US%26k%3Dk(vbaac10.chm10212)%3Bk(TargetFrameworkMoniker-Office.Version%3Dv15)%26rd%3Dtrue
'Font Size => https://learn.microsoft.com/en-us/office/vba/api/access.label.fontsize?f1url=%3FappId%3DDev11IDEF1%26l%3Den-US%26k%3Dk(vbaac10.chm10211)%3Bk(TargetFrameworkMoniker-Office.Version%3Dv15)%26rd%3Dtrue
'Font ForeColor => https://learn.microsoft.com/en-us/office/vba/api/access.label.forecolor?f1url=%3FappId%3DDev11IDEF1%26l%3Den-US%26k%3Dk(vbaac10.chm10209)%3Bk(TargetFrameworkMoniker-Office.Version%3Dv15)%26rd%3Dtrue
    If lTextAlign <> -1 Then frmProgress.lbl_Message1.TextAlign = lTextAlign
    If sFontName <> "" Then frmProgress.lbl_Message1.FontName = sFontName
    If lFontWeight <> 0 Then frmProgress.lbl_Message1.FontWeight = lFontWeight
    If iFontSize <> 0 Then frmProgress.lbl_Message1.fontSize = iFontSize
    If lFontColor <> 0 Then frmProgress.lbl_Message1.ForeColor = lFontColor
End Function

Public Function ProgressBar_Message2(sInput As String)
    frmProgress.lbl_Message2.Caption = sInput
    DoEvents
End Function

Public Function ProgressBar_Message2_Align(lTextAlign As pTextAlign)
    frmProgress.lbl_Message2.TextAlign = lTextAlign
End Function

Public Function ProgressBar_Message2_FontParam(Optional lTextAlign As pTextAlign = -1, Optional sFontName As String, Optional lFontWeight As pFontWeight, Optional iFontSize As Integer, Optional lFontColor As Long)
'Font Weight => https://learn.microsoft.com/en-us/office/vba/api/access.label.fontweight?f1url=%3FappId%3DDev11IDEF1%26l%3Den-US%26k%3Dk(vbaac10.chm10212)%3Bk(TargetFrameworkMoniker-Office.Version%3Dv15)%26rd%3Dtrue
'Font Size => https://learn.microsoft.com/en-us/office/vba/api/access.label.fontsize?f1url=%3FappId%3DDev11IDEF1%26l%3Den-US%26k%3Dk(vbaac10.chm10211)%3Bk(TargetFrameworkMoniker-Office.Version%3Dv15)%26rd%3Dtrue
'Font ForeColor => https://learn.microsoft.com/en-us/office/vba/api/access.label.forecolor?f1url=%3FappId%3DDev11IDEF1%26l%3Den-US%26k%3Dk(vbaac10.chm10209)%3Bk(TargetFrameworkMoniker-Office.Version%3Dv15)%26rd%3Dtrue
    If lTextAlign <> -1 Then frmProgress.lbl_Message2.TextAlign = lTextAlign
    If sFontName <> "" Then frmProgress.lbl_Message2.FontName = sFontName
    If lFontWeight <> 0 Then frmProgress.lbl_Message2.FontWeight = lFontWeight
    If iFontSize <> 0 Then frmProgress.lbl_Message2.fontSize = iFontSize
    If lFontColor <> 0 Then frmProgress.lbl_Message2.ForeColor = lFontColor
End Function

Public Function ProgressBar_Progress(lValue As Double)
'lValue = percentage => 0 to 1
    Dim iTotalWidth           As Long
    Dim iheight As Integer
    
    On Error Resume Next
    iTotalWidth = frmProgress.TotalProgress.Width - (lProgressBarOffset * 2)
    frmProgress.lbl_Progress.Width = (iTotalWidth * lValue)
    frmProgress.lbl_ProgressValue.Caption = VBA.FormatPercent(lValue, 2)
    
    If lValue = 1 Then
        iheight = frmProgress.InsideHeight
        frmProgress.Section(acFooter).Visible = True
        frmProgress.InsideHeight = iheight + frmProgress.Section(acFooter).Height
    End If
    
    DoEvents
End Function

Public Function ProgressBar_ProgressValue_Align(lTextAlign As pTextAlign)
    frmProgress.lbl_Progress.TextAlign = lTextAlign
    frmProgress.lbl_ProgressValue.TextAlign = lTextAlign
End Function

Public Function ProgressBar_ProgressOverlay(bOverlay As Boolean)
    If bOverlay = True Then
        frmProgress.lbl_ProgressValue.Top = frmProgress.TotalProgress.Top + ((frmProgress.TotalProgress.Height - frmProgress.lbl_ProgressValue.Height) / 2)
    Else
        frmProgress.lbl_ProgressValue.Top = frmProgress.TotalProgress.Top + frmProgress.TotalProgress.Height + 6
    End If
End Function

Public Function ProgressBar_Resize(sHeight As Single, sWidth As Single)
    Dim dHeightProportion     As Double
    Dim dWidthProportion      As Double
    Dim oCtrl                 As Access.control

    frmProgress.InsideHeight = sHeight
    frmProgress.InsideWidth = sWidth

    dHeightProportion = (sHeight / sHeight_Orig)
    dWidthProportion = (sWidth / sWidth_Orig)

    For Each oCtrl In frmProgress.controls
        oCtrl.Height = CInt(oCtrl.Height * dHeightProportion)
        oCtrl.Top = CInt(oCtrl.Top * dHeightProportion)
        oCtrl.Width = CInt(oCtrl.Width * dWidthProportion)
        oCtrl.Left = CInt(oCtrl.Left * dWidthProportion)
        oCtrl.TopPadding = CInt(oCtrl.TopPadding * dHeightProportion)

        If oCtrl.ControlType = acLabel And oCtrl.Visible = True Then
            oCtrl.fontSize = CInt(oCtrl.fontSize * dHeightProportion)
        End If
    Next
End Function



