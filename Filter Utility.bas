Attribute VB_Name = "Filter Utility"
Option Compare Database
Option Explicit
Public FilterCaption As String

Public Function FilterReportBetween(frm As Object, rptName, dateField, dateCaption, recordsetName)
    
    Dim StartDate, EndDate
    StartDate = frm("startDate")
    EndDate = frm("endDate")
    
    If ExitIfTrue(IsNull(StartDate), "Please supply a start date..") Then Exit Function
    If ExitIfTrue(IsNull(EndDate), "Please supply an end date..") Then Exit Function
    
    FilterCaption = dateCaption & " From " & StartDate & " and " & EndDate
    
    If ExitIfTrue(ECount(recordsetName, dateField & " BETWEEN #" & StartDate & "# And #" & EndDate & "#") = 0, "There is no record to show..") Then Exit Function
    
    DoCmd.OpenReport rptName, acViewPreview, , dateField & " BETWEEN #" & StartDate & "# And #" & EndDate & "#"

End Function

Public Function FilterDataEntryForm(frmName As String, fieldName As String, FilterControlName As String)
    
    Dim frm As Form: Set frm = Forms(frmName)
    
    Dim filterValue: filterValue = frm(FilterControlName)
    
    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.Name Like "fltr*" And ctl.Name <> FilterControlName Then
            ctl = Null
        End If
    Next ctl
    
    OpenFormByWhereClause frmName, fieldName & " = " & Esc(filterValue)
    
End Function
