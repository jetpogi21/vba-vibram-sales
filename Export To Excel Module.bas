Attribute VB_Name = "Export To Excel Module"
Option Compare Database
Option Explicit

Private Function BuildFileName(fileName) As String

    ''Default directory is on backend directory + Property Reports Folder
    Dim xlDirectory
    xlDirectory = CurrentProject.path & "\Files\"
    If Environ("computername") <> "LAPTOP-4EL19IO4" Then xlDirectory = "Z:\MY PANDA APP\Property Reports\"
    
    ''Check if the directory is existing, if not then create it.
    Dim strFolderExists
    strFolderExists = Dir(xlDirectory, vbDirectory)
    If strFolderExists = "" Then MkDir xlDirectory
    
    ''Build the search criteria
    ''Change File Name
    BuildFileName = xlDirectory & fileName & ".xlsx"
    
End Function

Private Function GetfieldsToExport(frm As Form)
    
    Dim frm2
    Set frm2 = frm("subform").Form
    
    Dim exceptedArr As New clsArray, fieldArr As New clsArray
    exceptedArr.arr = "IsFavorite,txtOpenInRPP,ExcludeFromReport"
    Dim ctl As control
    For Each ctl In frm2.controls
        ''Exlude Hidden columns, In excepted array and begins with sum
        If Not ctl.ColumnHidden And Not exceptedArr.InArray(ctl.Name) And Not ctl.Name Like "Sum*" Then
            fieldArr.Add ctl.Name
        End If
    Next ctl
    
    GetfieldsToExport = fieldArr.JoinArr(",")
    
End Function

Private Function GetSavePath(fileName)
    
    fileName = BuildFileName(fileName)
    Dim path: path = fileName
    
    Dim fd As FileDialog
    Set fd = FileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Choose a Location and Name of the File to Save This File"
        .ButtonName = "Click to Save"
        .InitialFileName = path
        If .Show <> 0 Then
            GetSavePath = .SelectedItems(1)
        End If
    End With
    
End Function


Public Function ExportToExcel(CustomReportID, frm As Form, dateCaption, dateField)

    Dim StartDate, EndDate
    StartDate = frm("startDate")
    EndDate = frm("endDate")
    
    If ExitIfTrue(IsNull(StartDate), "Please supply a start date..") Then Exit Function
    If ExitIfTrue(IsNull(EndDate), "Please supply an end date..") Then Exit Function
    
    FilterCaption = dateCaption & " From " & StartDate & " and " & EndDate

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReports WHERE CustomReportID = " & CustomReportID)
    
    Dim ReportName, ReportObjectName, FilterFormName, recordsetName, PreAppliedFilter, OrderBy, ReportOrientation, PaperSize
    ReportName = rs.fields("ReportName")
    recordsetName = rs.fields("RecordsetName")
    OrderBy = rs.fields("OrderBy")

    rs.Close
    
    If ExitIfTrue(ECount(recordsetName, dateField & " BETWEEN #" & StartDate & "# And #" & EndDate & "#") = 0, "There is no record to show..") Then Exit Function
    
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReportFields WHERE CustomReportID = " & CustomReportID & " And FieldOrder <> 0 ORDER BY FieldOrder ASC")
    
    Dim fieldNames As New clsArray, fieldCaptions As New clsArray, fieldTypes As New clsArray
    Dim CustomReportField, FieldTypeID, VerboseName
    
    Do Until rs.EOF
        
        CustomReportField = rs.fields("CustomReportField")
        FieldTypeID = rs.fields("FieldTypeID")
        VerboseName = rs.fields("VerboseName")
   
        fieldNames.Add CustomReportField
        fieldCaptions.Add VerboseName
        fieldTypes.Add FieldTypeID
        
        rs.MoveNext
        
    Loop
    
    rs.Close
    
    Dim xl As Object
    Dim sht As Object
    Dim xb As Object
    
    Set xl = CreateObject("Excel.Application")
    xl.Visible = True
    
    Set xb = xl.Workbooks.Add
    xb.Activate
    
    Set sht = xb.ActiveSheet
    
    sht.Cells(1, 1) = ReportName
    sht.Cells(2, 1) = FilterCaption
    
    Dim sqlStr
    sqlStr = "SELECT * FROM " & recordsetName & " WHERE " & dateField & " BETWEEN #" & StartDate & "# And #" & EndDate & "#"
    If Not IsNull(OrderBy) Then
        sqlStr = sqlStr & " ORDER BY " & OrderBy
    End If
    
    Set rs = ReturnRecordset(sqlStr)
    rs.MoveFirst
    
    Dim i As Integer, maxI As Integer, currentRow As Integer
    currentRow = 4
    maxI = fieldNames.count - 1
    
    For i = 0 To maxI
        sht.Cells(currentRow, i + 1) = fieldCaptions.arr(i)
        
        If fieldTypes.arr(i) = 8 Then
            sht.Columns(i + 1).NumberFormat = "m/d/yyyy"
        ElseIf fieldTypes.arr(i) = 7 Then
            sht.Columns(i + 1).NumberFormat = "#,##0.00"
        End If
        
    Next i
    currentRow = currentRow + 1
        
    Do Until rs.EOF
        For i = 0 To maxI
            sht.Cells(currentRow, i + 1) = rs.fields(fieldNames.arr(i))
        Next i
        currentRow = currentRow + 1
        rs.MoveNext
    Loop
    
    For i = 0 To maxI
    
        sht.Columns(i + 1).AutoFit
    
    Next i
    
End Function
