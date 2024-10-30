Attribute VB_Name = "Utility"
Option Compare Database
Option Explicit

Public ProgressPopup_StartTime As Long
Public ProgressPopup_LastProgress As Double

Dim DebugMode

Public Function ConvertToPascalCase(Text) As String
    
    Dim NewText: NewText = Text
    Dim matches As New clsArray: Set matches = GetAllMatchedGroup(NewText, "_\w")
    
    If matches.count = 0 Then
        ConvertToPascalCase = Capitalize(Text)
        Exit Function
    End If
    
    Dim item
    Dim i As Integer: i = 0
    For Each item In matches.arr
        ''Get the letter and convert to uppercase
        NewText = replace(NewText, item, UCase(Right(item, 1)))
        
        i = i + 1
    Next item
    
    NewText = Capitalize(NewText)
    ConvertToPascalCase = NewText

End Function

Public Function GoToLink(Link)

    CreateObject("Shell.Application").Open Link
    
End Function

Public Function SortStringsByLengthDescending(arr As Variant) As Variant
    Dim i As Long, j As Long
    Dim temp As String
    
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If Len(arr(i)) < Len(arr(j)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    SortStringsByLengthDescending = arr
End Function

Function Capitalize(inputString, Optional reverse As Boolean = False) As String
    If Len(inputString) > 0 Then
        Capitalize = IIf(reverse, LCase(Left(inputString, 1)), UCase(Left(inputString, 1))) & Right(inputString, Len(inputString) - 1)
    Else
        Capitalize = inputString
    End If
End Function

Public Function RemoveNewLines(Text) As String

    RemoveNewLines = replace(Text, vbNewLine, "")
    
End Function


Public Function GetSampleFieldValue(tblName, fieldName, Optional recordCount = 5, Optional DebugMode = True) As clsArray
    
    Dim values As New clsArray
    values.Add tblName
    values.Add fieldName
    
    Dim rs As Recordset
    
    Set rs = CurrentDb.OpenRecordset(GetReplacedString("SELECT [fieldName] FROM [tblName] GROUP BY [fieldName] ORDER BY [fieldName]", "tblName,fieldName", values))
    
    Dim sampleValues As New clsArray:
    
    Dim i As Integer: i = 0
    Do Until rs.EOF
        If DebugMode Then
            Debug.Print rs.fields(fieldName)
        End If
        sampleValues.Add rs.fields(fieldName)
        i = i + 1
        If i = recordCount Then
            GoTo ExitLoop
        End If
        rs.MoveNext
    Loop
ExitLoop:

    Set GetSampleFieldValue = sampleValues
    Exit Function
    
    
End Function


Public Function GetReplacedString(txtStr, wordsToReplacedWith, values As clsArray) As String
    
    Dim wordsArr As New clsArray: wordsArr.arr = wordsToReplacedWith
    
    Dim item
    Dim i As Integer: i = 0
    For Each item In wordsArr.arr
        txtStr = replace(txtStr, "[" & item & "]", values.arr(i))
        i = i + 1
    Next item
    
    GetReplacedString = txtStr
    
End Function


Public Function GetTranslation(English)

    GetTranslation = ELookup("tblTranslations", "English = " & Esc(English), "Arabic")
    
End Function

Public Function GetDatasheetCaption(ctl As control)

    GetDatasheetCaption = ctl.Properties("DatasheetCaption")
    
End Function

Public Function SetDatasheetCaption2(ctl As control, DatasheetCaption)

    ctl.Properties("DatasheetCaption") = DatasheetCaption
    
End Function

Public Function GetIsAdmin() As Boolean
    
    If isFalse(g_userID) Then Exit Function
    
    Dim IsAdmin: IsAdmin = ELookup("tblUsers", "UserID = " & g_userID, "IsAdmin")
    
    If isFalse(IsAdmin) Then Exit Function
    
    GetIsAdmin = CBool(IsAdmin)
    
End Function

Public Function GetBottom(ctl As control)
    
    GetBottom = ctl.Top + ctl.Height
    
End Function

Public Function GetRight(ctl As control)
    
    GetRight = ctl.Left + ctl.Width
    
End Function

Public Sub e(str As String)

    CopyToClipboard replace("Dim ModelID: ModelID = ELookup(""tblModels"",""Model = "" & Esc(""Model""),""ModelID"")", "Model", str)
    
End Sub

Public Function GetFormattedTimestamp()

    Dim formattedTimestamp As String
    GetFormattedTimestamp = Format$(Now, "YYYYMMDDhhmmss")
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : DeleteTableRelationships
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Delete all the relationships for the specified table
'             *Does not validate for the existence of the table
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: None required
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sTable - Name of the table for which to delete all its relationships
'
' Usage:
' ~~~~~~
' Call DeleteTableRelationships("Contacts")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2020-01-07              Initial Release
'---------------------------------------------------------------------------------------
Public Function DeleteTableRelationships(sTable As String) As Boolean
    Dim db                    As DAO.Database
    Dim rel                   As DAO.Relation

    On Error GoTo Error_Handler

    Set db = CurrentDb
    For Each rel In db.Relations
        If rel.Table = sTable Or rel.foreignTable = sTable Then
            db.Relations.Delete (rel.Name)
        End If
    Next rel
    DeleteTableRelationships = True

Error_Handler_Exit:
    On Error Resume Next
    If Not rel Is Nothing Then Set rel = Nothing
    If Not db Is Nothing Then Set db = Nothing
    Exit Function

Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: DeleteTableRelationships" & vbCrLf & _
           "Error Description: " & Err.description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Sub deb(str)
    
    Dim strArr As New clsArray: strArr.arr = str
    
    Dim lines As New clsArray
    
    Dim item, i As Integer: i = 0
    For Each item In strArr.arr
        lines.Add IIf(i > 0, vbTab, "") & "Debug.Print " & Esc(item & ": ") & "& " & item
        i = i + 1
    Next item
    
    CopyToClipboard lines.JoinArr(vbNewLine)
    
End Sub


Public Sub SetQueryDefSQL(QueryName, sqlStr)

    Dim qDef As QueryDef
    Dim db As Database
    Set db = CurrentDb
    Set qDef = db.QueryDefs(QueryName)
    qDef.sql = sqlStr
    
End Sub

Public Sub RowSource(fieldName)
        
    Dim arr As New clsArray
    
    arr.Add replace("Dim CustomerID: CustomerID = frm(""CustomerID"")", "CustomerID", fieldName)
    arr.Add replace(vbTab & "Dim filterStr: filterStr = ""CustomerID = 0""", "CustomerID", fieldName)
    arr.Add replace(vbTab & "If Not isFalse(CustomerID) Then", "CustomerID", fieldName)
    arr.Add replace(vbTab & vbTab & "filterStr = ""CustomerID = "" & CustomerID", "CustomerID", fieldName)
    arr.Add vbTab & "End If"
    
    arr.Add replace(vbTab & "Dim sqlStr: sqlStr = ""SELECT CustomerID FROM qryCustomerOrders WHERE "" & filterStr & "" ORDER BY OrderDate,CustomerOrderID""", "CustomerID", fieldName)
    arr.Add replace(vbTab & "frm(""fltrOrderDate"").rowsource = sqlStr", "CustomerID", fieldName)
    
    Dim str: str = arr.JoinArr(vbNewLine)
    CopyToClipboard str
    
End Sub

Public Function IsObjectAReport(obj As Object)

    IsObjectAReport = TypeName(obj) Like "Report_*"
    
End Function


Public Sub RemoveDuplicateRecords(TableName, fieldName)

    Dim sqlStr: sqlStr = "SELECT * FROM " & TableName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim RecordID, RecordIDsToDelete As New clsArray
    Dim TemporaryItem, TemporaryArr As New clsArray
    Dim PrimaryKey: PrimaryKey = GetPrimaryKeyFieldFromTable(TableName)
    
    Do Until rs.EOF
        Set TemporaryArr = New clsArray
        Dim fieldValue: fieldValue = rs.fields(fieldName)
        Dim PrimaryKeyValue: PrimaryKeyValue = rs.fields(PrimaryKey)
        Dim TemporaryIDs: TemporaryIDs = Elookups(TableName, fieldName & "=" & EscapeString(fieldValue, TableName, fieldName) & _
            " AND " & PrimaryKey & " <> " & PrimaryKeyValue, PrimaryKey)
        If Not isFalse(TemporaryIDs) Then
            TemporaryArr.arr = TemporaryIDs
            For Each TemporaryItem In TemporaryArr.arr
                RecordIDsToDelete.Add TemporaryItem
            Next TemporaryItem
        End If
        rs.MoveNext
    Loop
    
    If RecordIDsToDelete.count = 0 Then Exit Sub
    
    For Each RecordID In RecordIDsToDelete.arr
        RunSQL "DELETE FROM " & TableName & " WHERE " & PrimaryKey & " = " & RecordID
    Next RecordID
    
End Sub

Public Sub GetModelID(Model)

    Debug.Print Elookups("tblModels", "Model Like " & Esc("*" & Model & "*"), "ModelID", "ModelID DESC")
    
End Sub

Public Sub PrintControls(frmName, Optional textToSearch = "")

On Error GoTo ErrHandler:
    Dim frm As Form: Set frm = GetForm(frmName, True)
    
    If frm Is Nothing Then
        PrintFormsAndReports frmName, textToSearch
        Exit Sub
    End If

    Dim ctl As control, ctls As New clsArray
    
    For Each ctl In frm.controls
        Dim ControlName: ControlName = ctl.Name
        If ctl.Name Like "*" & textToSearch & "*" Then
            ctls.Add ControlName
        End If
    Next ctl
    
    Dim lines As New clsArray: lines.arr = SplitStringsIntoArray(ctls.JoinArr(" | "), 20)
    
    Dim line
    Dim ModifiedLines As New clsArray
    
    For Each line In lines.arr
        ModifiedLines.Add "''" & line
    Next line
    
    Debug.Print ModifiedLines.NewLineJoin
    
    Exit Sub
ErrHandler:
    If Err.Number = 2450 Then
        PrintFormsAndReports frmName, textToSearch
        Exit Sub
    End If
End Sub

Sub PrintFormsAndReports(filterStr, Optional contains = "")
    Dim db As DAO.Database
    Dim filterLower As String
    
    ' Get the current database
    Set db = CurrentDb
    
    ' Convert filterStr to lowercase for case-insensitive comparison
    filterLower = LCase(filterStr)
    
    Dim FormNames As New clsArray, ReportNames As New clsArray
    
    ' Populate the arrays with table and query names
    AddFilteredNames CurrentProject.AllForms, filterLower, FormNames
    AddFilteredNames CurrentProject.AllReports, filterLower, ReportNames
    
    ' Clean up
    Set db = Nothing
    
    Dim combinedArr As New clsArray
    
    If FormNames.count > 0 Then
        combinedArr.Add "FORMS: " & Join(SplitStringsIntoArray(FormNames.JoinArr(", "), 20), vbNewLine)
    End If
    
    If ReportNames.count > 0 Then
        combinedArr.Add "REPORTS: " & Join(SplitStringsIntoArray(ReportNames.JoinArr(", "), 20), vbNewLine)
    End If
    
    If FormNames.count + ReportNames.count = 1 Then
        PrintControls replace(replace(combinedArr.arr(0), "FORMS: ", ""), "REPORTS", ""), contains
        Exit Sub
    End If
    
    If combinedArr.count > 0 Then
        Debug.Print combinedArr.JoinArr(vbNewLine)
    Else
        Debug.Print "No form/report found."
    End If
End Sub


Public Function GetForm(frmName, Optional OpenFormWhenClosed As Boolean = False, Optional AsDesignView As Boolean = False) As Form

On Error GoTo ErrHandler:
    Dim frm As Form
    If Not IsFormOpen(frmName) And OpenFormWhenClosed Then
        DoCmd.OpenForm frmName
    End If
    
    If AsDesignView Then
        DoCmd.OpenForm frmName, acDesign, , , , acHidden
    End If
    
    Set frm = Forms(frmName)
    Set GetForm = frm

    Exit Function
ErrHandler:
    If Err.Number = 2450 Then
        Exit Function
    End If
    
End Function

Public Function GetReport(rptName, Optional OpenReportWhenClosed As Boolean = False, Optional AsDesignView As Boolean = False) As Report

On Error GoTo ErrHandler:
    Dim rpt As Report
    If Not IsReportOpen(rptName) And OpenReportWhenClosed Then
        DoCmd.OpenReport rptName, acDesign, , , , acHidden
    End If
    
    If AsDesignView Then
        DoCmd.OpenReport rptName, acDesign
    End If
    
    Set rpt = Reports(rptName)
    Set GetReport = rpt

    Exit Function
ErrHandler:
    If Err.Number = 2450 Then
        Exit Function
    End If
    
End Function


Public Sub CopyAndReplaceTable(ByVal TableName As String, ByVal newTableName As String)
    ' Delete the new table if it already exists
    On Error Resume Next
    DoCmd.DeleteObject acTable, newTableName
    On Error GoTo 0
    
    ' Copy the table to create a new one with the same structure and data
    DoCmd.CopyObject , newTableName, acTable, TableName
End Sub


Public Sub upsert()

    Dim arr As New clsArray
    arr.Add "Dim fields As New clsArray: fields.arr = ""field1,field2,field3"""
    arr.Add "Dim fieldValues As New clsArray"
    
    arr.Add "Set fieldValues = New clsArray"
    arr.Add "fieldValues.Add value1"
    arr.Add "fieldValues.Add value2"
    
    arr.Add "UpsertRecord ""TableName"", fields, fieldValues"
    
    CopyToClipboard arr.JoinArr(vbNewLine)
    
End Sub

Public Function Coalesce(ParamArray values() As Variant) As Variant
    Dim i As Integer
    
    ' Iterate through the values, except the last one
    For i = LBound(values) To UBound(values) - 1
        If Not isFalse(values(i)) Then
            Coalesce = values(i)
            Exit Function
        End If
    Next i
    
    ' If no value has passed the condition, return the last argument
    Coalesce = values(UBound(values))
End Function

Public Function QuoteAndJoin(possibleValues) As String

    If possibleValues = "" Then
        Exit Function
    End If
    
    Dim splitValues() As String
    Dim quotedValues() As String
    Dim i As Integer
    
    ' Split the values by comma
    splitValues = Split(possibleValues, ",")
    
    ' Initialize the array for quoted values
    ReDim quotedValues(LBound(splitValues) To UBound(splitValues))
    
    ' Loop through each value and add quotes
    For i = LBound(splitValues) To UBound(splitValues)
        quotedValues(i) = """" & Trim(splitValues(i)) & """"
    Next i
    
    ' Join the values back with semicolon
    QuoteAndJoin = Join(quotedValues, ";")
End Function

''accepts text from the module e.g. ClientPath & "src\components\" & ModelPath & "\" & ModelName & "FilterForm.tsx"
Public Function GetFilePathTemplate(Text As String) As String
    
    Dim inputString As String
    Dim regex As Object
    Dim match As Object
    Dim resultString As String
    Dim replacedText As String
    Dim firstIteration As Boolean
    
    ' The input string
    inputString = Text
    resultString = inputString
    firstIteration = True
    
    ' Create a new RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    
    Do While True
        If firstIteration Then
            ' Define the regex pattern to match (.*?)\s& at the start
            regex.pattern = "^(.*?)\s&\s"
            firstIteration = False
        Else
            ' Define the regex pattern to match &\s(.*?)\s
            regex.pattern = "\s&\s(.*?)\s&\s"
        End If
        
        regex.Global = False ' Apply the replacement one at a time
        
        If regex.Test(resultString) Then
            ' Execute the regex search and get the first match
            Set match = regex.Execute(resultString)(0)
            
            ' Replace the matched pattern with the appropriate group in square brackets
            If match.SubMatches(0) <> "" Then
                replacedText = "[" & match.SubMatches(0) & "]"
            End If
            
            ' Replace the first matched pattern in the result string
            resultString = Left(resultString, match.FirstIndex) & replacedText & Mid(resultString, match.FirstIndex + match.length + 1)
        Else
            Exit Do
        End If
    Loop
    
    ' Remove all double quotes from the result string
    resultString = replace(resultString, """", "")
    
    ' Output the result
    GetFilePathTemplate = resultString
    
End Function


Public Function UnescapeValue(Text) As Variant

    If Text = "Null" Then
        UnescapeValue = Null
        Exit Function
    End If
    ' Check if the Text starts and ends with double quotes
    If Len(Text) > 1 And Left(Text, 1) = """" And Right(Text, 1) = """" Then
        ' Remove the first and last character (the double quotes)
        UnescapeValue = Mid(Text, 2, Len(Text) - 2)
    Else
        ' Return the original Text if it does not start and end with double quotes
        UnescapeValue = Text
    End If
    
    If Len(UnescapeValue) > 1 And Left(UnescapeValue, 1) = "#" And Right(UnescapeValue, 1) = "#" Then
        ' Remove the first and last character (the double quotes)
        UnescapeValue = Mid(UnescapeValue, 2, Len(UnescapeValue) - 2)
    Else
        ' Return the original Text if it does not start and end with double quotes
        UnescapeValue = UnescapeValue
    End If
    
    
End Function


Public Sub iit(condition)
    
    Dim arr As New clsArray
    arr.Add "If isFalse(" & condition & ") Then"
    arr.Add vbTab & vbTab & "MsgBox ""Some message here"", vbOKOnly"
    arr.Add vbTab & vbTab & "Exit Function"
    arr.Add vbTab & "End If"
    
    Dim str: str = arr.JoinArr(vbNewLine)
    CopyToClipboard str
    
End Sub


Public Function IsEven(Number As Long) As Boolean
    ' Check if the number is even
    If Number Mod 2 = 0 Then
        IsEven = True
    Else
        IsEven = False
    End If
End Function


Public Function CountRecordset(rs As Recordset) As Long
    
    If rs Is Nothing Then Exit Function
    CountRecordset = rs.recordCount
    
    If CountRecordset <> 0 Then
        rs.MoveLast
        rs.MoveFirst
        CountRecordset = rs.recordCount
    End If

End Function

''GetControlProperty "frmQualityControl","cmdOK","BorderStyle,BorderWidth,BackColor,HoverColor,PressedColor,HoverForeColor,PressedForeColor,ForeColor"
Public Sub GetControlProperty(frmName, ctlName, propertyNames)

    If Not IsFormOpen(frmName) Then Exit Sub
    
    Dim frm As Form: Set frm = Forms(frmName)
    Dim ctl As control
    If DoesPropertyExists(frm, ctlName) Then
        Set ctl = frm(ctlName)
    End If
    
    Dim prop, propertyArr As New clsArray: propertyArr.arr = propertyNames
    
    For Each prop In propertyArr.arr
        Debug.Print "frm(""" & ctlName & """)(""" & prop & """) = " & ctl.Properties(prop)
    Next prop
    
End Sub

Public Function ParseISO8601ToDateTime(isoDateTimeString As Variant) As Variant
    
    If isFalse(isoDateTimeString) Then Exit Function

    Dim dateTimeParts() As String
    Dim dateStringPart As String
    Dim timeStringPart As String
    Dim yearPart As String
    Dim monthPart As String
    Dim dayPart As String
    Dim hourPart As String
    Dim minutePart As String
    Dim secondPart As String
    
    ' Split the ISO 8601 string into date and time components
    dateTimeParts = Split(isoDateTimeString, "T")
    dateStringPart = dateTimeParts(0)
    timeStringPart = dateTimeParts(1)
    
    ' Further split to get individual date components
    yearPart = Mid(dateStringPart, 1, 4)
    monthPart = Mid(dateStringPart, 6, 2)
    dayPart = Mid(dateStringPart, 9, 2)
    
    ' Further split to get individual time components
    hourPart = Mid(timeStringPart, 1, 2)
    minutePart = Mid(timeStringPart, 4, 2)
    secondPart = Mid(timeStringPart, 7, 2)
    
    ' Construct a VBA-friendly date string in MM/DD/YYYY format
    Dim vbaDateString As String
    vbaDateString = monthPart & "/" & dayPart & "/" & yearPart
    
    ' Attempt to combine date and time into a single Date value
    ' Note: This is a simplification and may not accurately reflect the original time zone or DST adjustments
    Dim combinedDateTime As Date
    combinedDateTime = CDate(vbaDateString) + TimeValue(hourPart & ":" & minutePart & ":" & secondPart)
    
    ParseISO8601ToDateTime = combinedDateTime
    
End Function



Function ListTableRowCounts()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim rs As DAO.Recordset
    Dim rowCount As Long
    Dim tableInfo() As Variant
    Dim i As Integer
    
    ' Set a reference to the current database
    Set db = CurrentDb
    
    ' Initialize an array to store table information
    ReDim tableInfo(0 To db.TableDefs.count - 1, 1)
    
    ' Loop through each table
    For Each tdf In db.TableDefs
        ' Check if the table is a user table (not a system table)
        If (tdf.attributes And dbSystemObject) = 0 Then
            ' Open a recordset for the table
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM [" & tdf.Name & "]")
            
            ' Get the row count
            rowCount = rs.fields(0).value
            
            ' Store table information in the array
            tableInfo(i, 0) = tdf.Name
            tableInfo(i, 1) = rowCount
            
            ' Close the recordset
            rs.Close
            
            ' Increment the index
            i = i + 1
        End If
    Next tdf
    
    ' Sort the array in descending order based on row count
    Sort2DArrayDescending tableInfo, 1
    
    ' Print the sorted table information
    For i = 0 To UBound(tableInfo, 1)
        Debug.Print "Table: " & tableInfo(i, 0) & ", Row Count: " & tableInfo(i, 1)
    Next i
    
    ' Clean up
    Set db = Nothing
    Set tdf = Nothing
    Set rs = Nothing
End Function

Sub Sort2DArrayDescending(ByRef arr() As Variant, ByVal sortIndex As Integer)
    Dim i As Integer, j As Integer, k As Integer
    Dim temp As Variant
    
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If arr(i, sortIndex) < arr(j, sortIndex) Then
                ' Swap rows if the current row has a smaller value than the next row
                For k = LBound(arr, 2) To UBound(arr, 2)
                    temp = arr(i, k)
                    arr(i, k) = arr(j, k)
                    arr(j, k) = temp
                Next k
            End If
        Next j
    Next i
End Sub

Public Function RunCommandSaveRecord()

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    
End Function

Public Function SelectAllContent(frm, ControlName)
    Dim txtBox As control: Set txtBox = frm(ControlName)
    txtBox.SetFocus
    txtBox.SelStart = 0
    txtBox.SelLength = Len(txtBox.Text)
End Function

Public Function GetKVPairs(QueryName, rs As Recordset) As String

    Dim KVPairs As New clsArray
    Dim rs2 As Recordset: Set rs2 = ReturnRecordset("SELECT * FROM qryConfigEnumeratorFields WHERE Not Exclude AND QueryName = " & Esc(QueryName) & " ORDER BY FieldOrder,ConfigEnumeratorFieldID")
    
    Do Until rs2.EOF
        Dim fieldName: fieldName = rs2.fields("fieldName"): If ExitIfTrue(isFalse(fieldName), "Field Name is empty..") Then Exit Function
        Dim VariableName: VariableName = rs2.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "Variable Name is empty..") Then Exit Function
        Dim fld As field: Set fld = rs.fields(fieldName)
        Dim fieldValue: fieldValue = fld.value
        
        If IsNull(fieldValue) Then
            fieldValue = "null"
'        ElseIf QueryName = "qrySeqModels" And FieldName = "NavItemIcon" Then
'            FieldValue = FieldValue
        Else
            If isFieldBoolean(fld.Type) Then
                fieldValue = IIf(fieldValue, "true", "false")
            ElseIf Not isFieldNumeric(fld.Type) Then
                fieldValue = Esc(fieldValue)
            End If
        End If
        
        KVPairs.Add VariableName & ": " & fieldValue & ","
        
        rs2.MoveNext
    Loop
    
    GetKVPairs = KVPairs.JoinArr(vbNewLine)
    
End Function

Public Function FirstCharLowercase(str) As String
    If Len(str) > 0 Then
        FirstCharLowercase = LCase(Left(str, 1)) & Mid(str, 2)
    Else
        FirstCharLowercase = str
    End If
End Function

Public Function isFieldBoolean(fldType) As Boolean

    isFieldBoolean = fldType = dbBoolean
    
End Function

Public Function isFieldNumeric(fldType) As Boolean

    isFieldNumeric = fldType = dbBoolean Or fldType = dbInteger Or fldType = dbLong Or fldType = dbSingle Or fldType = dbDouble Or fldType = dbByte Or fldType = dbDecimal
    
End Function

Public Function PluralizeWord(ByVal word, Optional Number As Long = 2) As String
    Dim pluralizedWord As String
    
    If Number <= 1 Then
        PluralizeWord = word
        Exit Function
    End If
    
    If Right(word, 1) = "y" Then
        pluralizedWord = Left(word, Len(word) - 1) & "ies"
    Else
        pluralizedWord = word & "s"
    End If
    
    PluralizeWord = pluralizedWord
End Function

Public Function GetPrimaryKeyFieldFromTable(TableName) As String
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.field
    
    Set db = CurrentDb
    Set td = db.TableDefs(TableName)
    
    For Each idx In td.indexes
        If idx.Primary Then
            For Each fld In idx.fields
                GetPrimaryKeyFieldFromTable = fld.Name
                Exit Function
            Next fld
        End If
    Next idx
    
    GetPrimaryKeyFieldFromTable = ""
    
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    Set db = Nothing
End Function

Public Function GetGeneratedByFunctionSnippet(str, FunctionName, Optional templateName = "", Optional jsx As Boolean = False, Optional AsOneLine As Boolean = False) As String
    Dim lines As New clsArray
    
    Dim Text: Text = "Generated by " & FunctionName
    
    If Not isFalse(templateName) Then
        Text = Text & " - " & templateName
    End If
    
    If jsx Then
        lines.Add "{/* " & Text & " */}"
    Else
        If AsOneLine Then
            lines.Add str & "//" & Text
        Else
            lines.Add "//" & Text
        End If
    End If
    
    If Not AsOneLine Then lines.Add str
    
    GetGeneratedByFunctionSnippet = lines.NewLineJoin
End Function

Public Function ReturnResultArray(ResultMessage As String, Optional result As String = "Error") As Variant

    Dim Results(1) As String
    Results(0) = result
    Results(1) = ResultMessage
    
    ReturnResultArray = Results
    
End Function

Public Function IfError(value) As Double
On Error GoTo ErrorHandler:
    IfError = IIf(IsError(value), 0, value)
    Exit Function
ErrorHandler:
    IfError = 0
End Function

Public Function EnumarateFields(tblName As String)

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(tblName)
        
    Dim fld As field
    
    DoCmd.SetWarnings False
    For Each fld In rs.fields
        
        CurrentDb.Execute "INSERT INTO tblTableFields (TableName,FieldName) VALUES ('" & tblName & "','" & fld.Name & "')"
    
    Next fld
    DoCmd.SetWarnings True
    
End Function

Public Function SplitStringsIntoArray(ByVal longString, ByVal n) As Variant
    Dim result() As String
    Dim i As Integer, j As Integer
    Dim currentPos As Integer
    Dim segment As String
    
    ' Initialize the starting position and the array size
    currentPos = 1
    i = 0
    ReDim result(0 To 0)
    
    Do While currentPos <= Len(longString)
        ' Get the segment of n characters
        segment = Mid(longString, currentPos, n)
        
        ' Check if the next character is a part of the word (not a whitespace)
        If (currentPos + n <= Len(longString)) And (Mid(longString, currentPos + n, 1) <> " ") Then
            ' Extend the segment to include the full word
            Do While (currentPos + n <= Len(longString)) And (Mid(longString, currentPos + n, 1) <> " ")
                n = n + 1
                segment = Mid(longString, currentPos, n)
            Loop
        End If
        
        ' Add the segment to the result array
        If i > UBound(result) Then
            ReDim Preserve result(0 To i)
        End If
        
        result(i) = Trim(segment)
        i = i + 1
        currentPos = currentPos + n
    Loop
    
    SplitStringsIntoArray = result
End Function


Public Function PrintFields(tblName, Optional contains = "")

On Error GoTo ErrHandler
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(tblName)
        
    Dim fld As field
    Dim fields As New clsArray
    
    For Each fld In rs.fields
        If isFalse(contains) Then
            fields.Add fld.Name
        Else
            If fld.Name Like "*" & contains & "*" Then
                fields.Add fld.Name
            End If
        End If
    Next fld
    
    Dim result As String
    result = "TABLE: " & tblName & " Fields: " & fields.JoinArr(" | ")
    
    If fields.count = 1 Then
        D fields.arr(0)
    End If
    
    Dim lines As New clsArray: lines.arr = SplitStringsIntoArray(result, 25)
    
    Dim line
    Dim ModifiedLines As New clsArray
    
    For Each line In lines.arr
        ModifiedLines.Add "''" & line
    Next line
    
    Debug.Print ModifiedLines.NewLineJoin
    
    Exit Function
    
ErrHandler:
    If Err.Number = 3078 Then
        PrintTables tblName, contains
    End If
    
    ''CopyToClipboard finalResult
End Function

Sub PrintTables(filterStr, Optional contains = "")
    Dim db As DAO.Database
    Dim filterLower As String
    
    ' Get the current database
    Set db = CurrentDb
    
    ' Convert filterStr to lowercase for case-insensitive comparison
    filterLower = LCase(filterStr)
    
    Dim TableNames As New clsArray, queryNames As New clsArray
    
    ' Populate the arrays with table and query names
    AddFilteredNames db.TableDefs, filterLower, TableNames
    AddFilteredNames db.QueryDefs, filterLower, queryNames
    
    ' Clean up
    Set db = Nothing
    
    Dim combinedArr As New clsArray
    
    If TableNames.count > 0 Then
        combinedArr.Add "TABLES: " & Join(SplitStringsIntoArray(TableNames.JoinArr(", "), 20), vbNewLine)
    End If
    
    If queryNames.count > 0 Then
        combinedArr.Add "QUERIES: " & Join(SplitStringsIntoArray(queryNames.JoinArr(", "), 20), vbNewLine)
    End If
    
    If combinedArr.count = 1 Then
        If TableNames.count = 1 Then
            PrintFields TableNames.arr(0), contains
            Exit Sub
        End If
        
        If queryNames.count = 1 Then
            PrintFields queryNames.arr(0), contains
            Exit Sub
        End If
    End If
    
    If combinedArr.count > 0 Then
        Debug.Print combinedArr.JoinArr(vbNewLine)
    Else
        Debug.Print "No recordset found."
    End If
End Sub

Private Sub AddFilteredNames(container As Object, filterLower As String, namesArray As clsArray)
    Dim item As Object
    For Each item In container
        If InStr(1, LCase(item.Name), filterLower) > 0 And NameHasNoTilde(item.Name) Then
            namesArray.Add item.Name
        End If
    Next item
End Sub

Private Function NameHasNoTilde(nameStr As String) As Boolean
    NameHasNoTilde = (InStr(1, nameStr, "~") = 0)
End Function

Public Function PrintRecordsetFields(rs As Recordset)

        
    Dim fld As field
    Dim fields() As String
    Dim i As Integer
    
    For Each fld In rs.fields
        ReDim Preserve fields(i)
        fields(i) = fld.Name
        i = i + 1
    Next fld
    
    Debug.Print Join(fields, "|")
    
End Function

Public Function Divide(Numerator, Denominator) As Double
    
    If IsNull(Numerator) Or IsNull(Denominator) Then
        Divide = 0
        Exit Function
    End If
    
    If Numerator = 0 Or Denominator = 0 Then
        Divide = 0
        Exit Function
    End If
    
    Divide = Numerator / Denominator
    
End Function


Public Function ArrayLength(arr As Variant) As Integer
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
On Error GoTo ArrayLengthError: 'array is empty
        ArrayLength = UBound(arr) + 1
        Exit Function
ArrayLengthError:
    On Error GoTo 0
    ArrayLength = 0
End Function

Public Function LogIn(Optional ctl As IRibbonControl)
    g_userID = 1
End Function

Public Function ShowError(ErrorStr As String)
    MsgBox ErrorStr, vbCritical + vbOKOnly
End Function

Public Function isFalse(value) As Boolean
On Error GoTo ErrHandler
    
    If IsMissing(value) Then
        isFalse = True
    Else
        isFalse = value = "" Or IsNull(value) Or IsEmpty(value)
    End If
    Exit Function
ErrHandler:
If Err.Number = 3167 Then
    isFalse = True
    Exit Function
End If
'Exit_isFalse:
'    Exit Function
'Err_isFalse:
'    LogError Err.Number, Err.description, "isFalse"
'    Resume Exit_isFalse
End Function

Public Function CdblNZ(val As Variant) As Double
    CdblNZ = CDbl(Nz(val, 0))
End Function

Public Function SetQueryDef()
    'Set recordsource -> tblAccOutInteractions
    ''qryLogs
    Dim logSQL As String
    logSQL = "SELECT * FROM qryLogs WHERE TableName = ""tblAccOutInteractions"" And EventName = ""ADD"""
    
    Dim sqlStr As String
    sqlStr = "SELECT tblAccOutInteractions.*,UserName,DateTime FROM tblAccOutInteractions LEFT JOIN (" & logSQL & ") As qryLogs ON tblAccOutInteractions.AccOutInteractionID = qryLogs.RecordID"
    
    Dim qDef As QueryDef
    Set qDef = CurrentDb.QueryDefs("qryAccOutInteractions")
    qDef.sql = sqlStr
End Function

Public Function PromptLogin()
    
    If isFalse(g_userID) Then
        If MsgBox("Login using developer user?", vbYesNo) = vbYes Then
            LogIn
        End If
    End If
    
End Function


Public Function ReturnYesNo(TF As Boolean) As String
    If TF Then
        ReturnYesNo = "YES"
    Else
        ReturnYesNo = "NO"
    End If
End Function


Public Function GenerateUPCID() As String

    Dim randNumber As Long, upperbound, lowerbound, UPCStr As String
    upperbound = 9999999: lowerbound = 1
    
    Do Until UPCStr <> "" Or isPresent("tblOrders", "PackCycleID = '" & UPCStr & "'")
        randNumber = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
        UPCStr = Format$(randNumber, "0000000")
    Loop
    
    GenerateUPCID = UPCStr
    
End Function

Public Function ExitIfTrue(condition As Boolean, msg As String) As Boolean

    ExitIfTrue = False
    If condition Then
        ShowError msg
        ExitIfTrue = True
    End If
    
End Function
Public Function GenerateFieldNamesString(ByVal tblName As String, Optional ByVal IgnoreFields = "", Optional prefix As Variant = "") As String

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(tblName)
        
    Dim fld As field
    Dim fields() As String, IgnoreArray() As String
    Dim i As Integer
    
    IgnoreArray = Split(IgnoreFields, ",")
    
    For Each fld In rs.fields
        If Not IsInArray(fld.Name, IgnoreArray) Then
            ReDim Preserve fields(i)
            If prefix <> "" Then
                fields(i) = prefix & "." & fld.Name
            Else
                fields(i) = fld.Name
            End If
            i = i + 1
        End If
    Next fld
    
    GenerateFieldNamesString = Join(fields, ", ")
    
End Function

Public Function GetTotalRecordCount(ByVal rs As Recordset) As Long

If rs.EOF Then
    GetTotalRecordCount = 0
    Exit Function
Else
    rs.MoveLast
    GetTotalRecordCount = rs.recordCount
    rs.MoveFirst
End If

End Function

Public Function InchToTwip(inch) As Double
    InchToTwip = CDbl(inch) * 1440
End Function

Public Function DoesTableExist(TableName, Optional DatabasePath = "") As Boolean

    Dim db As DAO.Database
    Dim td As TableDef
    
    If Not isFalse(DatabasePath) Then
        Set db = OpenDatabase(DatabasePath)
    Else
        Set db = CurrentDb
    End If
    
    For Each td In db.TableDefs
        If td.Name = TableName Then
            DoesTableExist = True
            Exit Function
        End If
    Next td
        
    Set td = Nothing
    Set db = Nothing
    
    DoesTableExist = False
    Exit Function
    

End Function

Public Function deleteTableIfExists(ByVal TableName As String, Optional DatabasePath = "") As Boolean

On Error GoTo Err_Handler:
    'Runs through all table names in CurrentDB and deletes table if name matches
    Dim db As DAO.Database
    Dim td As TableDef
    
    If Not isFalse(DatabasePath) Then
        Set db = OpenDatabase(DatabasePath)
    Else
        Set db = CurrentDb
    End If
    
    
    For Each td In db.TableDefs
        If td.Name = TableName Then
            db.TableDefs.Delete TableName
            db.TableDefs.Refresh
            Set td = Nothing
            Set db = Nothing
            Exit For
        End If
    Next td
        
    Set td = Nothing
    Set db = Nothing
    
    deleteTableIfExists = True
    Exit Function
    
Err_Handler:
    
    If Err.Number = 3211 Then
        ShowError """" & TableName & """ is open. Please close it first.."
        Exit Function
    Else
        Debug.Print "deleteTableIfExists error: " & Err.Number & ": " & Err.description
        Exit Function
    End If

End Function

Public Sub deleteQueryIfExists(ByVal QueryName As String)
'Runs through all table names in CurrentDB and deletes table if name matches
    Dim db As DAO.Database
    Dim qdf As QueryDef

    Set db = CurrentDb

    For Each qdf In db.QueryDefs
        If qdf.Name = QueryName Then
            db.QueryDefs.Delete QueryName
            db.QueryDefs.Refresh
            Set qdf = Nothing
            Set db = Nothing
            Exit For
        End If
    Next qdf
    
    Set qdf = Nothing
    Set db = Nothing
    
End Sub

Public Function DoesObjectExists(obj As Object)

    ''Checks wether a property exists within the parent obj
    ''obj => Parent object | propertyName => Name of the property
    Dim tempObj
On Error Resume Next
    Set tempObj = obj
    DoesObjectExists = (Err = 0)
On Error GoTo 0

End Function


Public Function DoesPropertyExists(obj As Object, propertyName)
    ''Checks wether a property exists within the parent obj
    ''obj => Parent object | propertyName => Name of the property
    Dim tempObj
On Error Resume Next
    Set tempObj = obj(propertyName)
    DoesPropertyExists = (Err = 0)
On Error GoTo 0

End Function

Function AddSpaces(pValue) As String
    'Update 20140723
    Dim xOut As String
    Dim i, xAsc
    xOut = VBA.Left(pValue, 1)
    For i = 2 To VBA.Len(pValue)
       xAsc = VBA.Asc(VBA.Mid(pValue, i, 1))
       If xAsc >= 65 And xAsc <= 90 Then
          xOut = xOut & " " & VBA.Mid(pValue, i, 1)
       Else
          xOut = xOut & VBA.Mid(pValue, i, 1)
       End If
    Next
    AddSpaces = xOut
End Function

Function RemoveSpaces(pValue) As String
    'Update 20140723
    RemoveSpaces = replace(pValue, " ", "")
End Function

Public Function ReturnStringBasedOnType(fieldVal As Variant, ControlType As Integer) As String

    If IsNull(fieldVal) Then
        ReturnStringBasedOnType = "Null"
        Exit Function
    End If
    
    Select Case ControlType
        Case 10, 12:
            ReturnStringBasedOnType = """" & fieldVal & """"
        Case 8:
            ReturnStringBasedOnType = "#" & SQLDate(fieldVal) & "#"
        Case Else:
            ReturnStringBasedOnType = fieldVal
    End Select
    
End Function

Public Function myRandVal(nm As String) As String
' Randomizes Text Strings and/or numbers
' Usage: In a query - NameOfFieldRnd: myRandVal([NameOfField])
'        In a Form Control's "ControlSource" property - =myRandVal([NameOfField])
' Test in debug window - ?myRandVal("12345 UPPER & lower Case")
' Modified by Bob Raskew [URL="tel:11/1/2003"]11/1/2003[/URL] - Added Number scrambling
Dim myChr As String
Dim myAsc As Integer
Dim i As Integer
    myRandVal = ""
    If Len(nm) = 0 Then
        myRandVal = vbNullString
        Exit Function
    Else
        For i = 1 To Len(nm)
            myChr = Mid(nm, i, 1)
            If Asc(myChr) >= 65 And Asc(myChr) <= 90 Then
                myAsc = Int((90 - 65 + 1) * Rnd + 65)
              ElseIf Asc(myChr) >= 97 And Asc(myChr) <= 122 Then
                myAsc = Int((122 - 97 + 1) * Rnd + 97)
              ElseIf Asc(myChr) >= 48 And Asc(myChr) <= 57 Then
                myAsc = Int((57 - 48 + 1) * Rnd + 48)
              Else
                myAsc = Asc(myChr)
            End If
            myChr = Chr(myAsc)
            myRandVal = myRandVal & myChr
        Next i
    End If
End Function

Public Function myRandEmail(nm As String) As String
' Randomizes Text Strings and/or numbers
' Usage: In a query - NameOfFieldRnd: myRandVal([NameOfField])
'        In a Form Control's "ControlSource" property - =myRandVal([NameOfField])
' Test in debug window - ?myRandVal("12345 UPPER & lower Case")
' Modified by Bob Raskew [URL="tel:11/1/2003"]11/1/2003[/URL] - Added Number scrambling
    Dim splStr() As String

    myRandEmail = ""
    If Len(nm) = 0 Then
        myRandEmail = vbNullString
        Exit Function
    ElseIf InStr(1, nm, "@") Then
        splStr = Split(nm, "@")
        splStr(0) = Left(CStr(10000000000# * Rnd), 10)
        splStr(1) = "example.com"
        myRandEmail = Join(splStr, "@")
    End If
End Function

Public Function makeQuery(sqlStr)

    Dim db As DAO.Database
    Dim qDef As DAO.QueryDef
    
    Set db = CurrentDb
    
    If DoesPropertyExists(db.QueryDefs, "qryTestQuery") Then
        Set qDef = db.QueryDefs("qryTestQuery")
    Else
        Set qDef = db.CreateQueryDef("qryTestQuery")
    End If

    qDef.sql = sqlStr
    qDef.Close
    db.Close
    
    DoCmd.OpenQuery "qryTestQuery", acViewDesign
    
End Function

Public Function GenerateUpdateStatements(targetFieldNames, targetTableName, updateFrom) As Variant

    Dim updateArr As New clsArray, filterArr As New clsArray, targetFieldArr As New clsArray, returnArr As New clsArray
    Dim TargetField As Variant, trimmedtargetField As String, origField As String, tempField As String
    
    targetFieldArr.arr = targetFieldNames
    
    For Each TargetField In targetFieldArr.arr
        trimmedtargetField = Trim(TargetField)
        origField = targetTableName & "." & trimmedtargetField
        tempField = "[" & updateFrom & "]![" & trimmedtargetField & "]"
        
        updateArr.Add origField & " = " & tempField
        filterArr.Add origField & " <> " & tempField
        
    Next TargetField
    
    Dim filterStatement As String, updateStatement As String
    updateStatement = updateArr.JoinArr
    filterStatement = filterArr.JoinArr(" OR ")
    
    returnArr.Add updateStatement
    returnArr.Add filterStatement
    
    GenerateUpdateStatements = returnArr.arr
    
End Function

Public Function HasProperty(obj As Object, strPropName) As Boolean
    'Purpose:   Return true if the object has the property.
    Dim varDummy As Variant
    
    On Error Resume Next
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)
End Function


Public Function GenerateJoinObj(Source, LeftFields, Optional Alias = "", Optional RightFields = "", Optional JoinType = "") As clsJoin

    Dim joinObj As clsJoin
    Set joinObj = New clsJoin
    With joinObj
      .Source = Source
      .LeftFields = LeftFields
    End With
    
    If Alias <> "" Then joinObj.Alias = Alias
    If RightFields <> "" Then joinObj.RightFields = RightFields
    If JoinType <> "" Then joinObj.JoinType = JoinType
    
    Set GenerateJoinObj = joinObj
  
End Function

Public Function concat(ParamArray var() As Variant) As String
    Dim i As Integer
    Dim tmp As String
    For i = LBound(var) To UBound(var)
        tmp = tmp & var(i)
    Next
    concat = tmp
End Function

Public Function CenterString(xInput As String, xLength As Long)
    Dim xM As Variant
    xM = Space(((xLength / 2) - (Len(xInput) / 2) + 1)) + xInput
    CenterString = xM + Space(xLength - Len(xM))
End Function

'Public Sub CopyToClipboard(str)
'
'    Dim objCP As Object
'    Set objCP = CreateObject("HtmlFile")
'
'    objCP.ParentWindow.ClipboardData.setData "text", str
'
'End Sub

Sub CopyToClipboard(str)
    'Create a new DataObject to store the string
    Dim clipboard As Object
    Set clipboard = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    'Set the string to the DataObject's text property
    clipboard.SetText str
    On Error Resume Next
    'Copy the DataObject to the clipboard
    clipboard.PutInClipboard

    ''MsgBox "Copied to clipboard."
End Sub

Public Sub rs()
    Dim str: str = "Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)"
    ''Debug.Print str
    CopyToClipboard str
End Sub

Public Sub D(str)

    Dim strArr As New clsArray: strArr.arr = str
    
    Dim lines As New clsArray
    
    Dim item, i As Integer: i = 0
    For Each item In strArr.arr
        lines.Add IIf(i > 0, vbTab, "") & replace("Dim [item]: [item] = rs.Fields(""[item]"")", "[item]", item)
        i = i + 1
    Next item
    
    CopyToClipboard lines.JoinArr(vbNewLine)
    
End Sub

Public Sub V(str)

    str = "If ExitIfTrue(IsFalse(" & str & ")," & Esc(str & " is empty..") & ") Then exit function"
    ''Debug.Print str
    CopyToClipboard str
    
End Sub

Public Sub g(FunctionName, Optional templateName = "")
    
    Dim str
    str = FunctionName & " = GetGeneratedByFunctionSnippet(" & FunctionName & "," & Esc(FunctionName) & "," & Esc(templateName) & ")"
    CopyToClipboard str
    
End Sub

Public Sub f(FunctionName)

    Dim str: str = FunctionName & " = GetGeneratedByFunctionSnippet(" & FunctionName & ", " & Esc(FunctionName) & ")"
    CopyToClipboard str
End Sub

Public Sub a(str, Optional FunctionName = "")

    Dim arr As New clsArray
    arr.Add "Dim " & str & " as new clsArray: " & str & ".arr = """""
    arr.Add str & ".add ""//Generated by " & IIf(Not isFalse(FunctionName), FunctionName, "replacewithyourfunctioname") & """"
    
    If Not isFalse(FunctionName) Then
        arr.Add str & ".add " & FunctionName
        arr.Add FunctionName & " = lines.NewLineJoin"
    End If
    
    str = arr.JoinArr(vbNewLine & vbTab)
    ''Debug.Print str
    CopyToClipboard str
End Sub

Public Sub w(Optional FunctionName = "")
    
    Dim lines As New clsArray
    lines.Add "Dim ModelPath: ModelPath = rs.fields(""ModelPath""): If ExitIfTrue(isFalse(ModelPath), ""ModelPath is empty.."") Then Exit Function"
    lines.Add "Dim ClientPath: ClientPath = rs.fields(""ClientPath""): If ExitIfTrue(isFalse(ClientPath), ""ClientPath is empty.."") Then Exit Function"
    
    lines.Add "Dim filePath: filePath = ClientPath & ""src\app\api\"" & ModelPath & ""\route.ts"""
    
    lines.Add "WriteToFile filePath, " & FunctionName & ", SeqModelID," & Esc(FunctionName)
    
    CopyToClipboard lines.JoinArr(vbNewLine)
    
End Sub

Public Sub onerror()
    
    Dim lines As New clsArray
    lines.Add "On Error GoTo ErrHandler:"
    lines.Add "ErrHandler:"
    
    lines.Add "If Err.Number = [SomeNumberHere] Then"
    lines.Add "End If"
    
    CopyToClipboard lines.JoinArr(vbNewLine)
    
End Sub


Public Sub dv(str)
    
    Dim strArr As New clsArray: strArr.arr = str
    
    Dim lines As New clsArray
    
    Dim item, i As Integer: i = 0
    
    Dim subLines As New clsArray
    
    
    For Each item In strArr.arr
        Set subLines = New clsArray
        subLines.Add replace("Dim [item]: [item] = rs.Fields(""[item]"")", "[item]", item)
        subLines.Add replace("If ExitIfTrue(IsFalse([item]),""""""[item]"""" is empty.."") Then exit function", "[item]", item)
        lines.Add IIf(i > 0, vbTab, "") & subLines.JoinArr(":")
        ''lines.Add IIf(i > 0, vbTab, "") & replace("Dim [item]: [item] = rs.Fields(""[item]"")", "[item]", item)
        i = i + 1
    Next item
    
    CopyToClipboard lines.JoinArr(vbNewLine)
    
End Sub

Public Sub s(str)

    str = "sqlStr = ""SELECT * FROM " & str & """"
    Debug.Print str
    CopyToClipboard str
    
End Sub

Public Sub p(str As String, Optional contains = "")

    PrintFields str, contains
    
End Sub

Public Sub r(FunctionName, strToBeReplaced)

    Dim str: str = FunctionName & " = Replace(" & FunctionName & ", ""[" & strToBeReplaced & "]"", " & strToBeReplaced & ")"
    CopyToClipboard str
    
End Sub

Public Function Esc(str) As String
    Esc = EscapeString(str)
End Function

Public Sub rsLoop()

    Dim strs As New clsArray
    strs.Add "Do until rs.EOF"
    strs.Add vbTab & vbTab & "rs.Movenext"
    strs.Add vbTab & "Loop"
    
    Dim str: str = strs.JoinArr(vbNewLine)
    CopyToClipboard str
    
End Sub

Public Sub forEachLoop(Optional VariableName = "item")

    Dim strs As New clsArray
    strs.Add "Dim " & VariableName & ", " & VariableName & "s as New clsArray"
    strs.Add "Dim i As Integer: i = 0"
    strs.Add vbTab & "For each " & VariableName & " in " & VariableName & "s.arr"
    strs.Add vbTab & vbTab & "Debug.Print " & VariableName
    strs.Add vbTab & vbTab & "i = i + 1"
    strs.Add vbTab & "Next " & VariableName
    
    Dim str: str = strs.JoinArr(vbNewLine)
    CopyToClipboard str
    
End Sub

Function EndOfMonth(currDate As Date) As Date
    EndOfMonth = DateSerial(Year(currDate), Month(currDate) + 1, 0)
End Function

Public Function GetFieldTypeFromRS(rsName, fieldName)

    Dim rs As Recordset: Set rs = ReturnRecordset(rsName)
    
    GetFieldTypeFromRS = rs.fields(fieldName).Type
    
End Function

Public Function SanitizeJSONString(inputStr As String) As String
    Dim lines() As String
    lines = Split(inputStr, vbCrLf)
    
    Dim level As Integer
    Dim output As String
    Dim i
    output = ""
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        If Len(line) = 0 Then
            ' Skip empty lines
            GoTo SkipIteration
        End If
        
        If Left(line, 1) = "}" Then
            ' Decrease level if the line is ending a block
            level = level - 1
        End If
        
        If level >= 0 Then
            ' Only sanitize lines inside a block
            output = output & String(level, "  ") & line & vbCrLf
        Else
            output = output & line & vbCrLf
        End If
        
        If Right(line, 1) = "{" Then
            ' Increase level if the line is starting a block
            level = level + 1
        End If
SkipIteration:
    Next
    
    SanitizeJSONString = Trim(output)
End Function

''Converts Option ('Character','Weapon','Power','Tactic') into an array
Public Function ConvertEnumToArray(value) As clsArray

    Dim convertedValue: convertedValue = Mid(value, 2, Len(value) - 2)
    convertedValue = replace(convertedValue, "'", """")
    Dim valueAsArray As New clsArray: valueAsArray.arr = convertedValue
    
    Set ConvertEnumToArray = valueAsArray
    
End Function

Public Function ConvertToVerboseCaption(word) As String

    Dim VerboseCaption As String
    Dim words() As String, i
    
    If word = "id" Then
        VerboseCaption = "ID"
    ElseIf RegExTest(word, "^[^ ]+ +[^ ]+$") Then
        ' Handle case where word has spaces already, assume it is already a verbose Caption
        VerboseCaption = word
    ElseIf InStr(word, "_") > 0 And Not RegExTest(word, "[A-Z]") Then
        ' Handle case where word uses underscores, replace underscores with spaces and capitalize each word
        words = Split(word, "_")
        For i = 0 To UBound(words)
            words(i) = UCase(Left(words(i), 1)) & LCase(Right(words(i), Len(words(i)) - 1))
        Next i
        VerboseCaption = Join(words, " ")
    ElseIf StrComp(word, LCase(word), vbTextCompare) <> 0 Then
        ' Handle case where every letter is lowercase
        VerboseCaption = UCase(Left(word, 1)) & LCase(Right(word, Len(word) - 1))
    ElseIf RegExTest(word, "([a-z]+)([A-Z][a-z]*)+") Then
        ' Handle camelCase or PascalCase, insert spaces before capital letters
        Dim separatedWords As New clsArray
        separatedWords.arr = SeparateWords(word)
        VerboseCaption = StrConv(separatedWords.JoinArr(" "), vbProperCase)
    
    Else
        ' Default case, return the original word
        VerboseCaption = StrConv(word, vbProperCase)

    End If
    
    VerboseCaption = CorrectID(VerboseCaption)
    ConvertToVerboseCaption = VerboseCaption

End Function

Public Function CorrectID(str)
    CorrectID = replace(str, "I D", "ID")
End Function


Function RegExTest(strInput, strPattern As String) As Boolean
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.pattern = strPattern
    RegExTest = objRegExp.Test(strInput)
End Function

Function RegExReplace(strInput, strPattern As String, strReplace As String) As String
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True
    objRegExp.pattern = strPattern
    RegExReplace = objRegExp.replace(strInput, strReplace)
End Function

Public Function GetReplacedTemplate(rs As Recordset, templateName, Optional skipFields = "", Optional directText = "")
    
    If rs.EOF Then Exit Function

    Dim TemplateContent
    If isFalse(directText) Then
        TemplateContent = GetTemplateContent(templateName)
    Else
        TemplateContent = directText
    End If

    Dim skipFieldsArr As New clsArray: skipFieldsArr.arr = skipFields
    Dim fld As field

    For Each fld In rs.fields
        If skipFieldsArr.count > 0 Then
            If skipFieldsArr.InArray(fld.Name) Then
                GoTo NextField:
            End If
        End If
        Dim fieldValue: fieldValue = rs.fields(fld.Name)
        Dim keyword: keyword = "[" & fld.Name & "]"
        If InStr(TemplateContent, keyword) > 0 Then
            If Not IsNull(fieldValue) Then
                TemplateContent = replace(TemplateContent, "[" & fld.Name & "]", fieldValue, , , vbBinaryCompare)
            Else
                MsgBox fld.Name & " can't be null", vbCritical + vbOKOnly
                Exit Function
            End If
        End If
NextField:
    Next fld
    
    GetReplacedTemplate = TemplateContent
    ''GetReplacedTemplate = replace(TemplateContent, "?", "...")
    ''GetReplacedTemplate = replace(TemplateContent, "?", "(e)")

End Function

Public Function SeparateWords(inputString) As String

    Dim outputString As String
    Dim i As Integer
    
    If Len(inputString) > 0 Then
        outputString = Left(inputString, 1)
    End If
    
    For i = 2 To Len(inputString)
        If Asc(Mid(inputString, i, 1)) >= 65 And Asc(Mid(inputString, i, 1)) <= 90 Then
            outputString = outputString & "," & LCase(Mid(inputString, i, 1))
        Else
            outputString = outputString & Mid(inputString, i, 1)
        End If
    Next i
    
    SeparateWords = outputString
    
End Function

Public Function ConvertToCustomTimestamp() As String
    Dim Timestamp As String
    
    Timestamp = Format(Now, "yyyymmddhhnnss")
    
    ConvertToCustomTimestamp = Timestamp
    
End Function

Public Function RemoveFirstAndLastCharacter(ByVal inputString As String) As String
    If Len(inputString) <= 2 Then
        RemoveFirstAndLastCharacter = ""
    Else
        RemoveFirstAndLastCharacter = Mid(inputString, 2, Len(inputString) - 2)
    End If
End Function

Public Function RemoveLastBracket(ByVal inputString As String) As String
    Dim regexPattern As String
    Dim regexMatch As Object
    
    ' Define the regex pattern to match the last "}" in the string
    regexPattern = "\}(\s*)$"
    
    ' Create a regex object
    Set regexMatch = CreateObject("VBScript.RegExp")
    
    With regexMatch
        .Global = True
        .Multiline = True
        .IgnoreCase = True
        .pattern = regexPattern
    End With
    
    ' Remove the last "}" from the input string
    RemoveLastBracket = regexMatch.replace(inputString, "")
End Function











