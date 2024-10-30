Attribute VB_Name = "MockData Mod"

Option Compare Database
Option Explicit

Public Function MockDataCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Sub LoadTestData(tblNames, Optional instanceNumber = 0)
    
    Dim tblNameArr As New clsArray: tblNameArr.arr = tblNames
    
    Dim tblName
    Dim rs As Recordset
    Dim fields As New clsArray
    For Each tblName In tblNameArr.arr
    
        
        ''Check first if there's a table to be used
       Dim SourceTable: SourceTable = tblName
       SourceTable = SourceTable & "_test_" & instanceNumber
       
'On Error GoTo ErrHandler:
       Set rs = ReturnRecordset(SourceTable)
       
       If Not rs.EOF Then
            rs.MoveFirst
           RunSQL "DELETE FROM " & tblName
           Dim rs2 As Recordset: Set rs2 = ReturnRecordset(tblName)
           Dim fld As field
           
           Dim field
           Set fields = New clsArray
           For Each fld In rs2.fields
               fields.Add fld.Name
           Next fld
           
           Dim fieldValues As New clsArray
           
           Do Until rs.EOF
               Set fieldValues = New clsArray
               For Each field In fields.arr
                   fieldValues.Add rs.fields(field)
               Next field
               UpsertRecord tblName, fields, fieldValues
               rs.MoveNext
           Loop
           
       End If
    Next tblName
ErrHandler:
 If Err.Number = 3078 Then
    MsgBox Esc(SourceTable) & " does not exist."
    Exit Sub
 End If
 
End Sub


Public Sub SaveTestData(tblNames, Optional instanceNumber = 0)
    
    Dim tblNameArr As New clsArray: tblNameArr.arr = tblNames
    
    Dim tblName
    For Each tblName In tblNameArr.arr
        Dim rs As Recordset: Set rs = ReturnRecordset(tblName)
    
        Dim DestinationTable: DestinationTable = tblName
        DestinationTable = DestinationTable & "_test_" & instanceNumber
        
        CopyAndReplaceTable tblName, DestinationTable
    Next tblName

ErrHandler:
    If Err.Number = 3078 Then
       MsgBox Esc(tblName) & " does not exist."
       Exit Sub
    End If
 
End Sub

Public Function AddCompaniesToContacts()

    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblContacts")
    
    Do Until rs.EOF
        rs.Edit
        rs.fields("CompanyID") = GetRandomID("tblCompanies", "CompanyID")
        rs.Update
        rs.MoveNext
    Loop
    
End Function

Public Function InsertMockStaff()

    ''TABLE: tblInternalLeads Fields: InternalLeadID|InternalLeadName|Timestamp|CreatedBy|RecordImportID|BusinessUnit|IsSelected|IsStaff
    Dim i, fName, lName, InternalLeadName
    For i = 0 To 10
        fName = GetRandomID("tblMockData", "FirstName")
        lName = GetRandomID("tblMockData", "LastName")
        InternalLeadName = fName & " " & lName
        RunSQL "INSERT INTO tblInternalLeads (InternalLeadName,IsStaff) VALUES (" & EscapeString(InternalLeadName) & ",-1)"
    Next i
    
End Function

Public Function GetRandomCode(Optional length As Integer = 10) As String
    Dim chars As String
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    
    Dim result As String
    Dim i As Integer
    
    Randomize ' Initialize random Number generator
    
    For i = 1 To length
        result = result & Mid(chars, Int(Rnd() * Len(chars)) + 1, 1)
    Next i
    
    GetRandomCode = result
End Function

Sub Test_GetRandomCode()
    Dim code As String
    code = GetRandomCode() ' Generates a default 10-character code
    Debug.Print code
End Sub

Public Function AddMockDataTo_tblOppurtunities()
    
    ''TABLE: tblOppurtunities Fields: OppurtunityID|DealName|CompanyID|ContactID|InternalLeadID|Fee|CloseDate|Weighting|Timestamp|CreatedBy|RecordImportID|Comments|SectorID|StageID|StatusID
    ''TABLE: tblMockData Fields: MockDataID|Phrase|FirstName|LastName|LongText|EmailAddress|PhoneNumber|Address|Timestamp|CreatedBy|RecordImportID
    Dim recordSize: recordSize = 300
    Dim i
    Dim DealName, CompanyID, ContactID, InternalLeadID, Fee, CloseDate, Weighting, Timestamp, CreatedBy, RecordImportID, Comments, SectorID, StageID, StatusID, InvestmentArea, Country, DatePitched
    Dim fldArr As New clsArray: fldArr.arr = "DealName,CompanyID,ContactID,InternalLeadID,Fee,CloseDate,Weighting,Comments,SectorID,StageID,StatusID,InvestmentArea,Country,DatePitched"
    Dim fldValArr As New clsArray
    
    For i = 0 To recordSize - 1
    
        Set fldValArr = New clsArray
        
        DealName = GetRandomID("tblMockData", "Phrase")
        CompanyID = GetRandomID("tblCompanies", "CompanyID")
        ContactID = GetRandomID("tblContacts", "ContactID")
        InternalLeadID = GetRandomID("tblInternalLeads", "InternalLeadID")
        Fee = CLng(1000000 * Rnd)
        CloseDate = CDate(GetRandomFromRange(CLng(#1/1/2020#), CLng(#12/31/2022#)))
        If Rnd > 0.8 Then CloseDate = Null
        Weighting = GetRandomFromRange(0, 100) / 100
        Comments = GetRandomID("tblMockData", "LongText")
        SectorID = GetRandomID("tblSectors", "SectorID")
        StageID = GetRandomID("tblStages", "StageID")
        If StageID <> 1 Then
            DatePitched = CDate(GetRandomFromRange(CLng(#1/1/2020#), CLng(#12/31/2022#)))
        Else
            DatePitched = Null
        End If
        StatusID = GetRandomID("tblStatus", "StatusID")
        InvestmentArea = GetRandomID("tblInvestmentAreas", "InvestmentArea")
        Country = GetRandomID("tblMockData", "Country")
        
        fldValArr.Add EscapeString(DealName, "tblOppurtunities", "DealName")
        fldValArr.Add EscapeString(CompanyID, "tblOppurtunities", "CompanyID")
        fldValArr.Add EscapeString(ContactID, "tblOppurtunities", "ContactID")
        fldValArr.Add EscapeString(InternalLeadID, "tblOppurtunities", "InternalLeadID")
        fldValArr.Add EscapeString(Fee, "tblOppurtunities", "Fee")
        fldValArr.Add EscapeString(CloseDate, "tblOppurtunities", "CloseDate")
        fldValArr.Add EscapeString(Weighting, "tblOppurtunities", "Weighting")
        fldValArr.Add EscapeString(Comments, "tblOppurtunities", "Comments")
        fldValArr.Add EscapeString(SectorID, "tblOppurtunities", "SectorID")
        fldValArr.Add EscapeString(StageID, "tblOppurtunities", "StageID")
        fldValArr.Add EscapeString(StatusID, "tblOppurtunities", "StatusID")
        fldValArr.Add EscapeString(InvestmentArea, "tblOppurtunities", "InvestmentArea")
        fldValArr.Add EscapeString(Country, "tblOppurtunities", "Country")
        fldValArr.Add EscapeString(DatePitched, "tblOppurtunities", "DatePitched")
        
        RunSQL "INSERT INTO tblOppurtunities (" & fldArr.JoinArr(",") & ") VALUES (" & fldValArr.JoinArr(",") & ")"
        
    Next i
    
End Function

Public Function SetFieldValueClipboard()

    
    Dim fldArr As New clsArray, fld: fldArr.arr = "DealName,CompanyID,ContactID,InternalLeadID,Fee,CloseDate,Weighting,Comments,SectorID,StageID,StatusID"
    Dim tblName: tblName = "tblOppurtunities"
    
    Dim clipboardArr As New clsArray
    For Each fld In fldArr.arr
        clipboardArr.Add "fldValArr.Add EscapeString(" & fld & "," & EscapeString(tblName) & "," & EscapeString(fld) & ")"
    Next fld
    
    CopyToClipboard clipboardArr.JoinArr(vbCrLf)
    
End Function

Public Function GetRandomFromRange(lower As Double, upper As Double, Optional allowDecimal As Boolean = False) As Double
    Dim randValue As Double
    
    If allowDecimal Then
        ' Generate a random Number with up to 2 decimal places
        randValue = lower + Rnd * (upper - lower)
        GetRandomFromRange = Round(randValue, 2)
    Else
        ' Generate a random integer
        GetRandomFromRange = Int(lower + Rnd * (upper - lower + 1))
    End If
End Function

Public Function GetRandomID(tblName, tableID, Optional AdditionalFilter = "")

    Dim rs As DAO.Recordset
    Dim recordCount As Long
    Dim randomRecord As Long
    
    Dim filtersArr As New clsArray
    filtersArr.Add "NOT " & tableID & " IS NULL"
    If Not isFalse(AdditionalFilter) Then
        filtersArr.Add AdditionalFilter
    End If
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & tblName & " WHERE " & filtersArr.JoinArr(" AND "))
    
    If rs.EOF Then
        GetRandomID = Null
        Exit Function
    End If
    
    rs.MoveLast 'To get the count
    rs.MoveFirst
    recordCount = rs.recordCount - 1
    randomRecord = CLng((recordCount) * Rnd)

    rs.Move randomRecord
    
    GetRandomID = rs.fields(tableID)
     
End Function


