Attribute VB_Name = "SubformNavigation Mod"
Option Compare Database
Option Explicit

Public Function SubformNavigationCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function GoToNewRecord(frm As Object, Optional RemoveFocusArgs As String = "") As Boolean
    ' Ensure the subform is referenced correctly
    Dim OriginalFormName As String
    OriginalFormName = frm.Name
    
    If Not isFalse(RemoveFocusArgs) Then
        RemoveFocusFromSubform frm, RemoveFocusArgs
    End If
    
    frm("subform").SetFocus
    
    Set frm = frm("subform").Form
    
    DoCmd.GoToRecord , , acNewRec
    
    Set frm = Forms(OriginalFormName)
    SetNavigationData frm

End Function

Public Function GoToSpecificRecord(frm As Object, ByVal recordNumber As Long, Optional AllowAddition As Boolean = True, Optional Caption = "position")

    ' Ensure the subform is referenced correctly
    Dim OriginalFormName As String
    OriginalFormName = frm.Name
    Set frm = frm("subform").Form
    
    ' Use RecordsetClone to navigate the records
    Dim rs As DAO.Recordset
    Set rs = frm.RecordsetClone

    ' Check if the record Number is valid
    If recordNumber > 0 And recordNumber <= rs.recordCount Then
        ' Move to the first record
        rs.MoveFirst
        
        ' Loop through the records to find the specific record Number
        Dim i As Long
        For i = 1 To recordNumber - 1
            rs.MoveNext
        Next i
        
        ' Synchronize the form's record with the recordset
        If Not rs.EOF Then
            frm.Bookmark = rs.Bookmark
        Else
            ' Optionally, you can add a message box to indicate that the record does not exist
            MsgBox "The requested record does not exist."
        End If
    Else
        ' Optionally, you can add a message box to indicate that the record Number is invalid
        MsgBox "Invalid record Number."
    End If
    
    ' Clean up
    rs.Close
    Set rs = Nothing
    
    Set frm = Forms(OriginalFormName)
    SetNavigationData frm, AllowAddition, , Caption

End Function

Public Function txtRecordNumber_AfterUpdate(frm As Object, Optional AllowAddition = False, Optional Caption = "position")

    Dim txtRecordNumber: txtRecordNumber = frm("txtRecordNumber")
    
    If isFalse(txtRecordNumber) Then Exit Function
    
    GoToSpecificRecord frm, txtRecordNumber, , Caption
    
End Function

Public Function SetNavigationData(frm As Object, Optional AdditionAllowed As Boolean = True, Optional PerformRequery As Boolean = False, Optional Caption = "position")
    
    If PerformRequery Then frm("subform").Form.Requery
    ''set the txtRecordNumber to be the current record Number of the subform control
    Dim CurrentRecord: CurrentRecord = frm("subform").Form.CurrentRecord
    Dim rs As Recordset: Set rs = frm("subform").Form.RecordsetClone
    Dim recordCount As Long: recordCount = CountRecordset(rs)
    
    CurrentRecord = IIf(CurrentRecord > recordCount, "[NEW]", CurrentRecord)
    
    If Not AdditionAllowed And CurrentRecord = "[NEW]" Then
        CurrentRecord = 0
    End If
    
    frm("txtRecordNumber") = CurrentRecord
    
    frm("lbl_of_x_positions").Caption = " of " & recordCount & " " & PluralizeWord(Caption, recordCount)
    
End Function

Public Function DeleteSubformRecord(frm As Object, Optional RemoveFocusArgs As String = "")
    
    Dim OriginalFormName As String
    OriginalFormName = frm.Name
    
    If Not isFalse(RemoveFocusArgs) Then
        RemoveFocusFromSubform frm, RemoveFocusArgs
    End If
    
    frm("subform").SetFocus
    
    Set frm = frm("subform").Form
    
    If frm.NewRecord Then
        GoToPreviousRecord Forms(OriginalFormName), True, RemoveFocusArgs
        Exit Function
    End If
    
    On Error GoTo ErrHandler:
    DoCmd.RunCommand acCmdDeleteRecord

    Set frm = Forms(OriginalFormName)
    SetNavigationData frm
    Exit Function
    
ErrHandler:
    If Err.Number = 2501 Then
        Exit Function
    End If
    
End Function

Private Function RemoveFocusFromSubform(frm As Object, RemoveFocusArgs As String)
    
    Dim RemoveFocusArgsItem, RemoveFocusArgsArr As New clsArray: RemoveFocusArgsArr.arr = RemoveFocusArgs
    Dim FunctionName, SubformName
    
    Dim itemNumber As Long: itemNumber = 0
    For Each RemoveFocusArgsItem In RemoveFocusArgsArr.arr
        If IsEven(itemNumber) Then
            FunctionName = RemoveFocusArgsItem
        Else
            SubformName = RemoveFocusArgsItem
            Run FunctionName, frm("subform").Form(SubformName).Form
        End If
        itemNumber = itemNumber + 1
    Next RemoveFocusArgsItem

End Function

Public Function GoToNextRecord(frm As Object, Optional AllowAddition As Boolean = True, Optional RemoveFocusArgs As String = "", Optional Caption = "position")

    Dim OriginalFormName As String
    OriginalFormName = frm.Name
    
    If Not isFalse(RemoveFocusArgs) Then
        RemoveFocusFromSubform frm, RemoveFocusArgs
    End If
    
    frm("subform").SetFocus
    
    Set frm = frm("subform").Form
    Dim CurrentRecord: CurrentRecord = frm.CurrentRecord
    Dim rs As Recordset: Set rs = frm.RecordsetClone
    Dim recordCount As Long: recordCount = CountRecordset(rs)

    If CurrentRecord = recordCount And Not AllowAddition Then Exit Function
    
    On Error GoTo SkipMovement:
    DoCmd.GoToRecord , , acNext
    
SkipMovement:
    
    Set frm = Forms(OriginalFormName)
    SetNavigationData frm, AllowAddition, , Caption
    
End Function

Public Function GoToPreviousRecord(frm As Object, Optional AllowAddition As Boolean = True, Optional RemoveFocusArgs As String = "", Optional Caption = "position")

    ' Ensure the subform is referenced correctly
    Dim OriginalFormName As String
    OriginalFormName = frm.Name
    
     If Not isFalse(RemoveFocusArgs) Then
        RemoveFocusFromSubform frm, RemoveFocusArgs
    End If
    
    frm("subform").SetFocus

    Set frm = frm("subform").Form
    
    On Error GoTo SkipMovement:
    DoCmd.GoToRecord , , acPrevious
    
SkipMovement:
    
    Set frm = Forms(OriginalFormName)
    SetNavigationData frm, AllowAddition, , Caption

End Function

