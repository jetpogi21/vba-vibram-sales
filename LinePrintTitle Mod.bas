Attribute VB_Name = "LinePrintTitle Mod"
Option Compare Database
Option Explicit

Public Function LinePrintTitleCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Private Function AlignText(inputStr As String, ColumnWidth As Long, alignment) As String

    Select Case alignment
            Case "L":
                AlignText = Left(inputStr & Space(ColumnWidth), ColumnWidth)
            Case "R":
                AlignText = Right(Space(ColumnWidth) & inputStr, ColumnWidth)
            Case "C":
                AlignText = CenterString(inputStr, ColumnWidth)
        End Select

End Function

Public Function RenderLine(LinePrintTitle, prntObj As clsThermalPrinting, ParamArray Args() As Variant)

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryLinePrintColumns WHERE LinePrintTitle = " & EscapeString(LinePrintTitle) & " ORDER BY ColumnOrder")
    
    Dim LinePrintColumnID, ColumnValue, ColumnOrder, FromRecordset, alignment, ColumnType, ColumnWidth As Long, LinePrintTitleID, FillWidth, isExpression
    Dim LineStr, inputStr As String, i, evalStr, rs2 As Recordset, txtValue As String
    
    LineStr = ""
    Do Until rs.EOF
        
        LinePrintColumnID = rs.fields("LinePrintColumnID")
        ColumnValue = rs.fields("ColumnValue")
        ColumnOrder = rs.fields("ColumnOrder")
        FromRecordset = rs.fields("FromRecordset")
        alignment = rs.fields("Alignment")
        ColumnType = rs.fields("ColumnType")
        ColumnWidth = rs.fields("ColumnWidth")
        LinePrintTitleID = rs.fields("LinePrintTitleID")
        FillWidth = rs.fields("FillWidth")
        isExpression = rs.fields("IsExpression")
        
        If isExpression Then
            For i = 0 To UBound(Args)
                evalStr = Args(i)
                ColumnValue = replace(ColumnValue, "[" & i & "]", evalStr)
            Next i
        End If
        
        If IsNull(ColumnValue) Then
            ColumnValue = Space(ColumnWidth)
        End If
        
        If FromRecordset Then
            Set rs2 = Args(0)
            ColumnValue = rs2.fields(ColumnValue)
        End If
        
        Select Case ColumnType
            Case "Currency":
                ColumnValue = Format(ColumnValue, "Standard")
        End Select
        
        If FillWidth Then
            ColumnValue = String(ColumnWidth, ColumnValue)
            alignment = "L"
        End If
        
        txtValue = ColumnValue
        inputStr = AlignText(txtValue, ColumnWidth, alignment)
        
        LineStr = LineStr & inputStr
        
        rs.MoveNext
        
    Loop
    
    prntObj.PrintLine LineStr
    
End Function
