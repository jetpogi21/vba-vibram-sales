Attribute VB_Name = "Lookups"
Option Compare Database
Option Explicit

Public Function isPresent(tblName, filterStr, Optional returnFalseOnError As Boolean = False) As Boolean

On Error GoTo ErrHandler:
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT *  FROM " & tblName & " WHERE " & filterStr)
    
    isPresent = Not rs.EOF
    
    Exit Function
ErrHandler:
    If returnFalseOnError Then
        isPresent = False
        Exit Function
    End If
    
    ShowError Err.Number & " - " & Err.description
    
End Function

Public Function Elookups(tblName, filterStr, fldName, Optional orderStr As String)
    
    Dim rs As Recordset
    Dim sqlStr As String
    sqlStr = "SELECT * FROM " & tblName & " WHERE " & filterStr
    
    If orderStr <> "" Then
        sqlStr = sqlStr & " ORDER BY " & orderStr
    End If
    
    Set rs = CurrentDb.OpenRecordset(sqlStr)
    
    If rs.EOF Then
        Elookups = ""
    Else
        Dim values As New clsArray
        Do Until rs.EOF
            Dim fieldValue: fieldValue = rs.fields(fldName)
            If Not isFalse(fieldValue) Then
                values.Add fieldValue, True
            End If
            rs.MoveNext
        Loop
        
        Elookups = values.JoinArr(",")
    End If
    
    Exit Function
    
End Function

Public Function ELookup(tblName, filterStr, fldName, Optional orderStr As String) As String
    
    Dim rs As Recordset
    Dim sqlStr As String
    sqlStr = "SELECT * FROM " & tblName & " WHERE " & filterStr
    
    If orderStr <> "" Then
        sqlStr = sqlStr & " ORDER BY " & orderStr
    End If
    
    Set rs = CurrentDb.OpenRecordset(sqlStr)
    
    If rs.EOF Then
        ELookup = ""
    Else
'On Error GoTo ErrHandler:
        If isFalse(rs.fields(fldName)) Then
            ELookup = ""
            Exit Function
        End If
        ELookup = rs.fields(fldName)
    End If
    
    Exit Function

'ErrHandler:
'    LogError Err.Number, Err.Description, "ELookup", , True
'    ELookup = ""

End Function

Public Function ELookupDate(tblName As String, filterStr As String, fldName As String, Optional orderStr As String) As Date

    Dim rs As Recordset
    Dim sqlStr As String
    sqlStr = "SELECT * FROM " & tblName & " WHERE " & filterStr
    
    If orderStr <> "" Then
        sqlStr = sqlStr & " ORDER BY " & orderStr
    End If
    
    Set rs = CurrentDb.OpenRecordset(sqlStr)
    
    If rs.EOF Then
        ELookupDate = #1/1/2100#
    Else
        If isFalse(rs.fields(fldName)) Then
            ELookupDate = #1/1/2100#
            Exit Function
        End If
        ELookupDate = SQLDate(rs.fields(fldName))
    End If

End Function

Public Function ReturnRecordset(sqlStr) As Recordset
   
On Error GoTo ErrHandler:
    Set ReturnRecordset = CurrentDb.OpenRecordset(sqlStr)
    Exit Function

    Exit Function
ErrHandler:
    If Err.Number = 3078 Then
        Exit Function
    End If

End Function

Public Function ESum(sqlStr As String, fieldName As String) As Double
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(sqlStr)
    
    If rs.EOF Then
        ESum = 0
        Exit Function
    End If
    
    If isFalse(rs.fields(fieldName)) Then
        ESum = 0
        Exit Function
    Else
        ESum = rs.fields(fieldName)
    End If
    
End Function

Public Function ESum2(tblName As String, filterStr, fieldName As String) As Double
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT SUM(" & fieldName & ") As SumOfRecord FROM " & tblName & " WHERE " & filterStr)
    
    If rs.EOF Then
        ESum2 = 0
        Exit Function
    End If
    
    If isFalse(rs.fields("SumOfRecord")) Then
        ESum2 = 0
        Exit Function
    Else
        ESum2 = rs.fields("SumOfRecord")
    End If
    
End Function

Public Function Emax(tblName As String, filterStr, fieldName As String) As Double
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Max(" & fieldName & ") As MaxValue FROM " & tblName & " WHERE " & filterStr)
    
    If rs.EOF Then
        Emax = 0
        Exit Function
    End If
    
    If isFalse(rs.fields("MaxValue")) Then
        Emax = 0
        Exit Function
    Else
        Emax = rs.fields("MaxValue")
    End If
    
End Function

Public Function Emin(tblName As String, filterStr, fieldName As String) As Double
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Min(" & fieldName & ") As MinValue FROM " & tblName & " WHERE " & filterStr)
    
    If rs.EOF Then
        Emin = 0
        Exit Function
    End If
    
    If isFalse(rs.fields("MinValue")) Then
        Emin = 0
        Exit Function
    Else
        Emin = rs.fields("MinValue")
    End If
    
End Function

Public Function ECount(tblName, filterStr As String) As Double
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT COUNT(*) As CountOfRecord FROM " & tblName & " WHERE " & filterStr)
    
    If rs.EOF Then
        ECount = 0
        Exit Function
    End If
    
    If isFalse(rs.fields("CountOfRecord")) Then
        ECount = 0
        Exit Function
    Else
        ECount = rs.fields("CountOfRecord")
    End If
    
End Function

