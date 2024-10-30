Attribute VB_Name = "Error Handling"
Option Compare Database
Option Explicit

Public Function IsInRange(arr() As String, i As Integer) As Boolean

On Error GoTo ErrHandle
    'Debug.Print Arr(i - 1)
    IsInRange = True
    Exit Function
    
ErrHandle:
    If Err.Number = 9 Then
        IsInRange = False
    End If
    
End Function

Public Function OpenFile(filePath As String) As Boolean

On Error GoTo ErrHandle

    Open filePath For Input As #1
    Exit Function
    
ErrHandle:

    If Err.Number = 55 Then
        Close #1
        Open filePath For Input As #1
    End If
    
End Function
