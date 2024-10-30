VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private vArray() As String

Property Get arr() As Variant

    arr = vArray
    
End Property

Property Let arr(arrValue As Variant)
    
    If IsArray(arrValue) Then
        vArray = arrValue
    Else
        vArray = Split(arrValue, ",")
    End If
    
End Property

Public Function Add(arrValue As Variant, Optional Unique As Boolean = False)
    
    If Unique Then
        If IsInArray(arrValue, vArray) Then Exit Function
    End If
    
    Dim CurrentCount As Long
    CurrentCount = count
    
    ReDim Preserve vArray(CurrentCount)
    
    If IsNull(arrValue) Then arrValue = "Null"
    vArray(CurrentCount) = arrValue

End Function

Public Function count() As Long

On Error GoTo ErrHandler:
    count = UBound(vArray) + 1
    Exit Function
ErrHandler:
    Select Case Err.Number
    Case 9:
        count = 0
    End Select
    
End Function

Public Function JoinArr(Optional Delimiter As String = ",") As String
    JoinArr = Join(vArray, Delimiter)
End Function

Public Function NewLineJoin() As String
    NewLineJoin = Join(vArray, vbNewLine)
End Function

Public Function items(Index As Integer) As String
    items = vArray(Index)
End Function

Public Function InArray(value) As Boolean
    InArray = IsInArray(value, vArray)
End Function

Public Function clearArr()

    Erase vArray
    
End Function

Public Function EscapeItems()
Dim i As Long
For i = LBound(vArray) To UBound(vArray)
vArray(i) = """" & vArray(i) & """"
Next i
End Function


