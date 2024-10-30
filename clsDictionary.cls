VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private dict As Object

Private Sub Class_Initialize()
    Set dict = CreateObject("Scripting.Dictionary")
End Sub

Public Function count() As Long
    count = dict.count
End Function

Public Function Keys() As Variant
    Keys = dict.Keys
End Function

Public Function values() As Variant
    values = dict.items
End Function

Public Function Add(key As Variant, value As Variant)
    If TypeOf value Is clsDictionary Then
        Set dict(key) = value
    Else
        dict(key) = value
    End If
End Function

Public Function Remove(key As Variant) As Boolean
    If dict.Exists(key) Then
        dict.Remove key
        Remove = True
    Else
        Remove = False
    End If
End Function

Public Function item(key As Variant) As Variant
    If dict.Exists(key) Then
        item = dict.item(key)
    Else
        item = Empty
    End If
End Function

Public Sub Clear()
    dict.RemoveAll
End Sub

Public Sub AddOrEdit(key As Variant, value As Variant)
    If Not dict.Exists(key) Then
        dict.Add key, value
    Else
        dict.item(key) = value
    End If
End Sub

Public Function ToFormatString(Optional inner As Boolean = False) As String
    Dim key As Variant
    Dim value As Variant
    Dim output As String
    Dim innerKey As Variant
    Dim innerDict As Object
    
    If dict.count = 0 Then
        If Not inner Then
            output = "{}"
        Else
            output = ""
        End If
        ToFormatString = output
        Exit Function
    End If
    
    If Not inner Then
        output = "{"
    End If
    
    For Each key In dict.Keys
        output = output & vbCrLf & "  " & key & ": "
        If TypeOf dict(key) Is clsDictionary Then
            output = output & "{"
            Set innerDict = dict.item(key)
            For Each innerKey In innerDict.Keys
                output = output & vbCrLf & "    " & innerKey & ": " & innerDict.item(innerKey) & ","
            Next innerKey
            output = Left(output, Len(output) - 1) & vbCrLf & "  },"
        Else
            output = output & dict.item(key) & ","
        End If
    Next key
    
    If Right(output, 1) = "," Then
        output = Left(output, Len(output) - 1) & vbCrLf
    End If
    
    If Not inner Then
        output = Left(output, Len(output) - 1) & vbCrLf & "}"
    End If
    
    ToFormatString = output
End Function


'Public Function ToFormatString() As String
'    Dim key As Variant
'    Dim value As Variant
'    Dim output As String
'    Dim innerKey As Variant
'    Dim innerDict As Object
'
'    output = "{"
'    For Each key In dict.Keys
'        output = output & vbCrLf & "  " & key & ": "
'        If TypeOf dict(key) Is clsDictionary Then
'            output = output & "{"
'            Set innerDict = dict.Item(key)
'            For Each innerKey In innerDict.Keys
'                output = output & vbCrLf & "    " & innerKey & ": " & innerDict.Item(innerKey) & ","
'            Next innerKey
'            output = Left(output, Len(output) - 1) & vbCrLf & "  },"
'        Else
'            output = output & dict.Item(key) & ","
'        End If
'    Next key
'    output = Left(output, Len(output) - 1) & vbCrLf & "}"
'    ToFormatString = output
'End Function
'
'Public Function ToInnerFormatString() As String
'    Dim key As Variant
'    Dim value As Variant
'    Dim output As String
'    Dim innerKey As Variant
'    Dim innerDict As Object
'
'    output = ""
'    For Each key In dict.Keys
'        output = output & vbCrLf & "  " & key & ": "
'        If TypeOf dict(key) Is clsDictionary Then
'            output = output & "{"
'            Set innerDict = dict.Item(key)
'            For Each innerKey In innerDict.Keys
'                output = output & vbCrLf & "    " & innerKey & ": " & innerDict.Item(innerKey) & ","
'            Next innerKey
'            output = Left(output, Len(output) - 1) & vbCrLf & "  }," ' remove trailing comma after inner dictionary
'        Else
'            output = output & dict.Item(key) & ","
'        End If
'    Next key
'    output = Left(output, Len(output) - 1) & vbCrLf ' remove trailing comma and add line break
'    ToInnerFormatString = output
'End Function








