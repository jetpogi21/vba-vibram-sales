VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsThermalPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private vCounter As Long
Private vPageLimit As Long
Private vMargin As Long

Private Sub Class_Initialize()
    
    Dim fso As Object, filePath
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    filePath = "C:\Users\" & Environ$("username") & "\Desktop\blank.txt"

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(filePath, True, True)
    Fileout.Close
    
    Dim myFile As String
    myFile = filePath
    Open myFile For Output As #1
    
    vCounter = 0
    vPageLimit = 69
    vMargin = 0
    
End Sub

Property Let PageLimit(value As Long)
    vPageLimit = value
End Property

Property Get PageLimit() As Long
    PageLimit = vPageLimit
End Property

Public Sub PrintArray(arrs As clsArray)

    Dim arr As Variant, lineToPrint, isFirst
    
    isFirst = True
    
    For Each arr In arrs.arr
        If isFirst Then
            isFirst = False
            lineToPrint = arr
        Else
            lineToPrint = lineToPrint & " " & arr
        End If
        
    Next arr
    
    PrintLine lineToPrint
    
End Sub

Public Sub PrintLine(toBePrinted)
    
    Print #1, Space(vMargin) & toBePrinted
    vCounter = vCounter + 1
    If vCounter = vPageLimit Then
        Print #1, Chr(12)
        vCounter = 1
    End If

End Sub

Public Sub CloseFile()
    
    Close #1
    
End Sub

Property Let Margin(value As Long)
    
    vMargin = value
    
End Property

Public Sub NewPage()

    Print #1, Chr(12)
    vCounter = 1

End Sub

Public Sub PrintEmptyLines(LineQTY)
    
    Dim i As Integer
    
    For i = 0 To LineQTY
        PrintLine ""
    Next i
    
End Sub


