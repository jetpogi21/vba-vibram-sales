Attribute VB_Name = "Reposition Helper"
Option Compare Database
Option Explicit


Public Sub RepositionControls(frm As Object, proportionArr As clsArray, controlArr As clsArray, x, y, totalWidth, Optional colSpaceWidth = 50)

    Dim proportionTotal, i, proportion, controlWidth
    proportionTotal = GetProportionTotal(proportionArr)
    
    For i = 0 To proportionArr.count - 1

        proportion = CDbl(proportionArr.arr(i)) / proportionTotal
        controlWidth = (totalWidth - ((proportionArr.count - 1) * colSpaceWidth * 2)) * proportion
        
        If controlArr.arr(i) <> "empty" Then
            frm(controlArr.arr(i)).Left = x
            frm(controlArr.arr(i)).Top = y
            frm(controlArr.arr(i)).Width = controlWidth
        End If
        
        x = x + (colSpaceWidth * 2) + controlWidth
       
    Next i
    
End Sub
