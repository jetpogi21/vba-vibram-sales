Attribute VB_Name = "Backend Mod"
Option Compare Database
Option Explicit

Public Function AlterBackendTable(filePath)
    

    Dim db As Database
    Set db = OpenDatabase(filePath)
    
    On Error GoTo Err_Handler:
    db.Execute "ALTER TABLE tblPropertyStatus ADD COLUMN IsShownOnFavorite BIT"
UpdateNow:
    'db.Execute "UPDATE tblPropertyStatus SET IsShownOnFavorite = -1"
    Exit Function
    
Err_Handler:
    If Err.Number = 3380 Then
        GoTo UpdateNow
    End If
    
End Function
