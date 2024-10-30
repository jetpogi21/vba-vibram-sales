Attribute VB_Name = "Workspace Test"
Option Compare Database
Option Explicit

Public Function WorkspaceTest()

    Dim wrkCurrent As DAO.Workspace, rs As DAO.Recordset, db As DAO.Database
    
    Set db = CurrentDb
    Set wrkCurrent = DBEngine.Workspaces(0)
    
    Dim fields(0) As String, fieldValues(0) As String
    fields(0) = "[Workspace]": fieldValues(0) = """Workspace"""
    
'    Set rs = db.OpenRecordset("tblWorkspaces")
    
    wrkCurrent.BeginTrans
'    rs.AddNew
'    rs!Workspace = "Workspaces"
'    rs.Update
    
    InsertAndLog "tblWorkspaces", fields, fieldValues
    
    If MsgBox("Save all changes?", vbQuestion + vbYesNo) = vbYes Then
        wrkCurrent.CommitTrans
    Else
        wrkCurrent.Rollback
    End If

    wrkCurrent.Close
    
    Set wrkCurrent = Nothing
    
End Function
