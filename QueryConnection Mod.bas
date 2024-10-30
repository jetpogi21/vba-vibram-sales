Attribute VB_Name = "QueryConnection Mod"
Option Compare Database
Option Explicit

Public Function GenerateQueryConFields(frm As Form)

    Dim QueryConnectionID, SummaryQueryID, QueryConnectionName, ChildKeys, ParentKeys, ConnectionType, Timestamp, CreatedBy, GroupBy
    
    QueryConnectionID = frm("QueryConnectionID")
    SummaryQueryID = frm("SummaryQueryID")
    QueryConnectionName = frm("QueryConnectionName")
    ChildKeys = frm("ChildKeys")
    ParentKeys = frm("ParentKeys")
    ConnectionType = frm("ConnectionType")
    Timestamp = frm("Timestamp")
    CreatedBy = frm("CreatedBy")
    GroupBy = frm("GroupBy")
        
    ''RunSQL "DELETE FROM tblQueryConnectionFields WHERE QueryConnectionID = " & QueryConnectionID
    
    Dim tblDef As Object, db As DAO.Database, fld As DAO.field
    Set db = CurrentDb
    
    If DoesPropertyExists(db.TableDefs, QueryConnectionName) Then
        Set tblDef = db.TableDefs(QueryConnectionName)
    Else
        Set tblDef = db.QueryDefs(QueryConnectionName)
    End If

    Dim fieldArr As New clsArray, fieldValues As New clsArray
    fieldArr.arr = "FieldTypeID,QueryConnectionID,QueryConnectionField,FieldCaption"
    
    For Each fld In tblDef.fields
        Set fieldValues = New clsArray
        fieldValues.Add fld.Type
        fieldValues.Add QueryConnectionID
        fieldValues.Add EscapeString(fld.Name)
        fieldValues.Add EscapeString(AddSpaces(fld.Name))
        
        If Not isPresent("tblQueryConnectionFields", "QueryConnectionID = " & QueryConnectionID & " And QueryConnectionField = " & EscapeString(fld.Name)) Then
        
            RunSQL "INSERT INTO tblQueryConnectionFields (" & fieldArr.JoinArr & ") VALUES (" & _
               fieldValues.JoinArr & " )"
               
        End If
        
    Next fld
    
    If DoesPropertyExists(Forms, "frmQueryConnections") Then
    
        If DoesPropertyExists(Forms("frmQueryConnections"), "subQueryConnectionFields") Then
            Forms("frmQueryConnections")("subQueryConnectionFields").Requery
        End If
        
    End If
    
    
End Function
