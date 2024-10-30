Attribute VB_Name = "SummaryQuery Mod"
Option Compare Database
Option Explicit

Public Function CreateQuery(frm As Form)
    
    Dim SummaryQueryID, MainTable, QueryName
    
    SummaryQueryID = frm("SummaryQueryID")
    MainTable = frm("MainTable")
    QueryName = frm("QueryName")
    
    If ExitIfTrue(IsNull(SummaryQueryID), "Please select a record below..") Then Exit Function

    Dim rs As Recordset, rs2 As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblQueryConnections WHERE SummaryQueryID = " & SummaryQueryID)
    
    Dim QueryConnectionID, QueryConnectionName, ChildKeys, ParentKeys, ConnectionType, GroupBy
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, fieldsArray As New clsArray, mainFieldsArray As New clsArray
    Dim sqlObj2 As clsSQL, sqlStr2, fldAliases As New clsArray
    Dim i As Integer
    i = 0
    
    mainFieldsArray.Add concat(MainTable, ".*")
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = MainTable
    End With
  
    Do Until rs.EOF
    
        QueryConnectionID = rs.fields("QueryConnectionID")
        QueryConnectionName = rs.fields("QueryConnectionName")
        ChildKeys = rs.fields("ChildKeys")
        ParentKeys = rs.fields("ParentKeys")
        ConnectionType = rs.fields("ConnectionType")
        GroupBy = rs.fields("GroupBy")
        
        ''Get the Fields
        Set rs2 = ReturnRecordset("SELECT * FROM tblQueryConnectionFields WHERE QueryConnectionID = " & QueryConnectionID & _
                                  " AND AggregateFunction IS NOT NULL")
        Set fieldsArray = New clsArray
        
        fieldsArray.Add GroupBy
        
        Dim fldName, fldAlias
        Do Until rs2.EOF
        
            fieldsArray.Add concat("CdblNz(", rs2.fields("AggregateFunction"), "([", rs2.fields("QueryConnectionField"), "])) AS ", _
                    rs2.fields("AggregateFunction"), "Of", rs2.fields("QueryConnectionField"))
            
            If IsNull(rs2.fields("FieldAlias")) Then
                fldAlias = rs2.fields("QueryConnectionField")
            Else
                fldAlias = rs2.fields("FieldAlias")
            End If
            
            fldAliases.Add fldAlias
            
            fldName = concat("temp_", i, "!", rs2.fields("AggregateFunction"), "Of", rs2.fields("QueryConnectionField"), " AS ", fldAlias)
            
            mainFieldsArray.Add fldName
            
            Debug.Print fldName
            
            rs2.MoveNext
        Loop
        
        ''Build the SQL of the QueryConnections
        Set sqlObj2 = New clsSQL
        With sqlObj2
            .Source = QueryConnectionName
            .fields = fieldsArray.JoinArr
            .GroupBy = GroupBy
            sqlStr2 = .sql
        End With
        
        sqlObj.joins.Add GenerateJoinObj(sqlStr2, ParentKeys, "temp_" & i, ChildKeys, ConnectionType)
        
        i = i + 1
        
        rs.MoveNext
    
    Loop
    
    ''Account for the additional query fields
    Set rs = ReturnRecordset("SELECT * FROM tblAdditionalQueryFields WHERE SummaryQueryID = " & SummaryQueryID)
    Do Until rs.EOF
        mainFieldsArray.Add rs.fields("FieldExpression") & " AS " & rs.fields("FieldAlias")
        rs.MoveNext
    Loop
    
    sqlObj.fields = mainFieldsArray.JoinArr
    sqlStr = sqlObj.sql
    
    Dim db As DAO.Database, qDef As DAO.QueryDef
    Set db = CurrentDb
    
    If DoesPropertyExists(db.QueryDefs, QueryName) Then
        Set qDef = db.QueryDefs(QueryName)
        qDef.sql = sqlStr
    Else
        db.CreateQueryDef QueryName, sqlStr
    End If

End Function
