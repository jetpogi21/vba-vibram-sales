VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private vSQLType As String
Private vSource As String
Private vFields As New clsArray
Private vFilter As String
Private vOrderBy As New clsArray
Private vJoins As New Collection
Private vSourceAlias As String
Private vSet As New clsArray
Private vInsertValues As New clsArray
Private vInsertSQL As String
Private vInsertFilterField As New clsArray
Private vGroupBy As New clsArray
Private VMakeTable As String
Private vInsertUseAsPlain As Boolean
Private vLastInsertID As Long
Private vHaving As String

Private Sub Class_Initialize()
    vSQLType = "SELECT"
End Sub

Property Let Having(value As String)
    vHaving = value
End Property

Property Let InsertFilterField(value As String)
    vInsertFilterField.arr = value
End Property

Property Let GroupBy(value As String)
    vGroupBy.arr = value
End Property

Property Let InsertUseAsPlain(value As Boolean)
    vInsertUseAsPlain = value
End Property

Property Let MakeTable(value As String)
    VMakeTable = value
End Property

Property Let SourceAlias(value As String)
    vSourceAlias = value
End Property

Public Function joins() As Collection

    Set joins = vJoins
    
End Function

Property Let SQLType(value As String)
    vSQLType = value
End Property

Property Get SQLType() As String
    SQLType = vSQLType
End Property

Property Let Source(value As String)
    vSource = value
End Property

Property Get Source() As String
    Source = vSource
End Property

Property Get LastInsertID() As Long
    LastInsertID = vLastInsertID
End Property

Property Let fields(value As Variant)
    
    vFields.arr = value
    
End Property

Property Let OrderBy(value As Variant)
    
    vOrderBy.arr = value

End Property

Property Let SetStatement(value As Variant)
    
    vSet.arr = value

End Property

Property Get fields() As clsArray
    
    Set fields = vFields
    
End Property

Property Get sql() As String
    
    Select Case vSQLType
    Case "SELECT":
        sql = GenerateSelectStatement()
    Case "UPDATE":
        sql = GenerateUpdateStatement()
    Case "INSERT":
        sql = GenerateInsertStatement()
    Case "DELETE":
        sql = GenerateDeleteStatement()
    Case "MAKE":
        sql = GenerateMakeStatement()
    Case Else:
        sql = ""
    End Select
    
End Property

Property Let insertValues(value As Variant)
    
    vInsertValues.arr = value

End Property

Property Let insertSQL(value As String)

    vInsertSQL = value
    
End Property

Private Function GenerateDeleteStatement() As String
    
    Dim sqlStr As String
    sqlStr = "DELETE"
    
    If vFields.count > 0 Then
        sqlStr = sqlStr & " " & vFields.JoinArr & " "
    Else
        sqlStr = sqlStr & " DISTINCTROW " & vSource & ".* "
    End If
    
    Dim openPar As String
    openPar = String(vJoins.count, "(")

    sqlStr = sqlStr & " FROM " & openPar & vSource & " "
    
    If vJoins.count > 0 Then
        sqlStr = sqlStr & GenerateJoinStatement
    End If
    
    If vFilter <> "" Then
        sqlStr = sqlStr & " WHERE " & vFilter
    End If
    
    GenerateDeleteStatement = sqlStr
    
End Function

Private Function GenerateInsertStatement() As String
    
    Dim sqlStr As String
    sqlStr = "INSERT INTO"

    sqlStr = sqlStr & " " & vSource & " "
    
    sqlStr = sqlStr & " (" & vFields.JoinArr & ") "
    
    ''Private vInsertValues As New clsArray
    ''Private vInsertSQL As String
    If vInsertValues.count > 0 Then
        sqlStr = sqlStr & " VALUES (" & vInsertValues.JoinArr & ")"
    ElseIf vInsertFilterField.count > 0 Then
        If vInsertUseAsPlain Then
            sqlStr = sqlStr & " SELECT " & vInsertFilterField.JoinArr & " FROM " & vInsertSQL
        Else
            sqlStr = sqlStr & " SELECT " & vInsertFilterField.JoinArr & " FROM " & "(" & vInsertSQL & ") tblTemp"
        End If
    Else
        sqlStr = sqlStr & " " & vInsertSQL
    End If

    GenerateInsertStatement = sqlStr
    
End Function

Private Function GenerateUpdateStatement() As String
    
    Dim sqlStr As String
    sqlStr = "UPDATE"
    
    Dim openPar As String
    openPar = String(vJoins.count, "(")

    sqlStr = sqlStr & " " & openPar & vSource
    
    If vJoins.count > 0 Then
        sqlStr = sqlStr & GenerateJoinStatement
    End If
    
    sqlStr = sqlStr & " SET " & vSet.JoinArr(",") & " "
    
    If vFilter <> "" Then
        sqlStr = sqlStr & " WHERE " & vFilter
    End If
    
    GenerateUpdateStatement = sqlStr
    
End Function

Private Function GenerateJoinStatement() As String

    ''Private vSource As String
    ''Private vAlias As String
    ''Private vJoinType As String
    ''Private vRightFields As New clsArray
    ''Private vLeftFields As New clsArray
    
    Dim vJoin As Object, joinStatementArr As clsArray
    Dim joinStr As String, i As Integer
    
    For Each vJoin In vJoins
        
        Set joinStatementArr = New clsArray
        
        joinStr = joinStr & " " & vJoin.JoinType & " JOIN "
        If vJoin.Alias <> "" Then
            joinStr = joinStr & "( " & vJoin.Source & " ) AS " & vJoin.Alias & " ON "
        Else
            joinStr = joinStr & " " & vJoin.Source & " ON "
        End If
        
        For i = 0 To vJoin.LeftFields.count - 1
            Dim linkStr As String
            
            If vSourceAlias <> "" Then
                linkStr = vSourceAlias & "." & vJoin.LeftFields.items(i) & " = "
            Else
                linkStr = vSource & "." & vJoin.LeftFields.items(i) & " = "
            End If
            
            If vJoin.Alias <> "" Then
                linkStr = linkStr & vJoin.Alias & "." & vJoin.RightFields.items(i)
            Else
                linkStr = linkStr & vJoin.Source & "." & vJoin.RightFields.items(i)
            End If

            joinStatementArr.Add "(" & linkStr & ")"
        Next i

        joinStr = joinStr & joinStatementArr.JoinArr(" AND ") & ") "
        
    Next vJoin
    
    GenerateJoinStatement = joinStr
    
End Function

Private Function GenerateSelectStatement() As String
    
    Dim sqlStr As String
    sqlStr = "SELECT"
    
    ''Fields
    If vFields.count > 0 Then
        sqlStr = sqlStr & " " & vFields.JoinArr(",")
    Else
        sqlStr = sqlStr & " *"
    End If

    Dim openPar As String
    openPar = String(vJoins.count, "(")

    If vSourceAlias = "" Then
        sqlStr = sqlStr & " FROM " & openPar & vSource
    Else
        sqlStr = sqlStr & " FROM " & openPar & "(" & vSource & ") AS " & vSourceAlias
    End If
    
    If vJoins.count > 0 Then
        sqlStr = sqlStr & GenerateJoinStatement
    End If
    
    If vFilter <> "" Then
        sqlStr = sqlStr & " WHERE " & vFilter
    End If
    
    If vGroupBy.count > 0 Then
        sqlStr = sqlStr & " GROUP BY " & vGroupBy.JoinArr(",")
    End If
    
    ''Having
    If vHaving <> "" Then
        sqlStr = sqlStr & " HAVING " & vHaving
    End If
    
    ''Order By
    If vOrderBy.count > 0 Then
        sqlStr = sqlStr & " ORDER BY " & vOrderBy.JoinArr(",")
    End If
    
    GenerateSelectStatement = sqlStr
    
End Function

Private Function GenerateMakeStatement() As String
    
    ''deleteTableIfExists VMakeTable
    
    Dim sqlStr As String
    sqlStr = "SELECT"
    
    ''Fields
    If vFields.count > 0 Then
        sqlStr = sqlStr & " " & vFields.JoinArr(",")
    Else
        sqlStr = sqlStr & " *"
    End If
    
    sqlStr = sqlStr & " INTO " & VMakeTable

    Dim openPar As String
    openPar = String(vJoins.count, "(")

    If vSourceAlias = "" Then
        sqlStr = sqlStr & " FROM " & openPar & vSource
    Else
        sqlStr = sqlStr & " FROM " & openPar & "(" & vSource & ") AS " & vSourceAlias
    End If
    
    If vJoins.count > 0 Then
        sqlStr = sqlStr & GenerateJoinStatement
    End If
    
    If vFilter <> "" Then
        sqlStr = sqlStr & " WHERE " & vFilter
    End If
    
    If vGroupBy.count > 0 Then
        sqlStr = sqlStr & " GROUP BY " & vGroupBy.JoinArr(",")
    End If
    
    ''Order By
    If vOrderBy.count > 0 Then
        sqlStr = sqlStr & " ORDER BY " & vOrderBy.JoinArr(",")
    End If
    
    GenerateMakeStatement = sqlStr
    
End Function

Public Function AddFilter(value, Optional Connector As String = "AND", Optional vEnclosureType As Long = 0)

    If vFilter = "" Then
        vFilter = EncloseValue(value, vEnclosureType)
    Else
        vFilter = vFilter & " " & Connector & " " & EncloseValue(value, vEnclosureType)
    End If
    
End Function

Private Function EncloseValue(value, vEnclosureType As Long) As String
    Select Case vEnclosureType
    Case 0:
        EncloseValue = value
    Case 1:
        EncloseValue = "(" & value
    Case 2:
        EncloseValue = value & ")"
    Case 3:
        EncloseValue = "(" & value & ")"
    End Select
End Function

Public Function Recordset() As Recordset

    Set Recordset = CurrentDb.OpenRecordset(sql())
    
End Function

Public Function Run() As Variant
    
    Run = RunSQL(sql())
    If vSQLType = "INSERT" Then
        vLastInsertID = CurrentDb.OpenRecordset("SELECT @@IDENTITY")(0)
    End If

End Function

Public Function ResetFilter()
    
    vFilter = ""
    
End Function

Public Function count() As Long

    Dim rs As Recordset
    Set rs = ReturnRecordset(sql())
    
    count = rs.recordCount
    
End Function


