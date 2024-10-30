Attribute VB_Name = "SQL Helper"

Option Compare Database
Option Explicit

Public Sub InsertItemsToTableField(ValueItems As String, fieldName As String, TableName As String)

    Dim fieldValueItems As New clsArray: fieldValueItems.arr = ValueItems
        
    Dim fields As New clsArray: fields.arr = fieldName
    Dim fieldValues As New clsArray
    
    Dim item
    Dim i As Integer: i = 0
    For Each item In fieldValueItems.arr
        Set fieldValues = New clsArray
        fieldValues.Add item
        UpsertRecord TableName, fields, fieldValues, fieldName & " = " & Esc(item)
        i = i + 1
    Next item
    
End Sub

Public Function UpsertRecord(tblName, fieldArr As clsArray, fieldValueArr As clsArray, Optional condition As String = "")
    
    Dim escapedFieldValues As New clsArray
    Dim fieldItem, i As Integer
    i = 0
    For Each fieldItem In fieldArr.arr
        escapedFieldValues.Add EscapeString(fieldValueArr.arr(i), tblName, fieldItem)
        i = i + 1
    Next fieldItem
    
    
    Dim DoUpdate As Boolean
    
    If isFalse(condition) Then
        DoUpdate = False
    Else
        DoUpdate = isPresent(tblName, condition)
    End If
    
    If DoUpdate Then
        ''Run Update
        
        Dim setStatements As New clsArray
        i = 0
        For Each fieldItem In fieldArr.arr
            setStatements.Add "[" & fieldItem & "] = " & escapedFieldValues.arr(i)
            i = i + 1
        Next fieldItem
        
        RunSQL "UPDATE " & tblName & " SET " & setStatements.JoinArr(",") & " WHERE " & condition
    Else
        ''Run Insert
        Dim parsedFieldArr As New clsArray
        For Each fieldItem In fieldArr.arr
            parsedFieldArr.Add "[" & fieldItem & "]"
        Next fieldItem
        RunSQL "INSERT INTO " & tblName & "(" & parsedFieldArr.JoinArr(",") & " ) VALUES (" & escapedFieldValues.JoinArr(",") & " )"
    End If
    
End Function


Public Function EscapeString(value, Optional tblName = "", Optional fieldName As Variant = "") As String

    If IsNull(value) Or value = "Null" Then
        EscapeString = "Null"
        Exit Function
    End If
    
    If tblName <> "" Then
        Dim defType As Object, FieldType
        If DoesPropertyExists(CurrentDb.TableDefs, tblName) Then
            Set defType = CurrentDb.TableDefs
        Else
            Set defType = CurrentDb.QueryDefs
        End If
    
        FieldType = defType(tblName).fields(fieldName).Type
        Select Case FieldType
            Case 10, 12:
                EscapeString = Chr(34) & replace(value, Chr(34), Chr(34) & Chr(34)) & Chr(34)
            Case 8:
                EscapeString = "#" & SQLDate(value) & "#"
            Case Else:
                EscapeString = replace(value, ",", ".")
        End Select
        
    Else
        EscapeString = Chr(34) & value & Chr(34)
    End If
    
End Function

Public Function ReplaceDoubleQuote(value) As String

    ReplaceDoubleQuote = replace(value, Chr(34), Chr(34) & Chr(34))
    
End Function

