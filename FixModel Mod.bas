Attribute VB_Name = "FixModel Mod"
Option Compare Database
Option Explicit

Public Function FixModelCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Sub FixModel()
    ''in hero skills convert SkillType a and p to Active and Passive respectively
    Dim valueArr As New clsArray: valueArr.arr = "a,p"
    Dim updatedValueArr As New clsArray: updatedValueArr.arr = "Active,Passive"
    
    RunDatabaseUpdate valueArr, updatedValueArr, "tblHeroSkills", "SkillType"
    
    ''in cards CardType c,w,p,t
    valueArr.arr = "c,w,p,t"
    updatedValueArr.arr = "Character,Weapon,Power,Tactic"
    RunDatabaseUpdate valueArr, updatedValueArr, "tblCards", "CardType"
    
    ''in cards BattleStyle a,g,s
    valueArr.arr = "a,g,s"
    updatedValueArr.arr = "Attack,Guardian,Support"
    RunDatabaseUpdate valueArr, updatedValueArr, "tblCards", "BattleStyle"
    
End Sub

Private Function RunDatabaseUpdate(valueArr As clsArray, updatedValueArr As clsArray, tblName, fldName)
    
    Dim i As Integer
    Dim value, updateValue
    For i = 0 To valueArr.count - 1
        Debug.Print valueArr.count
        value = valueArr.items(i)
        updateValue = updatedValueArr.items(i)
        ''Debug.Print "UPDATE " & tblName & " SET " & fldName & " = " & EscapeString(updateValue) & " WHERE " & fldName & " = " & EscapeString(value)
        RunSQL "UPDATE " & tblName & " SET " & fldName & " = " & EscapeString(updateValue) & " WHERE " & fldName & " = " & EscapeString(value)
    Next i
    
End Function

Public Function FixTableAutonumbering(TableName As String, Optional iteration = 0)
    
    Dim PrimaryKey As String: PrimaryKey = GetPrimaryKeyFieldFromTable(TableName)
    ''Get the maximum id of TableName, PrimaryKey
    Dim MaxPrimaryKey: MaxPrimaryKey = Emax(TableName, PrimaryKey & " > 0", PrimaryKey)
    
    Dim sqlStr: sqlStr = "SELECT * FROM " & TableName & " WHERE " & PrimaryKey & " = " & MaxPrimaryKey
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim vTimestamp: vTimestamp = rs.fields("Timestamp")
    
    Dim fields As New clsArray: fields.arr = GetFields(TableName, PrimaryKey & ",CreatedBy,RecordImportID")
    Dim fieldValues As New clsArray
    
    Dim item
    Dim i As Integer: i = 0
    Dim j As Integer
    
    For j = 0 To iteration
        Set fieldValues = New clsArray
        For Each item In fields.arr
            fieldValues.Add rs.fields(item)
            i = i + 1
        Next item
        
        UpsertRecord TableName, fields, fieldValues
        Dim PrimaryKeyValue: PrimaryKeyValue = ELookup(TableName, "[Timestamp] = #" & SQLDate(vTimestamp) & "#", PrimaryKey, "[Timestamp]")
        
        ''RunSQL "DELETE FROM " & TableName & " WHERE " & PrimaryKey & " = " & PrimaryKeyValue
    Next j
    
    MsgBox "Autonumbering fixed: make sure to delete the extra records."
    
End Function

Public Sub Update_tblSeqModelFields_price()
    
    Dim sqlStr: sqlStr = "Select * from qrySeqModelFields WHERE ModelName Like ""*Count"" AND DatabaseFieldName = ""price"""
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        RunSQL "UPDATE tblSeqModelFields SET Expression = Null, IsGeneratedField = False, DefaultValue = 0 WHERE SeqModelFieldID = " & SeqModelFieldID
        rs.MoveNext
    Loop
    
    sqlStr = "Select * from qrySeqModelFields WHERE ModelName Like ""*Count"" AND DatabaseFieldName = ""total"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        SeqModelFieldID = rs.fields("SeqModelFieldID")
        RunSQL "UPDATE tblSeqModelFields SET Expression = ""[qty] * [price]"" WHERE SeqModelFieldID = " & SeqModelFieldID
        rs.MoveNext
    Loop
    
End Sub
