Attribute VB_Name = "OneTimeFix Mod"
Option Compare Database
Option Explicit

''https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/create-table-statement-microsoft-access-sql

Public Function OneTimeFixCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function AddTaskNoteTotblTasks()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblTasks ADD COLUMN [TaskNote] MEMO"
    
    ''Fix the TaskNotes for each task using the GetTaskNotes(TaskID)
    RunSQL "UPDATE tblTasks SET TaskNote = GetTaskNotes(TaskID)"
    
    DoCmd.CopyObject BEPath, "Tasks", acQuery, "qryTasksToExport"
    
End Function

Public Function CopyEntityQueries()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    Dim sqlStr
    sqlStr = GetEntityMemberSQL
    
    ''Buyer,Seller,Tenant,Contact
    Dim catArr As New clsArray, i, qDef As QueryDef, sqlStr1
    catArr.arr = "Buyer,Seller,Tenant,Contact"
    
    For Each i In catArr.arr
        If i = "Contact" Then
            
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category],[Timestamp] FROM (" & sqlStr & ") temp WHERE EntityCategoryName = " & EscapeString(i)
            sqlStr1 = sqlStr1 & " ORDER BY [Timestamp]"
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category] FROM (" & sqlStr1 & ") temp2 GROUP BY [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category] ORDER BY Min([Timestamp])"
            
        ElseIf i = "Buyer" Then
        
            sqlStr1 = "SELECT TOP 20 [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category],[Timestamp],[Property Address] FROM (" & sqlStr & ") temp WHERE EntityCategoryName = " & EscapeString(i)
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],[Property Address],Website,[Timestamp] FROM (" & sqlStr1 & ") temp"
            sqlStr1 = sqlStr1 & " ORDER BY [Timestamp] ASC"
            
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],[Property Address],Website FROM (" & sqlStr1 & ") temp GROUP BY [Company Name],[Member],[Phone Number],[Email Address],[Property Address],Website ORDER BY Min([Timestamp])"
           
        Else
            
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category],[Timestamp],[Property Address] FROM (" & sqlStr & ") temp WHERE EntityCategoryName = " & EscapeString(i)
            sqlStr1 = sqlStr1 & " ORDER BY [Timestamp]"
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],[Property Address],Website FROM (" & sqlStr1 & ") temp ORDER BY [Timestamp]"
            ''makeQuery sqlStr1
            
        End If
        
        Set qDef = CurrentDb.QueryDefs("qryEntityQueries")
        qDef.sql = sqlStr1
        DoCmd.SetWarnings False
        DoCmd.CopyObject BEPath, i, acQuery, "qryEntityQueries"
        DoCmd.SetWarnings True
    Next i
    
End Function

Public Function GetEntityMemberSQL()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblEntities"
        .fields = "EntityID,EntityName,PhoneNumber,EmailAddress,Website,EntityCategoryName,ContactCategoryName,isSeller"
        .joins.Add GenerateJoinObj("tblEntityCategories", "EntityCategoryID")
        .joins.Add GenerateJoinObj("tblContactCategories", "ContactCategoryID", , , "LEFT")
        sqlStr = .sql
    End With
    
    Dim sqlStr1
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblPropertyEntities"
        .fields = "EntityName As [Company Name],MemberName As [Member],MemberPhoneNumber As [Phone Number] ,MemberEmailAddress As [Email Address]," & _
                  "StreetAddress As [Property Address],Website,ContactCategoryName As [Contact Category],EntityCategoryName,tblPropertyEntities.[Timestamp]"
        .joins.Add GenerateJoinObj("tblEntityMembers", "EntityID")
        .joins.Add GenerateJoinObj(sqlStr, "EntityID", "temp")
        .joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
        .OrderBy = "tblPropertyEntities.[Timestamp] DESC"
        sqlStr1 = .sql
    End With
    
    Dim sqlStr2
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblPropertyEntities"
        .AddFilter "isFavorite And isSeller"
        .fields = "EntityName As [Company Name],EntityName As [Member Name],PhoneNumber As [Phone Number],EmailAddress As [Email Address]," & _
                  "StreetAddress As [Property Address],Website,ContactCategoryName As [Contact Category],EntityCategoryName,tblPropertyEntities.[Timestamp]"
        .joins.Add GenerateJoinObj(sqlStr, "EntityID", "temp")
        .joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
        .OrderBy = "tblPropertyEntities.[Timestamp] DESC"
        sqlStr2 = .sql
    End With
    
    GetEntityMemberSQL = sqlStr1 & " UNION " & sqlStr2
    
End Function

Public Function AddCombinedOwnerToPropertyList()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function

    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [CombinedOwner] CHAR"
    RunSQL "UPDATE tblPropertyList SET CombinedOwner = JoinOwners(Owner1Name, Owner2Name, Owner3Name)"
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [FormattedTS] CHAR"
    
End Function

Public Function ModifyPropertyTempIndex()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "DROP INDEX UniqueProperty ON tblPropertyListTemp"
    RunSQLOnBackend BEPath, "CREATE UNIQUE INDEX UniqueProperty ON tblPropertyListTemp ([Street Address],[Owner 1 Name],[Owner 2 Name],[Owner 3 Name])"
    
End Function

Public Function DropPropertyUniqueIndex()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "DROP INDEX UniqueProperties ON tblPropertyList"
    
End Function

Public Function RunOneTimeFixes()

    Dim rs As Recordset
    Set rs = ReturnRecordset("select * from tblOneTimeFixes WHERE Not [Run] ORDER BY FunctionOrder,OneTimeFixID")
    
    Do Until rs.EOF
        Run rs.fields("FunctionName")
        RunSQL "UPDATE tblOneTimeFixes SET [Run] = -1 WHERE OneTimeFixID = " & rs.fields("OneTimeFixID")
        rs.MoveNext
    Loop
    
End Function

Public Function MakeBuyerStatusUnique()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "CREATE UNIQUE INDEX idxBuyerStatus ON tblBuyerStatus (BuyerStatus) WITH DISALLOW NULL"
    RunSQLOnBackend BEPath, "CREATE UNIQUE INDEX idxEntityNameIsSeller ON tblEntities (EntityName,PhoneNumber,EmailAddress,isSeller) WITH IGNORE NULL"

End Function

Public Function EditblEntitiesField()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    ''CustomerType,ToBeContacted & ToBeContactedDate
    RunSQLOnBackend BEPath, "ALTER TABLE tblEntities ADD COLUMN [CustomerType] CHAR, [ToBeContacted] BIT, [ToBeContactedDate] DATETIME"
    
End Function

Public Function AddPropertyAltLinkTotblPropertyList()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [PropertyAltLinks] MEMO"
        
End Function

Public Function AddExcludeFromReportTotblPropertyList()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [ExcludeFromReport] BIT"
    RunSQLOnBackend BEPath, "UPDATE tblPropertyList SET [ExcludeFromReport] = 0"
        
End Function

Public Function CopyActivityStudentsToBE()
    
    ''Change the tblName here"
    Dim tblName
    tblName = "tblActivityStudents"
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    DoCmd.CopyObject BEPath, tblName, acTable, tblName
    RunSQLOnBackend BEPath, "DELETE FROM " & tblName
    
End Function

Private Function GetFEAndBE(FEPath, BEPath)
    
    Dim ProjectPath, FEName, BEName
    ProjectPath = CurrentProject.path & "\"
    FEName = "PTS.accdb"
    BEName = "PTS Backend.accdb"
    
    FEPath = ProjectPath & FEName
    BEPath = ProjectPath & BEName
    
    If Environ("computername") <> "LAPTOP-4EL19IO4" Then
        ProjectPath = "Z:\MY PANDA APP"
        If Not DirectoryExists(ProjectPath) Then
            ProjectPath = "\\TRUENAS\database\MY PANDA APP"
            If Not DirectoryExists(ProjectPath) Then
                MsgBox "The database tables can't be linked to the backend file. The app will exit.", vbCritical
                DoCmd.Quit
                Exit Function
            End If
        End If
        BEPath = ProjectPath & "\PTS Backend.accdb"
    End If
        
End Function

Public Function FixBackendTable()

    ''Available Field types on tblFieldTypes, Use COUNTER for autonumber fields, FLOAT for DOUBLE, INTEGER (Long) AND SMALLINT
    ''CONSTRAINT MyTableConstraint UNIQUE & (FirstName, LastName, DateOfBirth));"
    ''SSN INTEGER CONSTRAINT MyFieldConstraint PRIMARY KEY"

End Function

Public Function RunSQLOnBackend(BEPath, sqlStr)

    Dim db As Database
    Set db = OpenDatabase(BEPath)

On Error GoTo Err_Handler:

    db.Execute sqlStr
    ''db.Close
    Exit Function
    
Err_Handler:
    
    MsgBox Err.description
    ''db.Close
    Exit Function
   
End Function

