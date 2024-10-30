Attribute VB_Name = "NonSystemTable Mod"
Option Compare Database
Option Explicit

Public Function NonSystemTableCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            frm.AllowAdditions = False
            frm("ModelID").DisplayAsHyperlink = 2
            frm("ModelID").OnDblClick = "=frmNonSystemTables_ModelID_OnDblClick([Form])"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function frmNonSystemTables_ModelID_OnDblClick(frm As Form)

    Dim ModelID: ModelID = frm("ModelID")
    Open_frmModels ModelID
    
End Function

Public Function EnumerateNonSystemTables()
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySupplementalModels WHERE NOT IsSystemTable ORDER BY SupplementalModelID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    RunSQL "DELETE FROM tblNonSystemTables"
    
    Dim fields As New clsArray: fields.arr = "TableName,ModelID,CustomOrder"
    Dim fieldValues As New clsArray
    Dim i: i = 0
    Do Until rs.EOF
        Set fieldValues = New clsArray
        Dim TableName: TableName = rs.fields("TableName")
        Dim ModelID: ModelID = rs.fields("ModelID")
        If Not isFalse(TableName) Then
            TableName = replace(TableName, "qry", "tbl")
            fieldValues.Add TableName
            fieldValues.Add ModelID
            fieldValues.Add i
            Dim rs2 As Recordset: Set rs2 = ReturnRecordset(TableName)
            If Not rs2 Is Nothing Then
                If Not isPresent("tblNonSystemTables", "TableName = " & Esc(TableName)) Then
                    UpsertRecord "tblNonSystemTables", fields, fieldValues, "TableName = " & Esc(TableName)
                End If
            End If
        End If
        i = i + 1
        rs.MoveNext
    Loop
    
    Dim frm As Form: Set frm = GetForm("mainNonSystemTables")
    If Not frm Is Nothing Then
        frm("subform").Form.Requery
    End If
    
End Function

Public Function PurgeNonSystemTableData(Optional askConfirmation As Boolean = True, Optional DebugMode As Boolean = False)
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblNonSystemTables ORDER BY CustomOrder DESC, NonSystemTableID DESC"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If askConfirmation Then
        Dim resp: resp = MsgBox("There will purge all the data from the enumareted tables. Do you want to proceed?", vbYesNo)
        If resp = vbNo Then Exit Function
    End If
    
    Dim PrimaryKeyName
    Do Until rs.EOF
        Dim TableName As String: TableName = rs.fields("TableName")
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(TableName)
        
        If Not rs2 Is Nothing Then
            RunSQL "DELETE FROM " & TableName
            If TableName = "tblCustomers" Then
                Debug.Print "Loop went to tblCustomers"
            End If
            rs2.Close
            Set rs2 = Nothing
            Set rs2 = ReturnRecordset(TableName)
            Dim recordCount: recordCount = CountRecordset(rs2)
            If TableName = "tblCustomers" Then
                Debug.Print "tblCustomers' record count is: " & recordCount
            End If
            CopyToClipboard TableName
            Set rs2 = Nothing
            DeleteTableRelationships TableName
            deleteTableIfExists TableName
            
            If TableName = "tblCustomers" Then
                Set rs2 = ReturnRecordset(TableName)
                If Not rs2 Is Nothing Then
                    Debug.Print "tblCustomers was not deleted. Autonumbering was not reset"
                End If
            End If
            
        End If
        
        rs.MoveNext
    Loop
    
End Function

Public Function RecreateNonSystemTables()

    Dim sqlStr: sqlStr = "SELECT * FROM tblNonSystemTables ORDER BY CustomOrder,NonSystemTableID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim ModelID: ModelID = rs.fields("ModelID")
        DoCmd.OpenForm "frmModels", , , "ModelID = " & ModelID, , acHidden
        Dim frm As Form: Set frm = GetForm("frmModels")
        If Not frm Is Nothing Then
            CreateTableDef frm, False
        End If
        rs.MoveNext
    Loop
    
End Function

Private Function Open_frmModels(ModelID)

    DoCmd.OpenForm "frmModels", , , "ModelID = " & ModelID, , acHidden
    
End Function

Public Function CreateBlankDatabase(Optional DebugMode As Boolean = False)
    
    If Not DebugMode Then
        Dim resp: resp = MsgBox("This will close all other forms. Do you want to proceed?", vbYesNo)
        If resp = vbNo Then Exit Function
    End If
    
    CloseAllForms "frmCustomDashboard"
    
    Dim frm As Form: Set frm = GetForm("frmStartCreateBlankDatabase")
    
    If Not DebugMode Then
        Dim PasswordOverride: PasswordOverride = InputBox("Enter the password to create a blank database out of this file.")
        If PasswordOverride <> "iamhandsome" Then
            ShowError "Incorrect password. Blank database creation failed."
            If Not frm Is Nothing Then
                DoCmd.Close acForm, frm.Name, acSaveNo
            End If
            Exit Function
        End If
    End If
    
    Dim destinationPath As String
    ''Create a copy of the current database -> This will serve as a backup of the current file..
    If Not DebugMode Then CreateDatabaseCopy destinationPath
    ''Run PurgeNonSystemTableData to remove all table data and delete all the relationships and tables
    PurgeNonSystemTableData False
    ''Run RecreateNonSystemTables to re-create all the tables with autonumbering correctly reset
    RecreateNonSystemTables
    
    If Not DebugMode Then MsgBox "Database File Backup created at " & Esc(destinationPath)
    
    If Not frm Is Nothing Then
        DoCmd.Close acForm, frm.Name, acSaveNo
    End If

End Function

Public Sub Test_CreateBlankDatabase()
    
    CreateBlankDatabase True
    
    ''Create Seed data for tblProducts then Look if the first record has id of 1
    SeedTableAndValidateTheFirstID "Product"
    SeedTableAndValidateTheFirstID "Customer"
    SeedTableAndValidateTheFirstID "CustomerOrderMain"
    SeedTableAndValidateTheFirstID "CustomerOrder"
    
End Sub

Private Sub SeedTableAndValidateTheFirstID(Model)
    
    Dim frm As Form: Set frm = GetForm("frmCustomDashboard", True, False)
    If frm Is Nothing Then Exit Sub
    
    Dim ModelID: ModelID = ELookup("tblModels", "Model = " & Esc(Model), "ModelID")
    
    Dim TableName: TableName = GetTableName(Model, , , True)
    Dim PrimaryKey: PrimaryKey = GetPrimaryKeyFieldFromTable(TableName)
    
    Dim rs As Recordset: Set rs = ReturnRecordset(TableName)
    Dim recordCount: recordCount = CountRecordset(rs)
    If recordCount > 0 Then
        Debug.Print Esc(TableName) & " has RecordCount: " & recordCount & ", supposed to be 0"
    End If
    
    SeedTable frm, ModelID, , 3, False
    
    Dim FirstID: FirstID = ELookup(TableName, PrimaryKey & " > 0", PrimaryKey)
    
    If FirstID <> "1" Then
        Debug.Print "First " & PrimaryKey & " : " & FirstID & " is not equal to 1."
    End If
    
End Sub

Private Sub CreateDatabaseCopy(ByRef destinationPath As String)

    Dim CurrentTimestamp: CurrentTimestamp = GetFormattedTimestamp()
    
    Dim sourcePath As String: sourcePath = CurrentProject.path & "\" & CurrentProject.Name
    Dim AppName: AppName = "Warehousing & Quality Control"
    destinationPath = CurrentProject.path & "\" & AppName & " " & CurrentTimestamp & ".accdb"
    
    Dim fso As Object
    Dim sourceFile, destinationFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile sourcePath, destinationPath
     
End Sub






