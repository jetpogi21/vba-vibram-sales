Attribute VB_Name = "BackendProject Mod"
Option Compare Database
Option Explicit

Public Function BackendProjectCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            ''ResizeSubforms frm
            frm("cmdProjectPath").OnClick = "=SelectDEDirectory([Form],""ProjectPath"",""SetAppName"")"
            frm("pgBackendNPMPackages").Caption = "Backend NPM Packages"
            frm("pgSeqModelRelationships").Caption = "Relationships"
            frm.PopUp = False
            
            frm.OnCurrent = "=frmBackendProjects_onCurrent([Form])"
            
            frm("listBackendProjectActions").Height = GetBottom(frm("tabCtl")) - frm("listBackendProjectActions").Top
            frm("listBackendProjectActions").VerticalAnchor = acVerticalAnchorBoth
            
        Case 5: ''Datasheet Form
            frm.AllowAdditions = False
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Private Function ResizeSubforms(frm As Form)

    Dim ctl As control
    
    Dim tabCtl As control
    Set tabCtl = frm("tabCtl")
    For Each ctl In frm.controls
        If ctl.ControlType = acSubform Then
            ctl.Width = tabCtl.Width - 250
        End If
    Next ctl
    
End Function

Public Function frmBackendProjects_onCurrent(frm As Form)
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    ''subSeqModelRelationships, LeftModelID, RightModelID
    Dim sqlStr: sqlStr = "SELECT SeqModelID,ModelName FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & " ORDER BY ModelName"
    
    frm("subSeqModelRelationships").Form.controls("LeftModelID").RowSource = sqlStr
    frm("subSeqModelRelationships").Form.controls("RightModelID").RowSource = sqlStr
    frm("subSeqModelRelationships").Form.controls("DeclareInModel").RowSource = sqlStr
    frm("subSeqModelRelationships").Form.controls("Through").RowSource = sqlStr
    
    Dim NextFontID: NextFontID = ELookup("tblNextFonts", "FontName = " & Esc("Roboto"), "NextFontID")
    
    If Not isFalse(NextFontID) Then
        frm("subAppConfigs").Form("NextFontID").DefaultValue = "=" & NextFontID
    End If
    
    
End Function

Public Function CreateBackendDirectory(frm As Form)

    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim shellArr As New clsArray
    shellArr.Add "cmd.exe /k mkdir " & Esc(ProjectPath)
    
    Call Shell(shellArr.JoinArr(" & "))
    
End Function

Public Function NodeInit(frm As Form)

    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /k cd /d " & Esc(ProjectPath)
    shellArr.Add "npm init -y"
    Call Shell(shellArr.JoinArr(" & "))
    
End Function

Public Function InstalNPMPackages(frm As Form)

    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    
    ''Get all the NPM packages -->
    ''TABLE: qryBackendNPMPackages Fields: BackendNPMPackageID|BackendProjectID|NPMPackageID|Timestamp|CreatedBy
    ''RecordImportID|NPMPackage|Location|DevOnly
    Dim NPMPackages As New clsArray: NPMPackages.arr = Elookups("qryBackendNPMPackages", "BackendProjectID = " & BackendProjectID & " AND Location = ""Backend""" & _
        "AND NOT DevOnly", "NPMPackage", "BackendNPMPackageID")
    Dim NPMPackagesDevOnly As New clsArray: NPMPackagesDevOnly.arr = Elookups("qryBackendNPMPackages", "BackendProjectID = " & BackendProjectID & " AND Location = ""Backend"" AND " & _
        "DevOnly", "NPMPackage", "BackendNPMPackageID")
    
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /k cd /d " & Esc(ProjectPath) ''CD into the VirtualEnvironmentPath
    
    If NPMPackages.count > 0 Then
        shellArr.Add "npm install " & NPMPackages.JoinArr(" ")
    End If
    
    If NPMPackagesDevOnly.count > 0 Then
        shellArr.Add "npm install --save-dev " & NPMPackagesDevOnly.JoinArr(" ")
    End If
    
    Call Shell(shellArr.JoinArr(" & "))
    
End Function

Public Function ModifyPackageJson(frm As Form)
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    ''TABLE: tblBackendProjects Fields: BackendProjectID|ProjectPath|Timestamp|CreatedBy|RecordImportID|AppName
    Dim AppName: AppName = frm("AppName")
    
    Dim filePath: filePath = ProjectPath & "package.json"
    
    ' Read the JSON file as text
    Open filePath For Input As #1
    Dim jsonText: jsonText = Input$(LOF(1), #1)
    Close #1
    
    Dim jsonObj As Object: Set jsonObj = JsonConverter.ParseJson(jsonText)
    
    ' Replace the name and type property
    jsonObj("name") = AppName
    jsonObj("type") = "commonjs"
    
    Dim scriptsObj As Object: Set scriptsObj = jsonObj("scripts")
    scriptsObj("start") = "nodemon dist/index.js"
    scriptsObj("build") = "tsc -p . -w"
    scriptsObj("dev") = "nodemon index.ts"
    
    ' Convert the dictionary object back to JSON text
    jsonText = JsonConverter.ConvertToJson(jsonObj, Whitespace:=2)
    
    ' Write the updated JSON text back to the file
    Open filePath For Output As #1
    Print #1, jsonText
    Close #1

End Function

Public Function CreateENVFile(frm As Object, Optional BackendProjectID = "")
    
    BackendProjectID = frm("BackendProjectID")
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim UseApp: UseApp = frm("UseApp")
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblBackendDatabaseConfigs WHERE BackendProjectID = " & BackendProjectID)
    
    If ExitIfTrue(rs.EOF, "Enter database information..") Then Exit Function
    
    ''Template - ENV File
    Dim TemplateContent: TemplateContent = GetReplacedTemplate(rs, "ENV File")
    CopyToClipboard TemplateContent
    Dim folder: folder = IIf(UseApp, ClientPath, ProjectPath)
    If ExitIfTrue(isFalse(folder), "Folder is empty..") Then Exit Function
    
    Dim filePath: filePath = folder & ".env"
    WriteToFile filePath, TemplateContent
    
End Function

Public Function GetTemplateContent(description)
        
    Dim descriptionTitle: descriptionTitle = "Template - " & description
    Dim sqlStr: sqlStr = "SELECT * FROM qrySnippetCategories WHERE CategoryName = ""Template"" And (SnippetDescription = " & _
        Esc(description) & " OR SnippetDescription = " & Esc(descriptionTitle) & ")"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    ''qrySnippetCategories
    
    If Not rs.EOF Then
        GetTemplateContent = rs.fields("Snippet")
    Else
        MsgBox "Template - " & Esc(description) & " is not existing.", vbCritical + vbOKOnly
        Exit Function
    End If
     
End Function

Public Function IntegrateTypescript(frm)

    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    
    ''Get all the NPM packages -->
    ''TABLE: qryBackendNPMPackages Fields: BackendNPMPackageID|BackendProjectID|NPMPackageID|Timestamp|CreatedBy
    ''RecordImportID|NPMPackage
    Dim NPMPackages As New clsArray: NPMPackages.arr = Elookups("qryBackendNPMPackages", "BackendProjectID = " & BackendProjectID, "NPMPackage", "BackendNPMPackageID")
    
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /k cd /d " & Esc(ProjectPath) ''CD into the VirtualEnvironmentPath
    shellArr.Add "npm install -g typescript"
    shellArr.Add "npm install --save-dev typescript @types/node @types/express @types/sequelize @types/cors"
    shellArr.Add "tsc --init"
    Call Shell(shellArr.JoinArr(" & "))
    
    Dim distPath: distPath = ProjectPath & "dist\"
    CreateFolder distPath
    
    Dim srcPath: srcPath = ProjectPath & "src\"
    CreateFolder srcPath
    
    ''Modify tsconfig.json
    Dim TemplateContent: TemplateContent = GetTemplateContent("tsconfig.json")
    Dim filePath: filePath = ProjectPath & "tsconfig.json"
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateDBConfigFile(frm As Object, Optional BackendProjectID = "")
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim UseApp: UseApp = frm("UseApp")
    
    Dim fileName: fileName = "db.ts"
    Dim templateName: templateName = IIf(UseApp, fileName & " next 13", fileName)
    Dim TemplateContent: TemplateContent = GetTemplateContent(templateName)
    
    Dim folder: folder = IIf(UseApp, ClientPath, ProjectPath)
    If ExitIfTrue(isFalse(folder), "Folder is empty..") Then Exit Function

    Dim filePath: filePath = folder & "src\config\" & fileName
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateGitIgnoreFile(frm As Object, Optional BackendProjectID = "")
    
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty.") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath")
    
    ''Modify db.ts
    Dim TemplateContent: TemplateContent = GetTemplateContent(".gitignore")
    Dim filePath: filePath = ProjectPath & ".gitignore"
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateEsLintFile(frm As Object, Optional BackendProjectID = "")
    
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty.") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath")
    
    ''Modify db.ts
    Dim TemplateContent: TemplateContent = GetTemplateContent(".eslintrc.json")
    Dim filePath: filePath = ProjectPath & ".eslintrc.json"
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateInterfaceFile(frm As Object, Optional BackendProjectID = "")
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim UseApp: UseApp = frm("UseApp")
    
    Dim fileName: fileName = "interface.ts"
    ''Dim templateName: templateName = IIf(UseApp, fileName & " next 13", fileName)
    Dim templateName: templateName = fileName
    Dim TemplateContent: TemplateContent = GetTemplateContent(templateName)
    
    Dim folder: folder = IIf(UseApp, ClientPath, ProjectPath)
    If ExitIfTrue(isFalse(folder), "Folder is empty..") Then Exit Function

    Dim filePath: filePath = folder & "src\interfaces\" & fileName
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateErrorHandlingFile(frm As Object, Optional BackendProjectID = "")
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim UseApp: UseApp = frm("UseApp")
    
    Dim fileName: fileName = "errorHandling.ts"
    Dim templateName: templateName = "errorHandling.ts next 13"
    Dim TemplateContent: TemplateContent = GetTemplateContent(templateName)
    ''Modify interface.ts
    
    ''Dim templateName: templateName = IIf(UseApp, fileName & " next 13", fileName)
    Dim folder: folder = IIf(UseApp, ClientPath, ProjectPath)
    If ExitIfTrue(isFalse(folder), "Folder is empty..") Then Exit Function

    Dim filePath: filePath = folder & "src\utils\" & fileName
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateGenericUtilFile(frm As Object, Optional BackendProjectID = "")
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim UseApp: UseApp = frm("UseApp")
    
    ''Modify interface.ts
    Dim fileName: fileName = "generic.ts"
    Dim templateName: templateName = IIf(UseApp, "generic.ts next 13", "generic.ts")
    Dim TemplateContent: TemplateContent = GetTemplateContent(templateName)
    
    Dim folder: folder = IIf(UseApp, ClientPath, ProjectPath)
    If ExitIfTrue(isFalse(folder), "Folder is empty..") Then Exit Function
    
    Dim filePath: filePath = folder & "src\utils\" & fileName
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateSQLHelperFile(frm As Form)
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim UseApp: UseApp = frm("UseApp")
    
    Dim fileName: fileName = "sqlHelper.ts"
    Dim templateName: templateName = fileName
    Dim TemplateContent: TemplateContent = GetTemplateContent(fileName)
    
    Dim folder: folder = IIf(UseApp, ClientPath, ProjectPath)
    If ExitIfTrue(isFalse(folder), "Folder is empty..") Then Exit Function
    
    Dim filePath: filePath = folder & "src\utils\sqlHelper.ts"
    WriteToFile filePath, TemplateContent
    
End Function

Public Function Create_clsSQLFile(frm As Form)
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    
    ''Modify clsSQL.ts
    Dim TemplateContent: TemplateContent = GetTemplateContent("clsSQL.ts")
    Dim filePath: filePath = ProjectPath & "src\utils\clsSQL.ts"
    WriteToFile filePath, TemplateContent
    
End Function

Public Function Create_clsJoinFile(frm As Form)
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    
    ''Modify clsSQL.ts
    Dim TemplateContent: TemplateContent = GetTemplateContent("clsJoin.ts")
    Dim filePath: filePath = ProjectPath & "src\utils\clsJoin.ts"
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateUtilsFile(frm As Form)
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim UseApp: UseApp = frm("UseApp")
    
    ''Modify interface.ts
    Dim TemplateContent: TemplateContent = GetTemplateContent(IIf(UseApp, "utils.ts next 13", "utils.ts"))
    Dim filePath: filePath = ProjectPath & "src\utils\utils.ts"
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateIndexFile(frm As Form)
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("index.ts")
    Dim filePath: filePath = ProjectPath & "src\index.ts"
    WriteToFile filePath, TemplateContent
    
End Function


Public Function SetAppName(frm As Form)
    
    Dim ProjectPath: ProjectPath = frm("ProjectPath")
    Dim inputString As String
    Dim lastBackslashPos As Integer
    Dim secondLastBackslashPos As Integer
    Dim extractedText As String
    
    ' Input string
    inputString = ProjectPath
    
    ' Find position of last backslash
    lastBackslashPos = InStrRev(inputString, "\")
    
    ' Find position of second to last backslash
    secondLastBackslashPos = InStrRev(inputString, "\", lastBackslashPos - 1)
    
    ' Extract text between the two backslashes
    extractedText = Mid(inputString, secondLastBackslashPos + 1, lastBackslashPos - secondLastBackslashPos - 1)
    
    frm("AppName") = extractedText
    
End Function

Public Function GenerateAllIndexRoute(frm As Form)
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    If isFalse(BackendProjectID) Then Exit Function
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim lines As New clsArray
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GenerateIndexRoute(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    CopyToClipboard lines.JoinArr(vbCrLf)
    
End Function

Public Function GenerateAllModelFiles(frm As Form)
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    If isFalse(BackendProjectID) Then Exit Function
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim lines As New clsArray
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        ImportCompleteModelFile frm, SeqModelID
        ImportCompleteControllerFile frm, SeqModelID
        ImportCompleteRouteFile frm, SeqModelID
        rs.MoveNext
    Loop
    
End Function

Public Function TestAllBackendAPIs(frm As Form)
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    If isFalse(BackendProjectID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendAPIs WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblBackendAPIs Fields: BackendAPIID|BackendProjectID|Endpoint|ReturnString|Timestamp|CreatedBy
    ''RecordImportID|StatusCode|ReturnStringUnformatted
    Do Until rs.EOF
        Dim BackendAPIID: BackendAPIID = rs.fields("BackendAPIID")
        DoCmd.OpenForm "frmBackendAPIs", , , "BackendAPIID = " & BackendAPIID, , acHidden
        TestApi Forms("frmBackendAPIs")
        rs.MoveNext
    Loop
    
    frm("subBackendAPIs").Form.Requery

End Function

Public Function InsertBackendAPIs(frm As Form)
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    If isFalse(BackendProjectID) Then Exit Function
    ''TABLE: tblBackendDatabaseConfigs Fields: BackendProjectID|Port|Host|User|DatabaseName|DatabasePassword
    ''Timestamp|CreatedBy|RecordImportID
    
    Dim Port: Port = ELookup("tblBackendDatabaseConfigs", "BackendProjectID = " & BackendProjectID, "Port")
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        Dim ModelPath: ModelPath = rs.fields("ModelPath")
        ''TABLE: tblBackendAPIs Fields: BackendAPIID|BackendProjectID|Endpoint|ReturnString|Timestamp|CreatedBy
        ''RecordImportID|StatusCode
        
        ''Select All
        Dim EndPoint: EndPoint = "http://localhost:" & Port & "/api/" & ModelPath
        If Not isPresent("tblBackendAPIs", "BackendProjectID = " & BackendProjectID & " AND EndPoint = " & Esc(EndPoint)) Then
            RunSQL "INSERT INTO tblBackendAPIs (BackendProjectID,EndPoint) VALUES " & _
                    "(" & BackendProjectID & "," & Esc(EndPoint) & ")"
        End If
        
        ''Select One
        EndPoint = EndPoint & "/1"
        If Not isPresent("tblBackendAPIs", "BackendProjectID = " & BackendProjectID & " AND EndPoint = " & Esc(EndPoint)) Then
            RunSQL "INSERT INTO tblBackendAPIs (BackendProjectID,EndPoint) VALUES " & _
                    "(" & BackendProjectID & "," & Esc(EndPoint) & ")"
        End If
        
        rs.MoveNext
    Loop
    
    frm("subBackendAPIs").Form.Requery
    
End Function


Public Function RunNpxCreateNextApp(frm As Object, Optional BackendProjectID = "")
    
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty.") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim resp: resp = MsgBox("This will install next.js with typescript at the path: " & Esc(ClientPath) & "." & _
        "Do you want to proceed?", vbYesNo)
    
    If resp = vbNo Then Exit Function
        
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /k cd /d " & Esc(ClientPath) ''CD into the VirtualEnvironmentPath
    shellArr.Add "npx create-next-app@latest ./ --ts --use-npm"
    Call Shell(shellArr.JoinArr(" & "))
    
End Function

Public Function InstallNPMClientPackages(frm As Form)

    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    Dim UseApp: UseApp = frm("UseApp")
    
    ''Get all the NPM packages -->
    Dim NPMPackages As New clsArray
    Dim NPMPackagesDevOnly As New clsArray
    
    If UseApp Then
        NPMPackages.arr = Elookups("qryBackendNPMPackages", "BackendProjectID = " & BackendProjectID & " AND " & _
        "NOT DevOnly", "NPMPackage", "BackendNPMPackageID")
        
        NPMPackagesDevOnly.arr = Elookups("qryBackendNPMPackages", "BackendProjectID = " & BackendProjectID & " AND " & _
        "DevOnly", "NPMPackage", "BackendNPMPackageID")
    Else
        NPMPackages.arr = Elookups("qryBackendNPMPackages", "BackendProjectID = " & BackendProjectID & " AND Location = ""Client"" AND " & _
        "NOT DevOnly", "NPMPackage", "BackendNPMPackageID")
        
        NPMPackagesDevOnly.arr = Elookups("qryBackendNPMPackages", "BackendProjectID = " & BackendProjectID & " AND Location = ""Client"" AND " & _
        "DevOnly", "NPMPackage", "BackendNPMPackageID")
    End If
    
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /k cd /d " & Esc(ClientPath)
    
    Dim resp
    If NPMPackages.count > 0 Then
        resp = MsgBox("This will install the following NPM package: " & NPMPackages.JoinArr(",") & " at the path: " & Esc(ClientPath) & "." & _
        "Do you want to proceed?", vbYesNo)
    
        If resp = vbNo Then Exit Function
        shellArr.Add "npm install " & NPMPackages.JoinArr(" ")
    End If
    
    If NPMPackagesDevOnly.count > 0 Then
        resp = MsgBox("This will install the following NPM package as dev dependency: " & NPMPackagesDevOnly.JoinArr(",") & " at the path: " & Esc(ClientPath) & "." & _
        "Do you want to proceed?", vbYesNo)
        shellArr.Add "npm install --save-dev " & NPMPackagesDevOnly.JoinArr(" ")
    End If
    
    Call Shell(shellArr.JoinArr(" & "))
    
End Function

Public Function CopyClientAPI_ts(frm As Object, Optional BackendProjectID = "")
    
    RunCommandSaveRecord
    
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty.") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    ''TABLE: tblBackendDatabaseConfigs Fields: BackendProjectID|Port|Host|User|DatabaseName|DatabasePassword
    ''Timestamp|CreatedBy|RecordImportID
    sqlStr = "SELECT * FROM tblBackendDatabaseConfigs WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim Port: Port = rs.fields("Port")
    Dim Host: Host = rs.fields("Host")
    If ExitIfTrue(isFalse(Port), "Port is empty..") Then Exit Function
    If ExitIfTrue(isFalse(Host), "Host is empty..") Then Exit Function
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("api.ts")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[Port]", Port)
    replacedContent = replace(replacedContent, "[Host]", Host)
    
    CopyClientAPI_ts = replacedContent
    
    CopyToClipboard CopyClientAPI_ts
    
    Dim filePath: filePath = ClientPath & "src\utils\api.ts"
    WriteToFile filePath, replacedContent
    
End Function

Public Function CopyUtilsFolder(frm As Form)
    
    RunCommandSaveRecord
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim UseApp: UseApp = frm("UseApp")
    
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendDatabaseConfigs WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim Port: Port = rs.fields("Port"): If ExitIfTrue(isFalse(Port), "Port is empty..") Then Exit Function
    If UseApp Then
        Port = Port & "/api"
    End If
    Dim Host: Host = rs.fields("Host"): If ExitIfTrue(isFalse(Host), "Host is empty..") Then Exit Function
    
    Dim folder: folder = "src\utils\"
    Dim files As New clsArray: files.arr = "constants.ts,api.ts,createEmotionCache.ts,formik.tsx,theme.ts,utilities.ts"
    Dim item
    Dim lines As New clsArray
    
    For Each item In files.arr
        Dim fileName: fileName = Trim(item)
        Dim TemplateContent: TemplateContent = GetTemplateContent(fileName)
        
        If item Like "*api.ts*" Then
            TemplateContent = replace(TemplateContent, "[Host]", Host)
            TemplateContent = replace(TemplateContent, "[Port]", Port)
        End If
        
        Set lines = New clsArray
        lines.Add "//Generated by CopyUtilsFolder"
        lines.Add TemplateContent
        TemplateContent = lines.NewLineJoin
    
        Dim filePath: filePath = ClientPath & folder & Trim(item)
        WriteToFile filePath, TemplateContent
    Next item
    
    
    If UseApp Then
    
        TemplateContent = GetTemplateContent("constants.ts next 13")
        filePath = ClientPath & folder & "constants.ts"
        WriteToFile filePath, TemplateContent
        
        Kill ClientPath & folder & "createEmotionCache.ts"
        Kill ClientPath & folder & "formik.tsx"
        Kill ClientPath & folder & "theme.ts"
        
    End If
    
End Function

Public Function CopyHooksFolder(frm As Form)
    
    RunCommandSaveRecord
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    Dim ClientPath: ClientPath = frm("ClientPath")
    
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim files As New clsArray: files.arr = "useLocalStorage, useSnackbar, useToggle, usePromiseAll"
    Dim item
    
    For Each item In files.arr
        Dim trimmedItem: trimmedItem = Trim(item) & ".ts"
        Dim fileName: fileName = trimmedItem
        Dim TemplateContent: TemplateContent = GetTemplateContent(fileName)
        Dim filePath: filePath = ClientPath & "src\hooks\" & trimmedItem
        WriteToFile filePath, TemplateContent
    Next item
    
End Function

Public Function CopyGlobalContextFile(frm As Form)

    RunCommandSaveRecord
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    Dim ClientPath: ClientPath = frm("ClientPath")
    
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Global.tsx")
    
    CopyGlobalContextFile = TemplateContent
    
    CopyToClipboard CopyGlobalContextFile
    
    Dim filePath: filePath = ClientPath & "src\contexts\Global.tsx"
    WriteToFile filePath, CopyGlobalContextFile
    
End Function

Public Function CopyComponentsFolder(frm As Form)
    
    RunCommandSaveRecord
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    Dim ClientPath: ClientPath = frm("ClientPath")
    
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim SourceFolder As String: SourceFolder = "C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Next js templates\components"
    Dim DestinationFolder As String: DestinationFolder = ClientPath & "src\components"
    
    CopyFolderContents SourceFolder, DestinationFolder
    
End Function

Public Function Copy_appFile(frm As Form)

    RunCommandSaveRecord
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim SidebarEnabled: SidebarEnabled = frm("SidebarEnabled")
    
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent(IIf(SidebarEnabled, "_app.tsx Sidebar", "_app.tsx"))
    
    ''_app.tsx Sidebar
    ''_app.tsx
    
    Copy_appFile = TemplateContent
    
    CopyToClipboard Copy_appFile
    
    Dim filePath: filePath = ClientPath & "src\pages\_app.tsx"
    WriteToFile filePath, Copy_appFile
    
End Function

Public Function Copy_documentFile(frm As Form)

    RunCommandSaveRecord
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    Dim ClientPath: ClientPath = frm("ClientPath")
    
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("_document.tsx")
    
    Copy_documentFile = TemplateContent
    
    CopyToClipboard Copy_documentFile
    
    Dim filePath: filePath = ClientPath & "src\pages\_document.tsx"
    WriteToFile filePath, Copy_documentFile
    
End Function

Public Function CopyInterfacesFolder(frm As Form)
    
    RunCommandSaveRecord
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    Dim ClientPath: ClientPath = frm("ClientPath")
    Dim UseApp: UseApp = frm("UseApp")
    
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim SourceFolder As String: SourceFolder = "C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Next js templates\interfaces"
    Dim DestinationFolder As String: DestinationFolder = ClientPath & "src\interfaces"
    
    CopyFolderContents SourceFolder, DestinationFolder
    
    ''Delete the EmotionInterfaces.ts
    If UseApp Then
        Kill DestinationFolder & "\EmotionInterfaces.ts"
    End If
    
    
End Function

Public Function GenerateNavbarWithModels(frm As Object, Optional BackendProjectID = "")
    
    RunCommandSaveRecord
    
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ProjectPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\Navbar.tsx"
    sqlStr = "SELECT * FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND Not NavItemOrder IS NULL ORDER BY NavItemOrder"
        
    Set rs = ReturnRecordset(sqlStr)
    Dim items As New clsArray
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
        items.Add GenerateModelNavbarItem(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    GenerateNavbarWithModels = GetTemplateContent("Navbar with Models")
    GenerateNavbarWithModels = replace(GenerateNavbarWithModels, "[Items]", items.JoinArr(vbNewLine))
    
    CopyToClipboard GenerateNavbarWithModels
    WriteToFile filePath, GenerateNavbarWithModels
    
End Function

Public Function Copytheme_ts(frm As Form)
    
    RunCommandSaveRecord
    
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    Dim ClientPath: ClientPath = frm("ClientPath")
    
    ''TABLE: tblBackendDatabaseConfigs Fields: BackendProjectID|Port|Host|User|DatabaseName|DatabasePassword
    ''Timestamp|CreatedBy|RecordImportID
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendDatabaseConfigs WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("theme.ts")
    
    Copytheme_ts = TemplateContent
    
    CopyToClipboard Copytheme_ts
    
    Dim filePath: filePath = ClientPath & "src\utils\theme.ts"
    WriteToFile filePath, Copytheme_ts
    
End Function

Public Function Copyclienttsconfigjson(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Copyclienttsconfigjson = GetReplacedTemplate(rs, "client tsconfig.json")
    CopyToClipboard Copyclienttsconfigjson
    
    Dim filePath: filePath = ClientPath & "tsconfig.json"
    WriteToFile filePath, Copyclienttsconfigjson
    
End Function

Public Function RunBackendAndFrontEndServer(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim shellArr As New clsArray
    'shellArr.Add "cmd.exe cd " & Esc("C:\Users")
    shellArr.Add "cmd.exe /k cd " & Esc(ProjectPath)
    shellArr.Add "npm run dev"
    
    Shell shellArr.JoinArr(" & "), vbNormalFocus
    
    Set shellArr = New clsArray
    shellArr.Add "cmd.exe /k cd " & Esc(ClientPath)
    shellArr.Add "npm run dev"
    
    Shell shellArr.JoinArr(" & "), vbNormalFocus
    
    
End Function

Public Function GenerateSidebarWithModels(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function

    GenerateSidebarWithModels = GetReplacedTemplate(rs, "Sidebar")
    
    Set rs = ReturnRecordset("SELECT * FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & " AND NOT NavItemOrder IS NULL ORDER BY NavItemOrder")
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetSidebarLink(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    GenerateSidebarWithModels = replace(GenerateSidebarWithModels, "[SidebarLinks]", lines.JoinArr(vbNewLine))
    
    CopyToClipboard GenerateSidebarWithModels
    
    Dim filePath: filePath = ClientPath & "src\components\Sidebar.tsx"
    WriteToFile filePath, GenerateSidebarWithModels
    
End Function

Public Function GeneratePages(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryPages WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Do Until rs.EOF
        Dim PageID: PageID = rs.fields("PageID")
        GeneratePageFile frm, PageID
        rs.MoveNext
    Loop
    
End Function


Public Function Modify_packagejson_port(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim AppName: AppName = rs.fields("AppName"): If ExitIfTrue(isFalse(AppName), "AppName is empty..") Then Exit Function
    
    sqlStr = "SELECT * FROM tblBackendDatabaseConfigs WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim Port: Port = rs.fields("Port"): If ExitIfTrue(isFalse(Port), "Port is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "package.json"
    
    ' Read the JSON file as text
    Open filePath For Input As #1
    Dim jsonText: jsonText = Input$(LOF(1), #1)
    Close #1
    
    Dim jsonObj As Object: Set jsonObj = JsonConverter.ParseJson(jsonText)
    ' Replace the name and type property
    jsonObj("name") = AppName
    ''jsonObj("type") = "commonjs"
    
    Dim overridesObj As Object
    Set overridesObj = CreateObject("Scripting.Dictionary")
    
    Set jsonObj("overrides") = overridesObj
    
    jsonObj("overrides")("@react-pdf/image") = "2.2.3"
    jsonObj("overrides")("@react-pdf/pdfkit") = "3.0.4"
    jsonObj("overrides")("@react-pdf/layout") = "3.5.0"
    
    Dim scriptsObj As Object: Set scriptsObj = jsonObj("scripts")
    scriptsObj("start") = "next start -p " & Port
    scriptsObj("dev") = "next dev -p " & Port
    scriptsObj("test") = "jest --watch"
    
    ' Convert the dictionary object back to JSON text
    jsonText = JsonConverter.ConvertToJson(jsonObj, Whitespace:=2)
    
    ' Write the updated JSON text back to the file
    Open filePath For Output As #1
    Print #1, jsonText
    Close #1
    
    Modify_packagejson_port = jsonText
    CopyToClipboard Modify_packagejson_port
    
End Function

Public Function Generate_middleware_ts(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim IsSupabase: IsSupabase = rs.fields("IsSupabase")
    
    Generate_middleware_ts = GetReplacedTemplate(rs, IIf(IsSupabase, "middleware.ts supabase", "middleware.ts"))
    Generate_middleware_ts = GetGeneratedByFunctionSnippet(Generate_middleware_ts, "Generate_middleware_ts")
    CopyToClipboard Generate_middleware_ts
    
    Dim filePath: filePath = ClientPath & "src\middleware.ts"
    
    ''middleware.ts supabase
    
    WriteToFile filePath, Generate_middleware_ts
    
End Function

Public Function Paste_interfacesFolder(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If
    
    Dim fileName: fileName = "GeneralInterfaces.ts"
    Dim AccessPath As String: AccessPath = CurrentProject.path & "\Files\Next js templates\interfaces\" & fileName

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    CopyFile ClientPath & "src\interfaces" & fileName, AccessPath
    
    
End Function

Public Function Paste_utilsFolder(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If
    
    Dim ClientPath: ClientPath = frm("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath As String: filePath = ClientPath & "src\utils\utilities.ts"
    Dim fileContent: fileContent = ReadTextFile(filePath)
    Dim templateName: templateName = "Template - utilities.ts"
    
    DoCmd.OpenForm "frmSnippets", , , "SnippetDescription = " & Esc(templateName)
    Forms("frmSnippets")("Snippet") = fileContent
    ''Dim sqlStr: sqlStr = "UPDATE tblSnippets SET Snippet = " & Esc(fileContent) & " WHERE SnippetDescription = " & Esc(templateName)
    
    ''RunSQL sqlStr
    
End Function


Public Function Generate_next_config(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Generate_next_config = GetReplacedTemplate(rs, "next.config.js")
    CopyToClipboard Generate_next_config
    
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "next.config.js"
    
    WriteToFile filePath, Generate_next_config
    
End Function

Public Function Create_formik_tsx_file_Next13(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Create_formik_tsx_file_Next13 = GetReplacedTemplate(rs, "formik.tsx")
    Create_formik_tsx_file_Next13 = GetGeneratedByFunctionSnippet(Create_formik_tsx_file_Next13, "Create_formik_tsx_file_Next13")
    CopyToClipboard Create_formik_tsx_file_Next13
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\utils\formik.tsx"
    WriteToFile filePath, Create_formik_tsx_file_Next13

End Function

Public Function Initialize_sequelize_cli(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    
    lines.Add "cmd.exe /k cd /d " & Esc(ProjectPath) & "src\"
    lines.Add "npx sequelize-cli init"
    Call Shell(lines.JoinArr(" & "))
    
    ''Edit the config file
    EditSequelizeCLIConfigFile frm, BackendProjectID
    
End Function

Public Function EditSequelizeCLIConfigFile(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    EditSequelizeCLIConfigFile = GetReplacedTemplate(rs, "sequelize-cli config.js")
    EditSequelizeCLIConfigFile = GetGeneratedByFunctionSnippet(EditSequelizeCLIConfigFile, "EditSequelizeCLIConfigFile")
    CopyToClipboard EditSequelizeCLIConfigFile
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim filePath: filePath = ProjectPath & "src\config\config.js"
    WriteToFile filePath, EditSequelizeCLIConfigFile

End Function

Public Function Generate_sequelizercFile(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID & " AND NOT IsSupabase"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then Exit Function
    Generate_sequelizercFile = GetReplacedTemplate(rs, ".sequelizerc")
    Generate_sequelizercFile = GetGeneratedByFunctionSnippet(Generate_sequelizercFile, "Generate_sequelizercFile")
    CopyToClipboard Generate_sequelizercFile
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim filePath: filePath = ProjectPath & ".sequelizerc"
    WriteToFile filePath, Generate_sequelizercFile
    
End Function

Public Function RunSequelizeDbMigrate(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    
    lines.Add "cmd.exe /k cd /d " & Esc(ProjectPath)
    lines.Add "npx sequelize db:migrate"
    Call Shell(lines.JoinArr(" & "))
    
End Function

Public Function RunSequelizeDbundoMigrate(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim count: count = InputBox("Undo last x migrations..", , 1)
    
    If ExitIfTrue(isFalse(count), "Invalid count value.") Then Exit Function
    
    Dim command: command = "npx sequelize db:migrate:undo"
    If count > 1 Then
        command = command & " --count " & count
    End If
    lines.Add "cmd.exe /k cd /d " & Esc(ProjectPath)
    lines.Add command
    Call Shell(lines.JoinArr(" & "))
    
End Function

Public Function CreateENVFileNext13ForProduction(frm As Object, Optional BackendProjectID = "")

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty.") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath")
    Dim ClientPath: ClientPath = rs.fields("ClientPath")
    Dim UseApp: UseApp = rs.fields("UseApp")
    Dim Port: Port = rs.fields("Port")
    Dim Host: Host = rs.fields("Host"): If ExitIfTrue(isFalse(Host), "Host is empty..") Then Exit Function
    ''Template - ENV File
    Dim TemplateContent: TemplateContent = GetReplacedTemplate(rs, "ENV File Next 13")
    
    Dim SanitizedAppName: SanitizedAppName = rs.fields("SanitizedAppName"): If ExitIfTrue(isFalse(SanitizedAppName), "SanitizedAppName is empty..") Then Exit Function
    SanitizedAppName = replace(SanitizedAppName, "_", "-")
    Dim Domain: Domain = "https://" & SanitizedAppName & ".vercel.app"

    TemplateContent = replace(TemplateContent, "[Domain]", Domain)
    TemplateContent = replace(TemplateContent, "NODE_ENV=development", "NODE_ENV=production")
    
    CopyToClipboard TemplateContent
    MsgBox ".env file copied"
    
End Function

Public Function CreateENVFileNext13(frm As Object, Optional BackendProjectID = "")

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty.") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath")
    Dim ClientPath: ClientPath = rs.fields("ClientPath")
    Dim UseApp: UseApp = rs.fields("UseApp")
    Dim Port: Port = rs.fields("Port")
    Dim Host: Host = rs.fields("Host"): If ExitIfTrue(isFalse(Host), "Host is empty..") Then Exit Function
    ''Template - ENV File
    Dim TemplateContent: TemplateContent = GetReplacedTemplate(rs, "ENV File Next 13")
    
    Dim Domain: Domain = "http://" & Host
    If Not isFalse(Port) Then
        Domain = Domain & ":" & Port
    End If
    
    TemplateContent = replace(TemplateContent, "[Domain]", Domain)
    
    CopyToClipboard TemplateContent
    Dim folder: folder = IIf(UseApp, ClientPath, ProjectPath)
    If ExitIfTrue(isFalse(folder), "Folder is empty..") Then Exit Function
    
    Dim filePath: filePath = folder & ".env"
    WriteToFile filePath, TemplateContent
    
End Function

Public Function CreateNextAuthRouteFile(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    CreateNextAuthRouteFile = GetReplacedTemplate(rs, "NextAuth route file")
    CreateNextAuthRouteFile = GetGeneratedByFunctionSnippet(CreateNextAuthRouteFile, "CreateNextAuthRouteFile")
    CopyToClipboard CreateNextAuthRouteFile
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ProjectPath & "src\app\api\auth\[...nextauth]\route.ts"
    WriteToFile filePath, CreateNextAuthRouteFile
    
End Function

Public Function CreateAuthProviderComponent(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    CreateAuthProviderComponent = GetReplacedTemplate(rs, "AuthProvider component")
    CreateAuthProviderComponent = GetGeneratedByFunctionSnippet(CreateAuthProviderComponent, "CreateAuthProviderComponent")
    CopyToClipboard CreateAuthProviderComponent
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ProjectPath & "src\components\auth-provider\AuthProvider.tsx"
    WriteToFile filePath, CreateAuthProviderComponent
    
End Function

Public Function Create_tailwind_config_js_shadcn_file(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Create_tailwind_config_js_shadcn_file = GetReplacedTemplate(rs, "tailwind.config.js shadcn")
    Create_tailwind_config_js_shadcn_file = GetGeneratedByFunctionSnippet(Create_tailwind_config_js_shadcn_file, "Create_tailwind_config_js_shadcn_file", "tailwind.config.js shadcn")
    CopyToClipboard Create_tailwind_config_js_shadcn_file
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    ''C:\Users\User\Desktop\Web Development\next-13-tutorial\tailwind.config.js
    Dim filePath: filePath = ClientPath & "tailwind.config.ts"
    WriteToFile filePath, Create_tailwind_config_js_shadcn_file
    
End Function

Public Function Create_globals_css_shadcn_file(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Create_globals_css_shadcn_file = GetReplacedTemplate(rs, "globals.css shadcn")
    
    CopyToClipboard Create_globals_css_shadcn_file
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    ''C:\Users\User\Desktop\Web Development\next-13-tutorial\src\app\globals.css
    Dim filePath: filePath = ClientPath & "src\app\globals.css"
    WriteToFile filePath, Create_globals_css_shadcn_file
    
End Function

Public Function Create_lib_utils_ts_file(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Create_lib_utils_ts_file = GetReplacedTemplate(rs, "lib/utils.ts")
    Create_lib_utils_ts_file = GetGeneratedByFunctionSnippet(Create_lib_utils_ts_file, "Create_lib_utils_ts_file")
    CopyToClipboard Create_lib_utils_ts_file
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\lib\utils.ts"
    WriteToFile filePath, Create_lib_utils_ts_file
    
End Function

Public Function Create_next_auth_d_ts(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Create_next_auth_d_ts = GetReplacedTemplate(rs, "next-auth.d.ts")
    Create_next_auth_d_ts = GetGeneratedByFunctionSnippet(Create_next_auth_d_ts, "Create_next_auth_d_ts")
    CopyToClipboard Create_next_auth_d_ts
    
    ''C:\Users\User\Desktop\Web Development\next-13-tutorial\types\next-auth.d.ts
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\types\next-auth.d.ts"
    WriteToFile filePath, Create_next_auth_d_ts
    
End Function

Public Function InstallShadComponents(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblAppShadComponents WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim AppShadComponentID: AppShadComponentID = rs.fields("AppShadComponentID")
        InstallShadComponent frm, AppShadComponentID
        rs.MoveNext
    Loop

End Function

Public Function Create_components_json(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Create_components_json = GetReplacedTemplate(rs, "components.json")
    CopyToClipboard Create_components_json
    
    ''C:\Users\User\Desktop\Web Development\next-13-tutorial\components.json
    Dim ClientPath: ClientPath = rs.fields("ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "components.json"
    WriteToFile filePath, Create_components_json
    
End Function


Public Function CopyAllTemplateFilesNext13(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord
    
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim IsSupabase: IsSupabase = rs.fields("IsSupabase")
    Dim ClientPath: ClientPath = rs.fields("ClientPath")
    
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim FolderName: FolderName = IIf(IsSupabase, "Next 13 templates - supabase", "Next 13 templates")
    
    Dim ProjectPath: ProjectPath = CurrentProject.path
    
    Dim SourceFolder As String: SourceFolder = ProjectPath & "\files\" & FolderName & "\"
    Dim DestinationFolder As String: DestinationFolder = ClientPath
    
    CopyFolderContents SourceFolder, DestinationFolder
    
End Function

Public Function OpenProjectInVscode(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /c cd /d " & Esc(ClientPath)
    shellArr.Add "code ."
    Call Shell(shellArr.JoinArr(" & "))
    
End Function

Public Function InstallNpmPackageForNext13(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    ''Get all the NPM packages -->
    
    Dim NPMPackages As New clsArray: NPMPackages.arr = Elookups("qryBackendNPMPackages", "NOT DevOnly AND NOT Inherent " & _
        "AND BackendProjectID = " & BackendProjectID, "NPMPackage", "NPMPackageID")
    Dim NPMPackagesDevOnly As New clsArray: NPMPackagesDevOnly.arr = Elookups("qryBackendNPMPackages", "DevOnly AND NOT Inherent " & _
        "AND BackendProjectID = " & BackendProjectID, "NPMPackage", "NPMPackageID")
    
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /k cd /d " & Esc(ClientPath) ''CD into the VirtualEnvironmentPath
    
    If NPMPackages.count > 0 Then
        shellArr.Add "npm install " & NPMPackages.JoinArr(" ")
    End If
    
    If NPMPackagesDevOnly.count > 0 Then
        shellArr.Add "npm install --save-dev " & NPMPackagesDevOnly.JoinArr(" ")
    End If
    
    Call Shell(shellArr.JoinArr(" & "))
    
End Function

Public Function GetAllHeaderLinks(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    sqlStr = "SELECT * FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & " AND NOT NavItemOrder IS NULL ORDER BY NavItemOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
        lines.Add GetModelHeaderLinkItem(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    GetAllHeaderLinks = GetTemplateContent("header-links")
    GetAllHeaderLinks = replace(GetAllHeaderLinks, "[HeaderLinks]", lines.JoinArr("," & vbCrLf))
    GetAllHeaderLinks = replace(GetAllHeaderLinks, "[GetAllItemIcons]", GetAllItemIcons(frm, BackendProjectID))
    GetAllHeaderLinks = GetGeneratedByFunctionSnippet(GetAllHeaderLinks, "GetAllHeaderLinks", "header-links")
    CopyToClipboard GetAllHeaderLinks
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\lib\header-links.ts
    Dim filePath: filePath = ClientPath & "src\lib\header-links.ts"
    WriteToFile filePath, GetAllHeaderLinks
    
End Function

Public Function WriteToLayout_tsx(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim IsSupabase: IsSupabase = rs.fields("IsSupabase")
    Dim SidebarEnabled: SidebarEnabled = rs.fields("SidebarEnabled")
    Dim templateName As String: templateName = IIf(SidebarEnabled, "GenerateMainLayoutSidebar", "GenerateMainLayout")

    WriteToLayout_tsx = GetReplacedTemplate(rs, templateName)
    WriteToLayout_tsx = GetGeneratedByFunctionSnippet(WriteToLayout_tsx, "WriteToLayout_tsx", templateName)
    CopyToClipboard WriteToLayout_tsx
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\app\layout.tsx
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\app\layout.tsx"
    WriteToFile filePath, WriteToLayout_tsx
    
End Function

Public Function WriteToFooter_tsx(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SidebarEnabled: SidebarEnabled = rs.fields("SidebarEnabled")
    Dim templateName As String: templateName = IIf(SidebarEnabled, "Footer.tsx sidebar", "Footer.tsx")

    WriteToFooter_tsx = GetReplacedTemplate(rs, templateName)
    WriteToFooter_tsx = GetGeneratedByFunctionSnippet(WriteToFooter_tsx, "WriteToFooter_tsx", templateName)
    CopyToClipboard WriteToFooter_tsx
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\components\Footer.tsx
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\Footer.tsx"
    WriteToFile filePath, WriteToFooter_tsx

End Function

Public Function DoNPMRunStart(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim shellArr As New clsArray
    shellArr.Add "cmd.exe /k cd " & Esc(ClientPath)
    shellArr.Add "npm run start"
    
    Shell shellArr.JoinArr(" & "), vbNormalFocus
    
End Function

Public Function DoNPMRunDev(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim shellArr As New clsArray
    shellArr.Add "cmd.exe /k cd " & Esc(ClientPath)
    shellArr.Add "npm run dev"
    
    Shell shellArr.JoinArr(" & "), vbNormalFocus
    
End Function

Public Function NpmInstallFromPackage_json(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim shellArr As New clsArray
    shellArr.Add "cmd.exe /k cd /d " & Esc(ClientPath) ''CD into the VirtualEnvironmentPath
    shellArr.Add "npm install"
    
    Call Shell(shellArr.JoinArr(" & "))
    
End Function

Public Function GenerateProjectModelRelationships(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray, fieldVals As New clsArray
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE NOT RelatedModelID IS NULL AND BackendProjectID = " & BackendProjectID & " ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
        Dim RelatedModelID: RelatedModelID = rs.fields("RelatedModelID"): If ExitIfTrue(isFalse(RelatedModelID), "RelatedModelID is empty..") Then Exit Function
        Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
        Dim RightForeignKey: RightForeignKey = ELookup("tblSeqModelFields", "SeqModelID = " & RelatedModelID & " AND PrimaryKey", "DatabaseFieldName")
        
        Set fields = New clsArray
        fields.Add "BackendProjectID"
        fields.Add "LeftModelID"
        fields.Add "RightModelID"
        fields.Add "LeftForeignKey"
        fields.Add "RightForeignKey"
        fields.Add "Relationship" ''1:M
        fields.Add "DeclareInModel"
        
        Set fieldVals = New clsArray
        fieldVals.Add BackendProjectID
        fieldVals.Add SeqModelID
        fieldVals.Add RelatedModelID
        fieldVals.Add Esc(DatabaseFieldName)
        fieldVals.Add Esc(RightForeignKey)
        fieldVals.Add Esc("1:M") ''1:M
        fieldVals.Add SeqModelID
        
        RunSQL "INSERT INTO tblSeqModelRelationships (" & fields.JoinArr & ") VALUES (" & fieldVals.JoinArr & ")"
        rs.MoveNext
    Loop
    ''Get all the Models and loop for each
    ''Loop at each Model Field and get all those that has RelatedModelID
    
    ''The related model id will be RightModelID the modelid will be the LeftModelID
    ''The database fieldName will be the LeftModelKey
    ''Get the PrimaryKey of the LeftModelID for the RightModelKey
    
    ''Requery the relationship subform
    frm("subSeqModelRelationships").Form.Requery
    
End Function

Public Function FixDatabaseFromApi(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If
    
    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim Host: Host = rs.fields("Host"): If ExitIfTrue(isFalse(Host), "Host is empty..") Then Exit Function
    Dim Port: Port = rs.fields("Port"): If ExitIfTrue(isFalse(Port), "Port is empty..") Then Exit Function
    
    Dim endpointUrl: endpointUrl = "http://" & Host & ":" & Port & "/api/fix-database" & "?rand=" & Rnd()
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "POST", endpointUrl, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.setRequestHeader "Pragma", "no-cache"
    http.send
    
    Dim StatusCode: StatusCode = http.status
    
    Dim json As Object
    Set json = JsonConverter.ParseJson(http.responseText)
    
    Dim prettyJson As String
    prettyJson = JsonConverter.ConvertToJson(json)
    prettyJson = UnescapeJson(prettyJson)
    
    Dim response: response = PrettifyJson(prettyJson)
    response = Left(response, 500)
    
    MsgBox "Finished fixing database" & vbNewLine & "Status Code: " & StatusCode & vbNewLine & "Response: " & response
    
End Function

Public Function GetAllItemIcons(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND NavItemOrder > 0 ORDER BY NavItemOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetItemIcons(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllItemIcons = lines.JoinArr(",")
        GetAllItemIcons = GetGeneratedByFunctionSnippet(GetAllItemIcons, "GetAllItemIcons")
        CopyToClipboard GetAllItemIcons
    End If

End Function

Public Function WriteToModelconfig_tsInterface(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToModelconfig_tsInterface = GetReplacedTemplate(rs, "ModelConfig.ts interface")
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetModelKeyInterface]", GetKVInterface("qrySeqModels"))
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetModelFieldKeyInterface]", GetKVInterface("qrySeqModelFields"))
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetModelFilterKeyInterface]", GetKVInterface("qrySeqModelFilters"))
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetModelSortKeyInterface]", GetKVInterface("qrySeqModelSorts"))
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetModelHookInterface]", GetKVInterface("qrySeqModelHooks"))
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetModelFieldGroups]", GetKVInterface("qrySeqModelFieldGroups"))
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetModelEmbeddings]", GetKVInterface("qrySeqModelEmbeddings"))
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetModelFilterOptionKeyInterface]", GetKVInterface("qrySeqModelFilterOptions"))
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetAllConfigControlTypes]", GenerateControlTypes)
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetAllConfigDataType]", GenerateDataTypes)
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetAllFilterOperator]", GenerateFilterOperators)
    WriteToModelconfig_tsInterface = replace(WriteToModelconfig_tsInterface, "[GetAllSummarizedBy]", GetAllSummarizedBy)
    WriteToModelconfig_tsInterface = GetGeneratedByFunctionSnippet(WriteToModelconfig_tsInterface, "WriteToModelconfig_tsInterface", "ModelConfig.ts interface")
    CopyToClipboard WriteToModelconfig_tsInterface
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\interfaces\ModelConfig.ts
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\interfaces\ModelConfig.ts"
    WriteToFile filePath, WriteToModelconfig_tsInterface

    
End Function

Public Function GetKVInterface(QueryName)

    Dim fld As field, db As Database
    Set db = CurrentDb
    
    Dim sqlStr: sqlStr = "SELECT * FROM qryConfigEnumeratorFields WHERE QueryName = " & Esc(QueryName) & " AND Not Exclude ORDER BY FieldOrder,ConfigEnumeratorFieldID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim lines As New clsArray
    Do Until rs.EOF
        
        Dim fieldName: fieldName = rs.fields("FieldName")
        Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "Variable Name is empty..") Then Exit Function
        Dim Nullable: Nullable = rs.fields("Nullable")
        
        Set fld = db.QueryDefs(QueryName).fields(fieldName)
        
        Dim FieldType
        
        If fieldName = "filterOperator" Then
            FieldType = "FilterOperator"
        ElseIf fieldName = "controlType" Then
            FieldType = "ControlType"
        ElseIf fieldName = "dataType" Then
            FieldType = "DataType"
        ElseIf fieldName = "summarizedBy" Then
            FieldType = "SummarizedBy"
        Else
            FieldType = GetFieldTypeInterface(fld.Type)
        End If
        If Nullable Then
            FieldType = FieldType & " | null"
        End If
        
        lines.Add VariableName & ": " & FieldType & ";"
        ''FieldName: FieldType;
        rs.MoveNext
    Loop
    
    GetKVInterface = lines.JoinArr(vbNewLine)
    
End Function


Private Function GenerateControlTypes()

    Dim rs As Recordset: Set rs = ReturnRecordset("Select WebControlType FROM tblWebControlTypes ORDER by WebControlType")
    Dim lines As New clsArray
    Do Until rs.EOF
        lines.Add rs.fields("WebControlType")
        rs.MoveNext
    Loop
    
    lines.EscapeItems
    GenerateControlTypes = lines.JoinArr(" | ")
    
End Function

Private Function GenerateDataTypes()

    Dim rs As Recordset: Set rs = ReturnRecordset("Select DataType FROM tblSeqDataTypes ORDER by DataType")
    Dim lines As New clsArray
    Do Until rs.EOF
        lines.Add rs.fields("DataType")
        rs.MoveNext
    Loop
    
    lines.EscapeItems
    GenerateDataTypes = lines.JoinArr(" | ")
    
End Function

Private Function GenerateFilterOperators()

    Dim rs As Recordset: Set rs = ReturnRecordset("Select FilterOperator FROM tblSeqModelFilterOperators ORDER by FilterOperator")
    Dim lines As New clsArray
    Do Until rs.EOF
        lines.Add rs.fields("FilterOperator")
        rs.MoveNext
    Loop
    
    lines.EscapeItems
    GenerateFilterOperators = lines.JoinArr(" | ")
    
End Function

Private Function GetAllSummarizedBy()

    Dim rs As Recordset: Set rs = ReturnRecordset("Select SummarizedBy FROM tblSummarizedBys ORDER by SummarizedBy")
    Dim lines As New clsArray
    Do Until rs.EOF
        lines.Add rs.fields("SummarizedBy")
        rs.MoveNext
    Loop
    
    lines.EscapeItems
    GetAllSummarizedBy = lines.JoinArr(" | ")
    
End Function


Public Function GetFieldTypeInterface(fldType)
        
    If isFieldBoolean(fldType) Then
        GetFieldTypeInterface = "boolean"
    ElseIf isFieldNumeric(fldType) Then
        GetFieldTypeInterface = "number"
    Else
        GetFieldTypeInterface = "string"
    End If
    
End Function

Public Function WriteToAppconfig_ts(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToAppconfig_ts = GetReplacedTemplate(rs, "AppConfig.ts")
    WriteToAppconfig_ts = replace(WriteToAppconfig_ts, "[GetRelationshipInterface]", GetKVInterface("qrySeqModelRelationships"))
    WriteToAppconfig_ts = replace(WriteToAppconfig_ts, "[GetAllSeqModelKeys]", GetAllSeqModelKeys(frm, BackendProjectID))
    WriteToAppconfig_ts = replace(WriteToAppconfig_ts, "[GetAllModelConfigImports]", GetAllModelConfigImports(frm, BackendProjectID))
    WriteToAppconfig_ts = replace(WriteToAppconfig_ts, "[GetAllSeqModelRelationshipKeys]", GetAllSeqModelRelationshipKeys(frm, BackendProjectID))
    WriteToAppconfig_ts = replace(WriteToAppconfig_ts, "[GeAllAppKeys]", GeAllAppKeys(frm, BackendProjectID))
    WriteToAppconfig_ts = replace(WriteToAppconfig_ts, "[GetAllAppKeyInterface]", GetKVInterface("qryBackendProjects"))
    WriteToAppconfig_ts = GetGeneratedByFunctionSnippet(WriteToAppconfig_ts, "WriteToAppconfig_ts", "AppConfig.ts")
    CopyToClipboard WriteToAppconfig_ts
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\lib\app-config.ts
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\app-config.ts"
    WriteToFile filePath, WriteToAppconfig_ts
    
End Function

Public Function GetAllSeqModelKeys(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID, ModelName FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
        lines.Add ModelName & "Config,"
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSeqModelKeys = lines.JoinArr(vbNewLine)
        GetAllSeqModelKeys = GetGeneratedByFunctionSnippet(GetAllSeqModelKeys, "GetAllSeqModelKeys")
        CopyToClipboard GetAllSeqModelKeys
    End If

End Function

Public Function GetAllModelConfigImports(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelConfigImports(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelConfigImports = lines.JoinArr(vbNewLine)
        GetAllModelConfigImports = GetGeneratedByFunctionSnippet(GetAllModelConfigImports, "GetAllModelConfigImports")
        CopyToClipboard GetAllModelConfigImports
    End If

End Function

Public Function GetAllSeqModelRelationshipKeys(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSeqModelRelationshipKeys(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSeqModelRelationshipKeys = lines.JoinArr(vbNewLine)
        GetAllSeqModelRelationshipKeys = GetGeneratedByFunctionSnippet(GetAllSeqModelRelationshipKeys, "GetAllSeqModelRelationshipKeys")
        CopyToClipboard GetAllSeqModelRelationshipKeys
    End If

End Function

Public Function GeAllAppKeys(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GeAllAppKeys = GetKVPairs("qryBackendProjects", rs)
    
End Function

Public Function WriteToBackendmodels_ts(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID & " AND NOT IsSupabase"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then Exit Function

    WriteToBackendmodels_ts = GetReplacedTemplate(rs, "backend-models.ts")
    WriteToBackendmodels_ts = replace(WriteToBackendmodels_ts, "[GetAllBackendModelImports]", GetAllBackendModelImports(frm, BackendProjectID))
    WriteToBackendmodels_ts = replace(WriteToBackendmodels_ts, "[GetAllBackedModelName]", GetAllBackedModelName(frm, BackendProjectID))
    WriteToBackendmodels_ts = GetGeneratedByFunctionSnippet(WriteToBackendmodels_ts, "WriteToBackendmodels_ts", "backend-models.ts")
    CopyToClipboard WriteToBackendmodels_ts
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\lib\backend-models.ts
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "Project Path is empty..") Then Exit Function
    Dim filePath: filePath = ProjectPath & "src\lib\backend-models.ts"
    WriteToFile filePath, WriteToBackendmodels_ts
    
End Function

Public Function GetAllBackendModelImports(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetBackendModelImports(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllBackendModelImports = lines.JoinArr(vbNewLine)
        GetAllBackendModelImports = GetGeneratedByFunctionSnippet(GetAllBackendModelImports, "GetAllBackendModelImports")
        CopyToClipboard GetAllBackendModelImports
    End If

End Function

Public Function GetAllBackedModelName(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetBackedModelName(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllBackedModelName = lines.JoinArr(vbNewLine)
        GetAllBackedModelName = GetGeneratedByFunctionSnippet(GetAllBackedModelName, "GetAllBackedModelName")
        CopyToClipboard GetAllBackedModelName
    End If

End Function

Public Function WriteToLucideicons_tsx(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToLucideicons_tsx = GetReplacedTemplate(rs, "LucideIcons.tsx")
    WriteToLucideicons_tsx = replace(WriteToLucideicons_tsx, "[GetAllLucideIcons]", GetAllLucideIcons(frm, BackendProjectID))
    WriteToLucideicons_tsx = replace(WriteToLucideicons_tsx, "[GetAllLucideIconItems]", GetAllLucideIconItems(frm, BackendProjectID))
    WriteToLucideicons_tsx = GetGeneratedByFunctionSnippet(WriteToLucideicons_tsx, "WriteToLucideicons_tsx", "LucideIcons.tsx")
    CopyToClipboard WriteToLucideicons_tsx
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\components\LucideIcons.tsx
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\LucideIcons.tsx"
    WriteToFile filePath, WriteToLucideicons_tsx
    
End Function

Public Function GetAllLucideIcons(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND Not NavItemIcon IS NULL AND NOT NavItemOrder IS NULL ORDER BY NavItemOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetLucideIcon(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllLucideIcons = lines.JoinArr(vbNewLine)
        GetAllLucideIcons = GetGeneratedByFunctionSnippet(GetAllLucideIcons, "GetAllLucideIcons")
        CopyToClipboard GetAllLucideIcons
    End If

End Function

Public Function GetAllLucideIconItems(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND Not NavItemIcon IS NULL AND NOT NavItemOrder IS NULL ORDER BY NavItemOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetLucideIconItem(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllLucideIconItems = lines.JoinArr(vbNewLine)
        GetAllLucideIconItems = GetGeneratedByFunctionSnippet(GetAllLucideIconItems, "GetAllLucideIconItems")
        CopyToClipboard GetAllLucideIconItems
    End If

End Function

Public Function WriteToGetmodeloptions_tsx(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetmodeloptions_tsx = GetReplacedTemplate(rs, "getModelOptions.tsx")
    WriteToGetmodeloptions_tsx = replace(WriteToGetmodeloptions_tsx, "[GetAllComponentImport]", GetAllComponentImport(frm, BackendProjectID))
    WriteToGetmodeloptions_tsx = replace(WriteToGetmodeloptions_tsx, "[GetAllExportConstModelOptions]", GetAllExportConstModelOptions(frm, BackendProjectID))
    
    WriteToGetmodeloptions_tsx = GetGeneratedByFunctionSnippet(WriteToGetmodeloptions_tsx, "WriteToGetmodeloptions_tsx", "getModelOptions.tsx")
    CopyToClipboard WriteToGetmodeloptions_tsx
    
    ''C:\Users\User\Desktop\Web Development\personal-finance-next-13\src\lib\getModelOptions.tsx
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\getModelOptions.tsx"
    WriteToFile filePath, WriteToGetmodeloptions_tsx
    
End Function

Public Function GetAllComponentImport(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetComponentImport(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllComponentImport = lines.JoinArr(vbNewLine)
        GetAllComponentImport = GetGeneratedByFunctionSnippet(GetAllComponentImport, "GetAllComponentImport")
        CopyToClipboard GetAllComponentImport
    End If

End Function

Public Function GetAllModelSingleColumn(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelSingleColumn(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelSingleColumn = lines.JoinArr(vbNewLine)
        GetAllModelSingleColumn = GetGeneratedByFunctionSnippet(GetAllModelSingleColumn, "GetAllModelSingleColumn")
        CopyToClipboard GetAllModelSingleColumn
    End If

End Function

Public Function GetModelReplacementForModelOption(frm As Object, Optional BackendProjectID = "", Optional templateName = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetBasicSeqModelReplacement(frm, SeqModelID, templateName)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetModelReplacementForModelOption = lines.JoinArr(vbNewLine)
        GetModelReplacementForModelOption = GetGeneratedByFunctionSnippet(GetModelReplacementForModelOption, "GetModelReplacementForModelOption")
        CopyToClipboard GetModelReplacementForModelOption
    End If

End Function

Public Function GetAllModelForm(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelForm(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelForm = lines.JoinArr(vbNewLine)
        GetAllModelForm = GetGeneratedByFunctionSnippet(GetAllModelForm, "GetAllModelForm")
        CopyToClipboard GetAllModelForm
    End If

End Function

Public Function GetGrantPrivelegeToSchema(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetGrantPrivelegeToSchema = GetReplacedTemplate(rs, "Grant Privelege to schema")
    'GetGrantPrivelegeToSchema = GetGeneratedByFunctionSnippet(GetGrantPrivelegeToSchema, "GetGrantPrivelegeToSchema", "Grant Privelege to schema")
    CopyToClipboard GetGrantPrivelegeToSchema
    
End Function

Public Function GetAllPostgresqlCreateTableStatement(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND ViewName IS NULL ORDER BY CreationOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetPostgresqlCreateTableStatement(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllPostgresqlCreateTableStatement = lines.JoinArr(vbNewLine)
        ''GetAllPostgresqlCreateTableStatement = GetGeneratedByFunctionSnippet(GetAllPostgresqlCreateTableStatement, "GetAllPostgresqlCreateTableStatement")
        ''CopyToClipboard GetAllPostgresqlCreateTableStatement
        
        DoCmd.OpenForm "frmClipboardForms"
        Forms("frmClipboardForms")("Snippet") = GetAllPostgresqlCreateTableStatement
        CopyFieldContent Forms("frmClipboardForms"), "Snippet"
        
    End If

End Function

Public Function GetDROPAllStatementForPostgreSQL(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY NavItemOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetDROPStatementForPostgreSQL(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetDROPAllStatementForPostgreSQL = lines.JoinArr(vbNewLine)
        ''GetDROPAllStatementForPostgreSQL = GetGeneratedByFunctionSnippet(GetDROPAllStatementForPostgreSQL, "GetDROPAllStatementForPostgreSQL")
        CopyToClipboard GetDROPAllStatementForPostgreSQL
    End If

End Function

Public Function GetAllPostgrePlainCreateStatement(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY NavItemOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetPostgrePlainCreateStatement(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllPostgrePlainCreateStatement = lines.JoinArr(vbNewLine)
        DoCmd.OpenForm "frmChatGPTPrompts", , , "Title = " & Esc("postgreSQL test data")
        Forms("frmChatGPTPrompts")("PrePrompt") = GetAllPostgrePlainCreateStatement
        CopyCombinedPrompt Forms("frmChatGPTPrompts")
        ''GetAllPostgrePlainCreateStatement = GetGeneratedByFunctionSnippet(GetAllPostgrePlainCreateStatement, "GetAllPostgrePlainCreateStatement")
        CopyToClipboard GetAllPostgrePlainCreateStatement
    End If

End Function

Public Function GetAllDeletePostgresql(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND ViewName IS NULL ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetDeletePostgresql(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllDeletePostgresql = lines.JoinArr(vbNewLine)
        ''GetAllDeletePostgresql = GetGeneratedByFunctionSnippet(GetAllDeletePostgresql, "GetAllDeletePostgresql")
        CopyToClipboard GetAllDeletePostgresql
    End If

End Function


Public Function GetAllAlterPolicyStatements(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND ViewName IS NULL ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetAlterPolicyStatements(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllAlterPolicyStatements = lines.JoinArr(vbNewLine)
        ''GetAllAlterPolicyStatements = GetGeneratedByFunctionSnippet(GetAllAlterPolicyStatements, "GetAllAlterPolicyStatements")
        CopyToClipboard GetAllAlterPolicyStatements
    End If

End Function

Public Function ReinitializeSchema(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    ReinitializeSchema = GetReplacedTemplate(rs, "Reinitialize schema")
    ReinitializeSchema = replace(ReinitializeSchema, "[GetGrantPrivelegeToSchema]", GetGrantPrivelegeToSchema(frm, BackendProjectID))
    ''ReinitializeSchema = GetGeneratedByFunctionSnippet(ReinitializeSchema, "ReinitializeSchema", "Reinitialize schema")
    CopyToClipboard ReinitializeSchema
    
End Function

Public Function GetAllCreateIndexFromRelationship(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID, LeftModelID, RightModelID FROM tblSeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        Dim RightModelID: RightModelID = rs.fields("RightModelID"): If ExitIfTrue(isFalse(RightModelID), "RightModelID is empty..") Then Exit Function
        
        If isPresent("qrySeqModels", "SeqModelID = " & LeftModelID & " AND (NoTable OR IsPureView)") Then
            GoTo NextRecord:
        End If
        
        If isPresent("qrySeqModels", "SeqModelID = " & RightModelID & " AND (NoTable OR IsPureView)") Then
            GoTo NextRecord:
        End If
    
        lines.Add GetCreateIndexFromRelationship(frm, SeqModelRelationshipID)
NextRecord:
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllCreateIndexFromRelationship = lines.JoinArr(vbNewLine)
        ''GetAllCreateIndexFromRelationship = GetGeneratedByFunctionSnippet(GetAllCreateIndexFromRelationship, "GetAllCreateIndexFromRelationship")
        CopyToClipboard GetAllCreateIndexFromRelationship
    End If

End Function

Public Function GetAllResetSerialAutonumber(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND ViewName IS NULL ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        If isPresent("tblSeqModelFields", "Autoincrement And PrimaryKey AND SeqModelID = " & SeqModelID) Then
            lines.Add GetResetSerialAutonumber(frm, SeqModelID)
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllResetSerialAutonumber = lines.JoinArr(vbNewLine)
        ''GetAllResetSerialAutonumber = GetGeneratedByFunctionSnippet(GetAllResetSerialAutonumber, "GetAllResetSerialAutonumber")
        CopyToClipboard GetAllResetSerialAutonumber
    End If

End Function

Public Function GetAllProjectIndexByFilter(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND ViewName IS NULL ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetAllIndextStatementByFilter(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllProjectIndexByFilter = lines.JoinArr(vbNewLine)
        ''GetAllProjectIndexByFilter = GetGeneratedByFunctionSnippet(GetAllProjectIndexByFilter, "GetAllProjectIndexByFilter")
        CopyToClipboard GetAllProjectIndexByFilter
    End If

End Function

Public Function WriteAllToHookPostRoute(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add WriteToHookPostRoutePerModel(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        WriteAllToHookPostRoute = lines.JoinArr(vbNewLine)
        WriteAllToHookPostRoute = GetGeneratedByFunctionSnippet(WriteAllToHookPostRoute, "WriteAllToHookPostRoute")
        CopyToClipboard WriteAllToHookPostRoute
    End If

End Function

Public Function CopySupabaseCustomFunctions(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    lines.Add GetReplacedTemplate(rs, "slugify")
    lines.Add GetReplacedTemplate(rs, "insert_with_children")
    lines.Add GetReplacedTemplate(rs, "upsert_with_children_text")
    lines.Add GetReplacedTemplate(rs, "upsert_with_children_date")
    lines.Add GetReplacedTemplate(rs, "multi_upsert")
    lines.Add GetReplacedTemplate(rs, "upsert_with_children")
    lines.Add GetReplacedTemplate(rs, "public.does_record_exists")
    lines.Add GetReplacedTemplate(rs, "public.does_record_exists_date")
    lines.Add GetReplacedTemplate(rs, "Create user_sessions table")
    lines.Add GetReplacedTemplate(rs, "public.dynamic_select")
    
    ''CopySupabaseCustomFunctions = GetGeneratedByFunctionSnippet(CopySupabaseCustomFunctions, "CopySupabaseCustomFunctions", Null)
    CopyToClipboard lines.JoinArr(vbNewLine)
    
End Function

Public Function AddAllSlugColumnSql(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND NOT SlugField IS NULL ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add AddSlugColumnSql(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        AddAllSlugColumnSql = lines.JoinArr(vbNewLine)
        ''AddAllSlugColumnSql = GetGeneratedByFunctionSnippet(AddAllSlugColumnSql, "AddAllSlugColumnSql")
        CopyToClipboard AddAllSlugColumnSql
    End If

End Function

Public Function WriteToProtectedLayout_tsx(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToProtectedLayout_tsx = GetReplacedTemplate(rs, "Protected Layout.tsx")
    WriteToProtectedLayout_tsx = GetGeneratedByFunctionSnippet(WriteToProtectedLayout_tsx, "WriteToProtectedLayout_tsx", "Protected Layout.tsx")
    CopyToClipboard WriteToProtectedLayout_tsx
    
    ''C:\Users\User\Desktop\Web Development\panda-realty\src\app\(protected)\layout.tsx
    ''Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\app\(protected)\layout.tsx"
    WriteToFile filePath, WriteToProtectedLayout_tsx
    
End Function

Public Function WriteToAuthLayout_tsx(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToAuthLayout_tsx = GetReplacedTemplate(rs, "Auth Layout.tsx")
    WriteToAuthLayout_tsx = GetGeneratedByFunctionSnippet(WriteToAuthLayout_tsx, "WriteToAuthLayout_tsx", "Auth Layout.tsx")
    CopyToClipboard WriteToAuthLayout_tsx
    
    ''C:\Users\User\Desktop\Web Development\panda-realty\src\app\(auth)\layout.tsx
    ''Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\app\(auth)\layout.tsx"
    WriteToFile filePath, WriteToAuthLayout_tsx
    
End Function

Public Function AddAllForeignKeyConstraintPostgres(frm As Object, Optional BackendProjectID = "") As String

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND VIewName IS NULL ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add AddForeignKeyConstraintPostgres(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        AddAllForeignKeyConstraintPostgres = lines.JoinArr(vbNewLine)
        ''AddAllForeignKeyConstraintPostgres = GetGeneratedByFunctionSnippet(AddAllForeignKeyConstraintPostgres, "AddAllForeignKeyConstraintPostgres")
        CopyToClipboard AddAllForeignKeyConstraintPostgres
    End If

End Function

Public Function WriteToJest_config_mjs(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToJest_config_mjs = GetReplacedTemplate(rs, "jest.config.mjs")
    WriteToJest_config_mjs = GetGeneratedByFunctionSnippet(WriteToJest_config_mjs, "WriteToJest_config_mjs", "jest.config.mjs")
    CopyToClipboard WriteToJest_config_mjs
    
    ''C:\Users\User\Desktop\Web Development\panda-realty\jest.config.mjs
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "jest.config.mjs"
    WriteToFile filePath, WriteToJest_config_mjs
    
End Function

Public Function PromptCreatePostgresqlView(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY NavItemOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetPostgrePlainCreateStatement(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        PromptCreatePostgresqlView = lines.JoinArr(vbNewLine)
        DoCmd.OpenForm "frmChatGPTPrompts", , , "Title = " & Esc("PostgreSQL view creator")
        Forms("frmChatGPTPrompts")("PrePrompt") = PromptCreatePostgresqlView
        CopyCombinedPrompt Forms("frmChatGPTPrompts")
        ''GetAllPostgrePlainCreateStatement = GetGeneratedByFunctionSnippet(GetAllPostgrePlainCreateStatement, "GetAllPostgrePlainCreateStatement")
        CopyToClipboard PromptCreatePostgresqlView
    End If
    
End Function

Public Function GoToExposeSchemaLink(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SupabaseURL: SupabaseURL = rs.fields("SupabaseURL"): If ExitIfTrue(isFalse(SupabaseURL), """SupabaseURL"" is empty..") Then Exit Function
    Dim SanitizedAppName: SanitizedAppName = rs.fields("SanitizedAppName"): If ExitIfTrue(isFalse(SanitizedAppName), "SanitizedAppName is empty..") Then Exit Function
    
    MsgBox "Edit the ""API settings"" > ""Exposed schemas"" and add the " & Esc(SanitizedAppName), vbOKOnly
    
    SupabaseURL = extractSupabaseURL(SupabaseURL)
    
    Dim url As String: url = "https://supabase.com/dashboard/project/" & SupabaseURL & "/settings/api" ' Replace this with the URL you want to open
    ' Use Shell function to open the default web browser with the specified URL
    Shell "cmd /c start " & url, vbNormalFocus
    
End Function

Public Function GetAllModelFormikControlsOnChanges(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelFormikControlsOnChange(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFormikControlsOnChanges = lines.JoinArr(vbNewLine)
        GetAllModelFormikControlsOnChanges = GetGeneratedByFunctionSnippet(GetAllModelFormikControlsOnChanges, "GetAllModelFormikControlsOnChanges")
        CopyToClipboard GetAllModelFormikControlsOnChanges
    End If

End Function

Public Function GetAllModelColumnsToBeOverriden(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelColumnsToBeOverriden(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelColumnsToBeOverriden = lines.JoinArr(vbNewLine)
        GetAllModelColumnsToBeOverriden = GetGeneratedByFunctionSnippet(GetAllModelColumnsToBeOverriden, "GetAllModelColumnsToBeOverriden")
        CopyToClipboard GetAllModelColumnsToBeOverriden
    End If

End Function

Public Function GetAllModelFormikControlsOnBlurs(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelFormikControlsOnBlurs(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFormikControlsOnBlurs = lines.JoinArr(vbNewLine)
        GetAllModelFormikControlsOnBlurs = GetGeneratedByFunctionSnippet(GetAllModelFormikControlsOnBlurs, "GetAllModelFormikControlsOnBlurs")
        CopyToClipboard GetAllModelFormikControlsOnBlurs
    End If

End Function

Public Function GetAllModelActions(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelAction(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelActions = lines.JoinArr(vbNewLine)
        GetAllModelActions = GetGeneratedByFunctionSnippet(GetAllModelActions, "GetAllModelActions")
        CopyToClipboard GetAllModelActions
    End If

End Function

Public Function GetAllUseModifiedRequiredLists(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetUseModifiedRequiredList(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllUseModifiedRequiredLists = lines.JoinArr(vbNewLine)
        GetAllUseModifiedRequiredLists = GetGeneratedByFunctionSnippet(GetAllUseModifiedRequiredLists, "GetAllUseModifiedRequiredLists")
        CopyToClipboard GetAllUseModifiedRequiredLists
    End If

End Function

Public Function GetAllGetRequiredListRowFilter(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetGetRequiredListRowFilter(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllGetRequiredListRowFilter = lines.JoinArr(vbNewLine)
        GetAllGetRequiredListRowFilter = GetGeneratedByFunctionSnippet(GetAllGetRequiredListRowFilter, "GetAllGetRequiredListRowFilter")
        CopyToClipboard GetAllGetRequiredListRowFilter
    End If

End Function

Public Function GetAllCustomizedListElements(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetCustomizedListElementsOption(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllCustomizedListElements = lines.JoinArr(vbNewLine)
        GetAllCustomizedListElements = GetGeneratedByFunctionSnippet(GetAllCustomizedListElements, "GetAllCustomizedListElements")
        CopyToClipboard GetAllCustomizedListElements
    End If

End Function

Public Function GetAllModifiedInitialValues(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModifiedInitialValuesLine(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModifiedInitialValues = lines.JoinArr(vbNewLine)
        GetAllModifiedInitialValues = GetGeneratedByFunctionSnippet(GetAllModifiedInitialValues, "GetAllModifiedInitialValues")
        CopyToClipboard GetAllModifiedInitialValues
    End If

End Function

Private Function extractSupabaseURL(url) As String
    Dim startIdx As Integer
    Dim endIdx As Integer
    Dim subStr As String
    
    ' Find the starting index of the substring
    startIdx = InStr(url, "https://") + Len("https://")
    
    ' Find the ending index of the substring
    endIdx = InStr(startIdx, url, ".supabase.co")
    
    ' Extract the substring
    subStr = Mid(url, startIdx, endIdx - startIdx)
    
    ' Return the extracted substring
    extractSupabaseURL = subStr
End Function

''Command Name: Open Supabase SQL Editor
Public Function OpenSupabaseSqlEditor(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SupabaseURL: SupabaseURL = rs.fields("SupabaseURL"): If ExitIfTrue(isFalse(SupabaseURL), "SupabaseURL is empty..") Then Exit Function
    SupabaseURL = extractSupabaseURL(SupabaseURL)
    
    CreateObject("Shell.Application").Open "https://supabase.com/dashboard/project/" & SupabaseURL & "/sql/new"
    
End Function

Public Function OpenSupabaseSqlEditorThenClose_frmClipboardForms(frm As Form, SeqModelID)
    
    Dim BackendProjectID: BackendProjectID = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "BackendProjectID")
    DoCmd.Close acForm, "frmClipboardForms", acSaveNo
    OpenSupabaseSqlEditor frm, BackendProjectID
    
End Function


''Command Name: Open Supabase Table Editor
Public Function OpenSupabaseTableEditor(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim SupabaseURL: SupabaseURL = rs.fields("SupabaseURL"): If ExitIfTrue(isFalse(SupabaseURL), "SupabaseURL is empty..") Then Exit Function
    SupabaseURL = extractSupabaseURL(SupabaseURL)
    
    CreateObject("Shell.Application").Open "https://supabase.com/dashboard/project/" & SupabaseURL & "/editor"
    
End Function

Public Function GetAllColumnWidthToOverride(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetColumnWidthToOverrideOption(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllColumnWidthToOverride = lines.JoinArr(vbNewLine)
        GetAllColumnWidthToOverride = GetGeneratedByFunctionSnippet(GetAllColumnWidthToOverride, "GetAllColumnWidthToOverride")
        CopyToClipboard GetAllColumnWidthToOverride
    End If

End Function

Public Function GetAllColumnVisibilityToOverride(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetColumnVisibilityToOverrideOption(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllColumnVisibilityToOverride = lines.JoinArr(vbNewLine)
        GetAllColumnVisibilityToOverride = GetGeneratedByFunctionSnippet(GetAllColumnVisibilityToOverride, "GetAllColumnVisibilityToOverride")
        CopyToClipboard GetAllColumnVisibilityToOverride
    End If

End Function

Public Function GetAllColumnOrderToOverride(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetColumnOrderToOverride(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllColumnOrderToOverride = lines.JoinArr(vbNewLine)
        GetAllColumnOrderToOverride = GetGeneratedByFunctionSnippet(GetAllColumnOrderToOverride, "GetAllColumnOrderToOverride")
        CopyToClipboard GetAllColumnOrderToOverride
    End If

End Function

''Command Name: Copy Dashboard Files
Public Function CopyDashboardFiles(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    CopyFolderContents CurrentProject.path & "\Files\Dashboard templates\", ClientPath & "src\app\(protected)\"
    
End Function

Public Function CopyAdminFiles(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    CopyFolderContents CurrentProject.path & "\Files\Admin templates\", ClientPath & "src\app\(protected)\"
    
End Function

Public Function GetAllModelPermissions(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelPermission(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelPermissions = lines.JoinArr(vbNewLine)
        GetAllModelPermissions = GetGeneratedByFunctionSnippet(GetAllModelPermissions, "GetAllModelPermissions")
        CopyToClipboard GetAllModelPermissions
    End If

End Function

''Command Name: Write to LoginForm.tsx
Public Function WriteToLoginform_tsx(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim AllowRegistration: AllowRegistration = rs.fields("AllowRegistration")
    
    Dim templateName: templateName = IIf(AllowRegistration, "LoginForm", "LoginForm registration not allowed")
    
    WriteToLoginform_tsx = GetReplacedTemplate(rs, templateName)
    WriteToLoginform_tsx = GetGeneratedByFunctionSnippet(WriteToLoginform_tsx, "WriteToLoginform_tsx", templateName)
    CopyToClipboard WriteToLoginform_tsx
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\components\login\LoginForm.tsx
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\login\LoginForm.tsx"
    WriteToFile filePath, WriteToLoginform_tsx
    
End Function

Public Function GetAllAdd_updated_at_Column(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID,TableName FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND Not NoTable AND Not IsPureView ORDER BY SeqModelID"
    Dim uqTables As New clsArray
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        Dim TableName: TableName = rs.fields("TableName")
        If Not uqTables.InArray(TableName) Then
            lines.Add GetAdd_updated_at_Column(frm, SeqModelID)
            uqTables.Add TableName
        End If
        
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllAdd_updated_at_Column = lines.JoinArr(vbNewLine)
        'GetAllAdd_updated_at_Column = GetGeneratedByFunctionSnippet(GetAllAdd_updated_at_Column, "GetAllAdd_updated_at_Column")
        CopyToClipboard GetAllAdd_updated_at_Column
    End If

End Function

Public Function GetAllAdd_created_at_Column(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID,TableName FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND Not NoTable AND Not IsPureView ORDER BY SeqModelID"
    Dim uqTables As New clsArray
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        Dim TableName: TableName = rs.fields("TableName")
        If Not uqTables.InArray(TableName) Then
            lines.Add GetAdd_created_at_Column(frm, SeqModelID)
            uqTables.Add TableName
        End If
        
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllAdd_created_at_Column = lines.JoinArr(vbNewLine)
        'GetAllAdd_created_at_Column = GetGeneratedByFunctionSnippet(GetAllAdd_created_at_Column, "GetAllAdd_created_at_Column")
        CopyToClipboard GetAllAdd_created_at_Column
    End If

End Function

Public Function GetAllSet_updated_at_to_NULL(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID,TableName FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND Not NoTable AND Not IsPureView ORDER BY SeqModelID"
    Dim uqTables As New clsArray
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        Dim TableName: TableName = rs.fields("TableName")
        If Not uqTables.InArray(TableName) Then
            lines.Add GetSet_updated_at_to_NULL(frm, SeqModelID)
            uqTables.Add TableName
        End If
        
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSet_updated_at_to_NULL = lines.JoinArr(vbNewLine)
        ''GetAllSet_updated_at_to_NULL = GetGeneratedByFunctionSnippet(GetAllSet_updated_at_to_NULL, "GetAllSet_updated_at_to_NULL")
        CopyToClipboard GetAllSet_updated_at_to_NULL
    End If

End Function

''Command Name: Copy public folder
Public Function CopyPublicFolder(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\public assets
    CopyFolderContents CurrentProject.path & "\Files\public assets\", ProjectPath & "public\"
    
End Function

Public Function GetAllCreatedAndUpdatedIndex(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID,TableName FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND Not NoTable AND Not IsPureView ORDER BY SeqModelID"
    Dim uqTables As New clsArray
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        Dim TableName: TableName = rs.fields("TableName")
        If Not uqTables.InArray(TableName) Then
            lines.Add GetCreatedAndUpdatedIndex(frm, SeqModelID)
            uqTables.Add TableName
        End If
        
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllCreatedAndUpdatedIndex = lines.JoinArr(vbNewLine)
        ''GetAllCreatedAndUpdatedIndex = GetGeneratedByFunctionSnippet(GetAllCreatedAndUpdatedIndex, "GetAllCreatedAndUpdatedIndex")
        CopyToClipboard GetAllCreatedAndUpdatedIndex
    End If

End Function

''Command Name: Sync Dashboard File To Template
Public Function SyncDashboardFileToTemplate(frm As Object, Optional BackendProjectID = "")
    
    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim SourceFolders As New clsArray: SourceFolders.arr = "dashboard,dashboard-postdated,dashboard-uncollected"
    Dim SourceFolder, DestinationFolder
    For Each SourceFolder In SourceFolders.arr
        Dim sourcePath As String: sourcePath = ClientPath & "src\app\(protected)\" & SourceFolder & "\"
        If DirectoryExists(sourcePath) Then
            CopyFolderContents sourcePath, CurrentProject.path & "\Files\Dashboard templates\", True
        End If
    Next SourceFolder
    
End Function

''Command Name: Sync Admin Files to Template
Public Function SyncAdminFilesToTemplate(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim SourceFolders As New clsArray: SourceFolders.arr = "admin"
    Dim SourceFolder, DestinationFolder
    For Each SourceFolder In SourceFolders.arr
        CopyFolderContents ClientPath & "src\app\(protected)\" & SourceFolder & "\", CurrentProject.path & "\Files\Admin templates\", True
    Next SourceFolder
     
End Function

''Command Name: Get Create deleted_records table
Public Function GetCreateDeleted_recordsTable(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetCreateDeleted_recordsTable = GetReplacedTemplate(rs, "create deleted_records table")
    CopyToClipboard GetCreateDeleted_recordsTable
    
End Function

Public Function GetAllDeleteTrigger(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND Not NoTable AND Not IsPureView ORDER BY SeqModelID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetDeleteTrigger(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllDeleteTrigger = lines.JoinArr(vbNewLine)
        ''GetAllDeleteTrigger = GetGeneratedByFunctionSnippet(GetAllDeleteTrigger, "GetAllDeleteTrigger")
        CopyToClipboard GetAllDeleteTrigger
    End If

End Function

Public Function GetAllModelFormGridTemplateAreas(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelFormGridTemplateAreasOption(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFormGridTemplateAreas = lines.JoinArr(vbNewLine)
        GetAllModelFormGridTemplateAreas = GetGeneratedByFunctionSnippet(GetAllModelFormGridTemplateAreas, "GetAllModelFormGridTemplateAreas")
        CopyToClipboard GetAllModelFormGridTemplateAreas
    End If

End Function

Public Function GetAllCustomModelFormElements(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetCustomModelFormElements(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllCustomModelFormElements = lines.JoinArr(vbNewLine)
        GetAllCustomModelFormElements = GetGeneratedByFunctionSnippet(GetAllCustomModelFormElements, "GetAllCustomModelFormElements")
        CopyToClipboard GetAllCustomModelFormElements
    End If

End Function

Public Function GetAllFormikControlsOnTabKeyDowns(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetFormikControlsOnTabKeyDowns(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllFormikControlsOnTabKeyDowns = lines.JoinArr(vbNewLine)
        GetAllFormikControlsOnTabKeyDowns = GetGeneratedByFunctionSnippet(GetAllFormikControlsOnTabKeyDowns, "GetAllFormikControlsOnTabKeyDowns")
        CopyToClipboard GetAllFormikControlsOnTabKeyDowns
    End If

End Function

Public Function GetAllModelFormContainerStyles(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " AND (NOT APIOnly OR APIOnly IS NULL OR ForceModelOption) ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add GetModelFormContainerStylesOption(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFormContainerStyles = lines.JoinArr(vbNewLine)
        GetAllModelFormContainerStyles = GetGeneratedByFunctionSnippet(GetAllModelFormContainerStyles, "GetAllModelFormContainerStyles")
        CopyToClipboard GetAllModelFormContainerStyles
    End If

End Function

Private Function ExtractPathFromSrc(filePath) As String
    ' Find the position of "src\" in the file path
    Dim srcPosition As Integer
    srcPosition = InStr(1, filePath, "src\", vbTextCompare)
    
    ' Check if "src\" was found
    If srcPosition > 0 Then
        ' Extract the substring from "src\" to the end of the string
        ExtractPathFromSrc = "src\" & Mid(filePath, srcPosition + 4) ' +4 to account for the length of "src\"
    Else
        ' If "src\" is not found, return an empty string or handle as needed
        ExtractPathFromSrc = ""
    End If
End Function

''Command Name: Delete Unnecessary files
Public Function DeleteUnnecessaryFiles(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Set rs = ReturnRecordset("SELECT * FROM tblFilesToBeDeleted ORDER BY FilesToBeDeletedID")
    
    ''BaseFilePath == C:\Users\User\Desktop\Web Development\backpack-battles\
    ''Example filePath == C:\Users\User\Desktop\Web Development\backpack-battles\src\components\FormikFormControlGenerator.tsx
    Do Until rs.EOF
        Dim filePath As String: filePath = rs.fields("filePath"): If ExitIfTrue(isFalse(filePath), "filePath is empty..") Then Exit Function
        filePath = ClientPath & ExtractPathFromSrc(filePath)
        DeleteFileIfExists filePath
        rs.MoveNext
    Loop
    
    ''Delete the tailwind.config.js file
    filePath = replace("[ClientPath]tailwind.config.js", "[ClientPath]", ClientPath)
    DeleteFileIfExists filePath
    
    Dim urlTemplate, urlTemplates As New clsArray
    urlTemplates.Add "[ClientPath]src\components\[ModelPath]\[ModelName]Table.tsx"
    urlTemplates.Add "[ClientPath]src\components\[ModelPath]\[ModelName]Form.tsx"
    urlTemplates.Add "[ClientPath]src\components\[ModelPath]\[ModelName]FilterForm.tsx"
    urlTemplates.Add "[ClientPath]src\components\[ModelPath]\[ModelName]StateHolder.tsx"
    
    Dim frmName: frmName = "frmDeleteSeqModelFiles"
    DoCmd.OpenForm frmName
    Set frm = Forms(frmName)
    
    frm("BackendProjectID") = BackendProjectID
    
    For Each urlTemplate In urlTemplates.arr
        frm("FileNamePattern") = urlTemplate
        DeleteSeqModelFiles frm, True
    Next urlTemplate
    
    DoCmd.Close acForm, frmName, acSaveNo
    
    DeleteAllUnnecessaryModelFile frm, BackendProjectID
    
    MsgBox "Files sucessfully deleted.."
    
End Function

''Command Name: Get Create user_sessions table
Public Function GetCreateUser_sessionsTable(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetCreateUser_sessionsTable = GetReplacedTemplate(rs, "Create user_sessions table")
    CopyToClipboard GetCreateUser_sessionsTable
    
End Function

''Command Name: Copy Git Diff Detailed
Public Function CopyGitDiffDetailed(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    CopyGitDiffToClipboard rs, "git diff detailed"
    
    GoToLink "https://www.notion.so/ff9245675adb48328f3c12f885b8f8b7?v=d798464ec8164a74906ff671fefd9e4f"
    GoToLink "https://chatgpt.com/"
    
End Function

Private Function CopyGitDiffToClipboard(rs As Recordset, templateName As String)

    Dim wsh As Object
    Dim waitOnReturn As Boolean
    Dim windowStyle As Integer

    ' Set waitOnReturn to True if you want to wait for the command to complete
    waitOnReturn = True
    ' Set windowStyle to 1 to show the command window, or 0 to hide it
    windowStyle = 1
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    ProjectPath = Left(ProjectPath, Len(ProjectPath) - 1)
    Dim diffFile: diffFile = "my.diff"
    Dim shellArr As New clsArray
    
    Dim MiddlePath: MiddlePath = IIf(Environ("ComputerName") = "DESKTOP-UOL72RE", "Clarisse", "Jet\OneDrive")
    Dim DiffPyPath: DiffPyPath = "C:\Users\" & MiddlePath & "\Desktop\Programming Tool"
    
    shellArr.Add "cmd.exe /c cd /d " & Esc(DiffPyPath)
    shellArr.Add "py ""git diff.py"" " & Esc(ProjectPath)

    Set wsh = CreateObject("WScript.Shell")
    
    wsh.Run shellArr.JoinArr(" & "), windowStyle, waitOnReturn

    ''Read the content of the text file
    Dim filePath As String: filePath = ProjectPath & "\" & diffFile
    Dim fileContent: fileContent = ReadTextFile(filePath)
    
    If isFalse(fileContent) Then
        MsgBox "File content is empty.", vbCritical + vbOKOnly
        Exit Function
    End If

    CopyGitDiffToClipboard = GetReplacedTemplate(rs, templateName)
    CopyGitDiffToClipboard = replace(CopyGitDiffToClipboard, "[gitdiff]", fileContent)
    CopyToClipboard CopyGitDiffToClipboard
    
    MsgBox "Prompt copied to clipboard.."

End Function

''Command Name: Copy Git Diff Brief
Public Function CopyGitDiffBrief(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    CopyGitDiffToClipboard rs, "git diff brief"
    
End Function

Public Function ValidateAllSeqModel(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        lines.Add ValidateSeqmodel(frm, SeqModelID)
        rs.MoveNext
    Loop
    
    ''Validate if there's an OPEN API Key
    If isPresent("tblBackendDatabaseConfigs", "OpenAPIKey IS NULL AND BackendProjectID = " & BackendProjectID) Then
        RunSQL "UPDATE tblBackendDatabaseConfigs SET OpenAPIKey=" & Esc("INSERT YOUR API KEY HERE") & " WHERE BackendProjectID = " & BackendProjectID
    End If
    
End Function

''Command Name: Sync Chatbot Files to Template
Public Function SyncChatbotFilesToTemplate(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim SourceFolders As New clsArray: SourceFolders.arr = "chat"
    Dim SourceFolder, DestinationFolder
    For Each SourceFolder In SourceFolders.arr
        CopyFolderContents ClientPath & "src\app\(protected)\" & SourceFolder & "\", CurrentProject.path & "\Files\Chatbot Files\", True
    Next SourceFolder
    
End Function

''Command Name: Copy Chatbot Files
Public Function CopyChatbotFiles(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    CopyFolderContents CurrentProject.path & "\Files\Chatbot Files\", ClientPath & "src\app\(protected)\"
    
End Function

''Command Name: Sync Prompt Maker Files to Template
Public Function SyncPromptMakerFilesToTemplate(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim SourceFolders As New clsArray: SourceFolders.arr = "prompt"
    Dim SourceFolder, DestinationFolder
    For Each SourceFolder In SourceFolders.arr
        CopyFolderContents ClientPath & "src\app\(protected)\" & SourceFolder & "\", CurrentProject.path & "\Files\Prompt Files\", True
    Next SourceFolder
    
End Function

Public Function SyncPDFFilesToTemplate(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function

    CopyFolderContents ClientPath & "src\components\pdf\", CurrentProject.path & "\Files\PDF Files\", True
       
End Function

Public Function CopyPDFFiles(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    CopyFolderContents CurrentProject.path & "\Files\PDF Files\pdf\", ClientPath & "src\components\"
    
End Function

''Command Name: Copy Prompt Files
Public Function CopyPromptFiles(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''C:\Users\User\Desktop\Web Development\backpack-battles\src\app\(protected)\dashboard\page.tsx
    ''C:\Users\User\Desktop\Programming Tools\Programming Guides\Files\Dashboard templates
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    CopyFolderContents CurrentProject.path & "\Files\Prompt Files\", ClientPath & "src\app\(protected)\"
    
End Function

''Command Name: Create updated_embeddings Table
Public Function CreateUpdated_embeddingsTable(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    CreateUpdated_embeddingsTable = GetReplacedTemplate(rs, "create updated_embeddings table")
    CopyToClipboard CreateUpdated_embeddingsTable
    
End Function

Public Function GetAllFunctionsAndTriggersForEmbedding(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        If isPresent("qrySeqModelEmbeddings", "SeqModelID = " & SeqModelID) Then
            lines.Add GetFunctionsAndTriggersForEmbedding(frm, SeqModelID)
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllFunctionsAndTriggersForEmbedding = lines.JoinArr(vbNewLine)
        CopyToClipboard GetAllFunctionsAndTriggersForEmbedding
    End If

End Function

Public Function DeleteAllUnnecessaryModelFile(frm As Object, Optional BackendProjectID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelID FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & _
        " ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
        DeleteUnnecessaryModelFile frm, SeqModelID
        rs.MoveNext
    Loop
    
End Function

''Command Name: Get All Export Const Model Options
Public Function GetAllExportConstModelOptions(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected, rs2 As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblFunctionChainItems"
          .AddFilter "tblFunctionChainItems.Note Like " & Esc("*Special File*")
          .AddFilter "Not FilePathTemplate IS NULL"
          .AddFilter "Not ModelOptionImportAs IS NULL"
          .AddFilter "Not ExportConst IS NULL"
          .AddFilter "Not TypescriptInference IS NULL"
          .fields = "ModelOptionImportAs,ExportConst,TypescriptInference,FilePathTemplate"
          .joins.Add GenerateJoinObj("tblModelButtons", "ModelButtonID")
          .OrderBy = "FunctionOrder,FunctionChainItemID"
          Set rs2 = .Recordset
    End With
    
    Dim ExportFunctions As New clsArray
    Do Until rs2.EOF
        Dim ModelOptionImportAs: ModelOptionImportAs = rs2.fields("ModelOptionImportAs")
        Dim ExportConst: ExportConst = rs2.fields("ExportConst")
        Dim TypescriptInference: TypescriptInference = rs2.fields("TypescriptInference")
        Dim FilePathTemplate: FilePathTemplate = rs2.fields("FilePathTemplate")
        
        Dim exportConsts As New clsArray, exportConstStr As String
        Set exportConsts = New clsArray
        exportConstStr = ""
        Set rs = ReturnRecordset("SELECT * FROM qrySeqModels WHERE BackendProjectID = " & BackendProjectID & " ORDER BY SeqModelID")
        Do Until rs.EOF
            Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
            Dim filePath: filePath = GetReplacedTemplate(rs, "", , FilePathTemplate)
            Dim ShouldImport: ShouldImport = isPresent("tblSeqModelFiles", "filePath = " & Esc(filePath) & " AND SeqModelID = " & _
                SeqModelID & " AND IsProtected")
            If ShouldImport Then
                exportConsts.Add GetReplacedTemplate(rs, "", , vbTab & vbTab & "[ModelName]: " & ExportConst & ",")
            End If
            rs.MoveNext
            
        Loop
        
        If exportConsts.count > 0 Then
            exportConstStr = exportConsts.NewLineJoin
        End If
        
        ExportFunctions.Add "export const " & ModelOptionImportAs & " = () => {" & vbNewLine & _
                        vbTab & "return {" & vbNewLine & _
                        exportConstStr & vbNewLine & _
                        vbTab & "} as " & TypescriptInference & vbNewLine & _
                        "};" & vbNewLine
        rs2.MoveNext
    Loop
    
    If ExportFunctions.count > 0 Then
        GetAllExportConstModelOptions = ExportFunctions.NewLineJoin
    End If
    
    GetAllExportConstModelOptions = GetGeneratedByFunctionSnippet(GetAllExportConstModelOptions, "GetAllExportConstModelOptions")
    CopyToClipboard GetAllExportConstModelOptions
    
End Function
''Command Name: Run git clone
Public Function RunGitClone(frm As Object, Optional BackendProjectID = "")

    RunCommandSaveRecord

    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackendProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim GithubLink: GithubLink = rs.fields("GithubLink"): If ExitIfTrue(isFalse(GithubLink), """GithubLink"" is empty..") Then Exit Function
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath")
    ProjectPath = RemoveMatchedPattern(ProjectPath, "([a-zA-Z0-9\-]+\\)$")
    
    Dim shellArr As New clsArray
    
    shellArr.Add "cmd.exe /k cd /d " & Esc(ProjectPath)
    
    Dim GitStatement: GitStatement = GetReplacedTemplate(rs, "", , "git clone [GithubLink]")
    shellArr.Add GitStatement
    Call Shell(shellArr.JoinArr(" & "))
    

End Function
