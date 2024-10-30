Attribute VB_Name = "SeqModel Mod"
Option Compare Binary
Option Explicit
Public NoHasWriteToFilePrompt As Boolean

Public Function SeqModelCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4, 5: ''Data Entry Form
            ''CreateViewFunctionButton frm
            If FormTypeID = 4 Then
                frm("pgSeqModelFields").Caption = "Fields"
                frm.PopUp = 0
                frm.OnCurrent = "=SeqModelDEFormOnCurrent([Form])"
                
                frm("listSeqModelActions").Height = GetBottom(frm("tabCtl")) - frm("listSeqModelActions").Top
                frm("listSeqModelActions").VerticalAnchor = acVerticalAnchorBoth
            End If
            
            frm("ModelName").AfterUpdate = "=SeqModelModelNameAfterUpdate([Form])"
            
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Private Sub CreateViewFunctionButton(frm As Form)

    Dim ctl As control
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 0, 0, 0, 0)
    CopyProperties frm, ctl.Name, "ButtonControl"
    ctl.Name = "cmdViewFunctionFromSeqModel"
    ctl.Caption = "View Function"
    ctl.VerticalAnchor = acVerticalAnchorBottom
    ctl.HorizontalAnchor = acHorizontalAnchorRight
    ctl.OnClick = "=ViewFunctionFromSeqModel([Form])"
    
    frm("listSeqModelActions").VerticalAnchor = acVerticalAnchorBoth
    
    ''Also add a view files to this code
    ''button name -> cmdAddSeqModelFiles
    ''subform name -> subSeqModelFiles
    ''control name -> filePath
    
    ''OpenFolderLocation -> The function
    Set ctl = frm("cmdAddSeqModelFiles")
    ctl.Caption = "Open File"
    ctl.OnClick = "=OpenFolderLocation([Form]![subSeqModelFiles].[Form]![filePath])"
    
End Sub

Public Function ViewFunctionFromSeqModel(frm As Form)
    
    Dim ModelButtonID: ModelButtonID = frm("listSeqModelActions"): If ExitIfTrue(isFalse(ModelButtonID), "ModelButtonID is empty..") Then Exit Function
    Dim FunctionName: FunctionName = ELookup("tblModelButtons", "ModelButtonID = " & ModelButtonID, "FunctionName")
    
    DoCmd.OpenModule , FunctionName
    
End Function

Public Function SeqModelModelNameAfterUpdate(frm As Form)

    Dim ModelName: ModelName = frm("ModelName")
    Dim AppName: AppName = frm("BackendProjectID").Column(1)
    
    frm("ExportAs") = ModelName
    frm("CapitalizedName") = UCase(ModelName)
    frm("ModelFileName") = ModelName & "Model.ts"
    frm("ControllerFileName") = ModelName & "Controller.ts"
    frm("InterfaceFileName") = ModelName & "Interfaces.ts"
    frm("RouteFileName") = ModelName & "Route.ts"
    frm("PluralizedModelName") = PluralizeWord(ModelName)
    If isFalse(frm("TableName")) Then frm("TableName") = GetSeqModelTableName(AppName, ModelName)
    frm("ModelPath") = ConvertToModelPath(ModelName)
    frm("VariableName") = FirstCharLowercase(ModelName)
    frm("VariablePluralName") = PluralizeWord(frm("VariableName"))
    frm("VerboseModelName") = ConvertToVerboseCaption(ModelName)
    frm("PluralizedVerboseModelName") = PluralizeWord(ConvertToVerboseCaption(ModelName))
    
End Function

Private Function GetSeqModelTableName(AppName, ModelName) As String
    
    Dim words As New clsArray: words.arr = SeparateWords(ModelName)
    GetSeqModelTableName = LCase(words.JoinArr("_")) & "s"
    
End Function



Public Function CopyModelOptionDict(frm As Form)

    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    
    If isFalse(SeqModelID) Then Exit Function
    
    Dim ModelOptionDict As clsDictionary: Set ModelOptionDict = GetModelOptionDict(SeqModelID)
    CopyToClipboard ModelOptionDict.ToFormatString
    
End Function

Public Function GetModelOptionDict(SeqModelID) As clsDictionary
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Function
    Dim Timestamps:  Timestamps = rs.fields("Timestamps")
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    Dim dict As New clsDictionary
    dict.Add "name", "{singular: " & Esc(ModelName) & ",plural:" & Esc(PluralizedModelName) & "}"
    If Not isFalse(TableName) Then dict.Add "tableName", Esc(TableName)
    If Not Timestamps Then dict.Add "timestamps", "false"
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Not UniqueWith IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    
    If Not rs.EOF Then
    
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim UniqueWith: UniqueWith = rs.fields("UniqueWith"): If ExitIfTrue(isFalse(UniqueWith), "UniqueWith is empty..") Then Exit Function
        
        Dim indexes: indexes = "[{unique: true,fields: [" & Esc(fieldName) & ", " & Esc(UniqueWith) & "],},]"
        dict.Add "indexes", indexes
        
    End If
    
    Set GetModelOptionDict = dict
    
End Function


Public Function GenerateIncludeOption(frm As Object, Optional SeqModelID)
    
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID")
    Dim IsTable: IsTable = rs.fields("IsTable")
    
    Dim filter: filter = "BackendProjectID = " & BackendProjectID & " AND (RightModelID = " & SeqModelID & _
        " OR LeftModelID = " & SeqModelID & ") AND Relationship <> ""M:M"""
    
    If IsTable Then
        filter = filter & " AND NOT ExcludeInTable"
    End If
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE " & filter & " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim includeArray As New clsArray
    
    Do Until rs.EOF
        Dim includeItemDict As New clsDictionary
        If rs.fields("RightModelID") = SeqModelID Then
            includeItemDict.Add "model", rs.fields("LeftModelName")
            includeItemDict.Add "attributes", replace(GenerateAttributesOption(frm, rs.fields("LeftModelID"), True), "attributes:", "")
        Else
            includeItemDict.Add "model", rs.fields("RightModelName")
            includeItemDict.Add "attributes", replace(GenerateAttributesOption(frm, rs.fields("RightModelID"), True), "attributes:", "")
        End If
        
        includeArray.Add includeItemDict.ToFormatString
        rs.MoveNext
    Loop
    
    Dim IncludeDict As New clsDictionary
    IncludeDict.Add "include", "[" & includeArray.JoinArr("," & vbNewLine) & "]"
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateIncludeOption"
    lines.Add IncludeDict.ToFormatString(True)
    GenerateIncludeOption = lines.JoinArr(vbNewLine)
    
    CopyToClipboard GenerateIncludeOption
    
End Function

Public Function CopyModelHooks(frm As Form)

    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    
    If isFalse(SeqModelID) Then Exit Function
    
    Dim ModelHooks: ModelHooks = GetModelHooks(frm, SeqModelID)
    CopyToClipboard ModelHooks
    
End Function


Public Function GetModelHooks(frm As Object, SeqModelID)
    
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    Dim fieldsAsJS: fieldsAsJS = GenerateSQLFieldList(frm, SeqModelID)
    
    Dim ModelName: ModelName = rs.fields("ModelName")
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim VariableName: VariableName = rs.fields("VariableName")

    Dim TemplateContent: TemplateContent = GetTemplateContent("Model Hooks")
    
    Dim replacedContent, replacedContent1, replacedContent2
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[VariableName]", ModelName)
    replacedContent = replace(replacedContent, "[SlugField]", SlugField)
    replacedContent1 = replace(replacedContent, "[CreateOrUpdate]", "Create")
    replacedContent2 = replace(replacedContent, "[CreateOrUpdate]", "Update")
    
    Dim hookArr As New clsArray
    hookArr.Add replacedContent1
    hookArr.Add replacedContent2
    
    GetModelHooks = hookArr.JoinArr(vbNewLine & vbNewLine)
    
    Dim lines As New clsArray
    lines.Add "//Generated by GetModelHooks"
    lines.Add GetModelHooks
    GetModelHooks = lines.NewLineJoin
    
    CopyToClipboard GetModelHooks
    
End Function


Public Function CopyModelFieldsDictionary(frm As Form) As String

    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    
    If isFalse(SeqModelID) Then Exit Function
    
    CopyToClipboard GetModelFieldsDictionary(SeqModelID)
    
End Function

Public Function GetModelFieldsDictionary(SeqModelID, Optional migrateMode As Boolean = False, Optional withBraces As Boolean = True) As String
    
    If isFalse(SeqModelID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Not isExpression ORDER BY FieldOrder"
    ''Loop into each tblSeqModelFields and get their string
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim Timestamps: Timestamps = rs.fields("Timestamps")
    
    Dim fieldDicts As New clsArray
    Do Until rs.EOF
        ''TABLE: tblSeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
        ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        Dim dict As New clsDictionary: Set dict = GetModelFieldDict(SeqModelFieldID, migrateMode)
        
        ''Handle the addition of relationship here as key to the dictionary
        ''Get the fieldName then look up at the relationship where it matches the rightModel and the rightForeignKey
        
        Dim ModelFieldDict: ModelFieldDict = dict.ToFormatString(True)
        
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
        
        Dim SeqModelRelationshipID: SeqModelRelationshipID = ELookup("qrySeqModelRelationships", "LeftModelID = " & SeqModelID & " AND LeftForeignKey = " & Esc(DatabaseFieldName), "SeqModelRelationshipID")
        If Not isFalse(SeqModelRelationshipID) Then
        
            ModelFieldDict = RemoveLastBracket(ModelFieldDict)
            
            Dim refDict As New clsDictionary, frm As Form
            Set frm = Forms("frmSeqModels")
            refDict.Add "references", GetReferencesKeyForModelCreationMigration(frm, SeqModelRelationshipID)
            refDict.Add "onUpdate", """CASCADE"""
            
            Dim IsNullAllowed: IsNullAllowed = isPresent("tblSeqModelFields", "SeqModelID = " & SeqModelID & _
                " AND DatabaseFieldName = " & Esc(DatabaseFieldName) & " AND AllowNull")
            ''This shouldn't always be CASCADE if the related model from
            refDict.Add "onDelete", Esc(IIf(IsNullAllowed, "SET NULL", "CASCADE"))
            
            ModelFieldDict = ModelFieldDict & "," & refDict.ToFormatString(True) & "}"
        End If
        
        fieldDicts.Add ModelFieldDict
        rs.MoveNext
    Loop
    
    ''Dictionary for the slug
    Dim dict3 As New clsDictionary, dict2 As New clsDictionary
    Dim SlugField: SlugField = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "SlugField")
    If Not isFalse(SlugField) Then
        dict3.Add "type", "DataTypes.STRING"
        dict3.Add "unique", "true"
        dict2.Add "slug", dict3 ''produces the slug option
        fieldDicts.Add dict2.ToFormatString(True)
    End If
    
    If Timestamps Then
        Set dict2 = New clsDictionary
        dict2.Add "type", "Sequelize.DATE"
        dict2.Add "defaultValue", "Sequelize.literal(""CURRENT_TIMESTAMP"")"
        dict2.Add "field", """createdAt"""

        fieldDicts.Add "createdAt: " & dict2.ToFormatString

        Set dict2 = New clsDictionary
        dict2.Add "type", "Sequelize.DATE"
        dict2.Add "defaultValue", "Sequelize.literal(""CURRENT_TIMESTAMP"")"
        dict2.Add "field", """updatedAt"""

        fieldDicts.Add "updatedAt: " & dict2.ToFormatString

    End If
    
    GetModelFieldsDictionary = fieldDicts.JoinArr("," & vbCrLf)
    If withBraces Then
        GetModelFieldsDictionary = "{" & GetModelFieldsDictionary & "}"
    End If
    
    GetModelFieldsDictionary = SanitizeJSONString(GetModelFieldsDictionary)
    GetModelFieldsDictionary = GetGeneratedByFunctionSnippet(GetModelFieldsDictionary, "GetModelFieldsDictionary")
    
End Function

Public Function CopyModelDefinition(frm As Form)
    
    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    
    If isFalse(SeqModelID) Then Exit Function

    CopyToClipboard GetModelDefinition(SeqModelID)
    
End Function

Public Function GetModelDefinition(SeqModelID) As String
    
    If isFalse(SeqModelID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName, ExportAs, TableName, Timestamps, BackendProjectID
    ModelName = rs.fields("ModelName")
    ExportAs = rs.fields("ExportAs")
    TableName = rs.fields("TableName")
    Timestamps = rs.fields("Timestamps")
    BackendProjectID = rs.fields("BackendProjectID")
    
    Dim lines As New clsArray
    lines.Add "//Generated by GetModelDefinition"
    ''export const Deck = sequelize.define(
    lines.Add "export const " & ExportAs & " = sequelize.define<" & ModelName & ">("
    ''"Deck",
    lines.Add Esc(ModelName) & ","
    lines.Add GetModelFieldsDictionary(SeqModelID) & ","
    
    Dim options As New clsDictionary
    Set options = GetModelOptionDict(SeqModelID)
    lines.Add "//Generated By GetModelOptionDict"
    lines.Add options.ToFormatString
    
    '');
    lines.Add ");"
    
    GetModelDefinition = lines.JoinArr(vbCrLf)
    GetModelDefinition = replace(GetModelDefinition, "Sequelize.DATE", "DataTypes.DATE")
    
End Function

Public Function CopyModelInterface(frm As Form)
    
    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    
    If isFalse(SeqModelID) Then Exit Function

    CopyToClipboard GetModelInterface(SeqModelID)
    
End Function

Public Function GetModelInterface(SeqModelID) As String
    
'    export default interface Deck
'      extends Model<InferAttributes<Deck>, InferCreationAttributes<Deck>> {
'      id: CreationOptional<number>;
'      name: string;
'    }

    If isFalse(SeqModelID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT isExpression ORDER BY SeqModelFieldID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: qrySeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName|DataType|DataTypeInterface
    
    ''Construct the Field Interface dictionary
    Dim fieldInterface As New clsArray
    Do Until rs.EOF
    
        Dim SeqModelFieldID, Unique, fieldName, SeqDataTypeID, Autoincrement, PrimaryKey, AllowNull, DataTypeOption, DataType, DataTypeInterface
        SeqModelFieldID = rs.fields("SeqModelFieldID")
        Unique = rs.fields("Unique")
        fieldName = rs.fields("FieldName")
        SeqDataTypeID = rs.fields("SeqDataTypeID")
        Autoincrement = rs.fields("Autoincrement")
        PrimaryKey = rs.fields("PrimaryKey")
        AllowNull = rs.fields("AllowNull")
        DataTypeOption = rs.fields("DataTypeOption")
        DataType = rs.fields("DataType")
        DataTypeInterface = rs.fields("DataTypeInterface")
        
        Dim value: value = DataTypeInterface
        
        ''If Enum, use the choices provided
        If DataType = "ENUM" Then
            value = ConvertEnumToArray(DataTypeOption).JoinArr(" | ")
        End If
        
        If PrimaryKey And Autoincrement Then
            ''CreationOptional<number>;
            value = "CreationOptional<number>"
        End If
        
        If AllowNull Then
            value = value & " | null"
        End If
        
        
        fieldInterface.Add fieldName & IIf(AllowNull Or DataType = "BOOLEAN", "?", "") & ": " & value

        rs.MoveNext
    Loop

    Dim SlugField: SlugField = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "SlugField")
    If Not isFalse(SlugField) Then
        fieldInterface.Add "slug : CreationOptional<string>"
    End If
    
    ''Construct the Interface
    sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelID"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName, ExportAs, TableName, Timestamps, BackendProjectID
    ModelName = rs.fields("ModelName")
    ExportAs = rs.fields("ExportAs")
    TableName = rs.fields("TableName")
    Timestamps = rs.fields("Timestamps")
    BackendProjectID = rs.fields("BackendProjectID")
    
    
    If Timestamps Then
        fieldInterface.Add "createdAt: CreationOptional<Date>"
        fieldInterface.Add "updatedAt: CreationOptional<Date>"
    End If
    
    Dim lines As New clsArray
    ''export default interface Deck extends Model<InferAttributes<Deck>, InferCreationAttributes<Deck>>
    lines.Add "//Generated by GetModelInterface"
    lines.Add "export default interface " & ModelName & " extends Model<InferAttributes<" & ModelName & ">, InferCreationAttributes<" & ModelName & ">>"
    lines.Add "{" & fieldInterface.JoinArr(";" & vbCrLf) & "}"
    
    GetModelInterface = lines.JoinArr(vbCrLf)
    
End Function

Public Function CopyModelImports(frm As Form)
    
    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    
    If isFalse(SeqModelID) Then Exit Function

    CopyToClipboard GetModelImports(SeqModelID)
    
End Function

Public Function GetModelImports(SeqModelID) As String
    
    If isFalse(SeqModelID) Then Exit Function
    
    Dim lines As New clsArray
    
    Dim fromSequelize As New clsArray: fromSequelize.arr = "CreationOptional,DataTypes,InferAttributes,InferCreationAttributes,Model,Sequelize"
    Dim importFromSequelize: importFromSequelize = "import {" & fromSequelize.JoinArr & "} from ""sequelize"";"
    lines.Add "//Generated by GetModelImports"
    lines.Add importFromSequelize
    lines.Add "import sequelize from ""../config/db"";"
    
    ''Declare slugify import here (if being slugified)
    Dim SlugField: SlugField = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "SlugField")
    If Not isFalse(SlugField) Then
        lines.Add "import slugify from ""slugify"";"
    End If
    
    ''Declare any relationship imports here.
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE DeclareInModel = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim ModelToUse: ModelToUse = rs.fields("LeftModelID")
        If SeqModelID = ModelToUse Then
            ModelToUse = rs.fields("RightModelID")
        End If
        ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
        ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath
        Dim ModelName: ModelName = ELookup("tblSeqModels", "SeqModelID = " & ModelToUse, "ModelName")
        ''import { Hero } from "./HeroModel";
        lines.Add "import { " & ModelName & " } from ""./" & ModelName & "Model"";"
        
'        Dim ThroughModelName: ThroughModelName = rs.fields("ThroughModelName")
'        If Not isFalse(ThroughModelName) Then
'            lines.Add "import { " & ThroughModelName & " } from ""./" & ThroughModelName & "Model"";"
'        End If
        rs.MoveNext
    Loop
    
    GetModelImports = lines.JoinArr(vbCrLf)
    
End Function

Public Function CopyCompleteModelFile(frm As Form)
    
    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    
    If isFalse(SeqModelID) Then Exit Function

    CopyToClipboard GetCompleteModelFile(SeqModelID)
    
End Function

Public Function GetCompleteModelFile(SeqModelID) As String
    
    If isFalse(SeqModelID) Then Exit Function
    
    Dim lines As New clsArray
    
    lines.Add "//Generated by GetCompleteModelFile"
    ''Get the imports
    lines.Add GetModelImports(SeqModelID)
    
    ''Get the interfaces
    lines.Add GetModelInterface(SeqModelID)
    
    ''Get the fields
    lines.Add GetModelDefinition(SeqModelID)
    
    Dim SlugField: SlugField = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "SlugField")
    
    Dim frm As Form: Set frm = Screen.ActiveForm
    
    If Not isFalse(SlugField) Then
        lines.Add GetModelHooks(frm, SeqModelID)
    End If
    
    lines.Add GenerateSyncModel(frm, SeqModelID)
    ''Add the relationships to be declared in this model
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE DeclareInModel = " & SeqModelID & " ORDER BY SeqModelRelationshipID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GenerateModelRelationship(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    GetCompleteModelFile = lines.JoinArr(vbCrLf & vbCrLf)
    
End Function

Public Function ImportCompleteModelFile(frm As Object, Optional SeqModelID = Null)
    
    RunCommandSaveRecord
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''Get the path and file name of the model
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''Skip if IsSupabase
    Dim IsSupabase: IsSupabase = rs.fields("IsSupabase")
    If IsSupabase Then Exit Function
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName
    Dim ModelFileName: ModelFileName = rs.fields("ModelFileName")
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID")
    Dim NextAuthModel: NextAuthModel = rs.fields("NextAuthModel")
    
    If NextAuthModel Then
        GetCompleteNextAuthModelFile frm, SeqModelID
        Exit Function
    End If
    If ExitIfTrue(isFalse(ModelFileName), "Please provide the Model File Name..") Then Exit Function
    
    Dim ProjectPath: ProjectPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ProjectPath")
    Dim filePath: filePath = ProjectPath & "src\models\" & ModelFileName
    Dim ModelFileContent: ModelFileContent = GetCompleteModelFile(SeqModelID)
    
    ImportCompleteModelFile = ModelFileContent
    
    Dim lines As New clsArray
    lines.Add "//Generated by ImportCompleteModelFile"
    lines.Add ImportCompleteModelFile
    ImportCompleteModelFile = lines.NewLineJoin
    
    CopyToClipboard ImportCompleteModelFile
    WriteToFile filePath, ImportCompleteModelFile, SeqModelID
    
End Function

Public Function CreateModelInterfaceFile(frm As Object, Optional SeqModelID = Null)
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    ''Get the path and file name of the model
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName
    Dim InterfaceFileName: InterfaceFileName = rs.fields("InterfaceFileName")
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID")
    
    If ExitIfTrue(isFalse(InterfaceFileName), "Please provide the Model Interface File Name..") Then Exit Function
    
    Dim ProjectPath: ProjectPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ProjectPath")
    Dim filePath: filePath = ProjectPath & "src\interfaces\" & InterfaceFileName
    CreateModelInterfaceFile = GetCompleteModelInterface(frm, SeqModelID)
    CreateModelInterfaceFile = GetGeneratedByFunctionSnippet(CreateModelInterfaceFile, "CreateModelInterfaceFile")
    
    WriteToFile filePath, CreateModelInterfaceFile, SeqModelID
    
End Function

Public Function DeclareModelBody(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelFieldsArr As New clsArray
    Do Until rs.EOF
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        ModelFieldsArr.Add fieldName & ": string"
        rs.MoveNext
    Loop
    
    ModelFieldsArr.Add "checked : boolean"
    ModelFieldsArr.Add "touched : boolean"
    
    DeclareModelBody = "{" & ModelFieldsArr.JoinArr(";" & vbNewLine) & "}"
    
    Dim lines As New clsArray
    lines.Add "//Generated by DeclareModelBody"
    lines.Add DeclareModelBody
    DeclareModelBody = lines.NewLineJoin
    
    CopyToClipboard DeclareModelBody
    
End Function

Public Function GetCompleteModelInterface(frm As Object, SeqModelID) As String
    
    RunCommandSaveRecord
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If isFalse(SeqModelID) Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelID"
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    ''TABLE: qrySeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName|DataType|DataTypeInterface
    Set rs = ReturnRecordset("SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID)
    Dim ModelFieldsArr As New clsArray
    Do Until rs.EOF
        ''should produce //name: string;
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        ModelFieldsArr.Add fieldName & ": " & DataTypeInterface & ";"
        rs.MoveNext
    Loop
    
    If Not isFalse(SlugField) Then
        ModelFieldsArr.Add "slug: string;"
    End If
    
    Dim relatedPluralizedModelNames As New clsArray
    ''Add any relationships here (only simple relationship for now)...
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & _
        "AND RightModelID = " & SeqModelID & " AND Relationship = ""1:M"""
    Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        Dim IsSimpleRelationship: IsSimpleRelationship = rs.fields("IsSimpleRelationship")
        Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName")
        relatedPluralizedModelNames.Add LeftPluralizedModelName
        If IsSimpleRelationship Then
            ''Should produce string like this -> CardCardKeywords: string[];
            ModelFieldsArr.Add LeftPluralizedModelName & ": string[];"
        Else
            ModelFieldsArr.Add LeftPluralizedModelName & ": " & DeclareModelBody(frm, LeftModelID) & "[]; //Generated by DeclareModelBody"
        End If
        rs.MoveNext
    Loop
    
    Dim ModelFields: ModelFields = ModelFieldsArr.JoinArr(vbNewLine)
    
    ''TABLE: qrySeqModelFilters Fields: SeqModelFilterID|SeqModelID|SeqModelFieldID|IsMultiple|FilterQueryName
    ''FilterOperator|Timestamp|CreatedBy|RecordImportID|DatabaseFieldName|FieldName|SeqDataTypeID|DataType
    Set rs = ReturnRecordset("SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " ORDER BY FilterOrder")
    Dim ModelQueriesArr As New clsArray
    Dim uniqueFilters As New clsArray
    Do Until rs.EOF
        ''should produce //name: string;
        Dim FilterOperator: FilterOperator = rs.fields("FilterOperator"): If ExitIfTrue(isFalse(FilterOperator), "FilterOperator is empty..") Then Exit Function
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
        If Not uniqueFilters.InArray(FilterQueryName) Then
            uniqueFilters.Add FilterQueryName, True
            If FilterOperator = "Between" And DataType = "DATEONLY" Then
                ModelQueriesArr.Add "start_" & FilterQueryName & ": string;"
                ModelQueriesArr.Add "end_" & FilterQueryName & ": string;"
            Else
                ModelQueriesArr.Add FilterQueryName & ": string;"
            End If
        End If
        rs.MoveNext
    Loop
    
    Dim ModelQueries: ModelQueries = ModelQueriesArr.JoinArr(vbNewLine)
    
    Dim arrayItem
    Dim ModelFieldsForUpdateArr As New clsArray
    If relatedPluralizedModelNames.count > 0 Then
        For Each arrayItem In relatedPluralizedModelNames.arr
            ModelFieldsForUpdateArr.Add "deleted" & arrayItem & ": string[];"
        Next arrayItem
    End If
    
    Dim ModelFieldsForUpdate: ModelFieldsForUpdate = ModelFieldsForUpdateArr.JoinArr(vbNewLine)
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Backend Model Interface")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[ModelFields]", ModelFields)
    replacedContent = replace(replacedContent, "[ModelQueries]", ModelQueries)
    replacedContent = replace(replacedContent, "[ModelFieldsForUpdate]", ModelFieldsForUpdate)
     
    GetCompleteModelInterface = replacedContent
    Dim lines As New clsArray
    lines.Add "//Generated by GetCompleteModelInterface"
    lines.Add GetCompleteModelInterface
    GetCompleteModelInterface = lines.NewLineJoin
    
    CopyToClipboard GetCompleteModelInterface
    
End Function

Public Function GenerateBackendModelInterfaceForListForm(frm As Object, Optional SeqModelID = "") As String
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If isFalse(SeqModelID) Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelID"
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim rs2 As Recordset: Set rs2 = rs
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim InterfaceFileName: InterfaceFileName = rs.fields("InterfaceFileName"): If ExitIfTrue(isFalse(InterfaceFileName), "InterfaceFileName is empty..") Then Exit Function
    
    ''TABLE: qrySeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName|DataType|DataTypeInterface
    Set rs = ReturnRecordset("SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID)
    Dim ModelFieldsArr As New clsArray
    Do Until rs.EOF
        ''should produce //name: string;
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        ModelFieldsArr.Add fieldName & ": string;"
        rs.MoveNext
    Loop
    
    If Not isFalse(SlugField) Then
        ModelFieldsArr.Add "slug : string;"
    End If
    
    Dim ModelFields: ModelFields = ModelFieldsArr.JoinArr(vbNewLine)
    
    ''TABLE: qrySeqModelFilters Fields: SeqModelFilterID|SeqModelID|SeqModelFieldID|IsMultiple|FilterQueryName
    ''FilterOperator|Timestamp|CreatedBy|RecordImportID|DatabaseFieldName|FieldName|SeqDataTypeID|DataType
    Set rs = ReturnRecordset("SELECT FilterQueryName FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " GROUP BY FilterQueryName")
    Dim ModelQueriesArr As New clsArray
    Do Until rs.EOF
        ''should produce //name: string;
        ModelQueriesArr.Add rs.fields("FilterQueryName") & ": string;"
        rs.MoveNext
    Loop
    
    Dim ModelQueries: ModelQueries = ModelQueriesArr.JoinArr(vbNewLine)
    
    Dim replacedContent: replacedContent = GetReplacedTemplate(rs2, "Backend Model Interface For List Form")
    replacedContent = replace(replacedContent, "[Fields]", ModelFields)
    replacedContent = replace(replacedContent, "[ModelQueries]", ModelQueries)
     
    GenerateBackendModelInterfaceForListForm = replacedContent
    GenerateBackendModelInterfaceForListForm = GetGeneratedByFunctionSnippet(GenerateBackendModelInterfaceForListForm, "GenerateBackendModelInterfaceForListForm")
    CopyToClipboard GenerateBackendModelInterfaceForListForm
    
    Dim ProjectPath: ProjectPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ProjectPath")
    Dim filePath: filePath = ProjectPath & "src\interfaces\" & InterfaceFileName
    
    WriteToFile filePath, GenerateBackendModelInterfaceForListForm, SeqModelID
    
End Function

Public Function CopyCompleteControllerFile(frm As Form)
    
    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    If isFalse(SeqModelID) Then Exit Function
    
    ImportCompleteControllerFile frm, SeqModelID
    
End Function

Public Function GenerateControllerForListForm(frm As Object, Optional SeqModelID = Null)
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''Get the path and file name of the model
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim TemplateContent: TemplateContent = GetReplacedTemplate(rs, "Controller For List Form")
    
    ''updateModels,getModels,getModel
    ''Generate_updateModelsFunction
    ''Generate_getModelsSimpleFilter
    ''GetGenericController("get", ModelName, "GetOne", FindOptions)
    Dim updateModels: updateModels = Generate_updateModelsFunction(frm, SeqModelID)
    Dim getModels: getModels = Generate_getModelsSimpleFilter(frm, SeqModelID)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim findOptions: findOptions = GeneratefindOptions(frm, SeqModelID)
    Dim getModel: getModel = GetGenericController("get", ModelName, "GetOne", Generate_findOptionsCopy(frm, SeqModelID))
    
    ''Relationships
    Dim item, relatedModels As New clsArray: relatedModels.arr = Elookups("qrySeqModelRelationships", "RightModelID = " & SeqModelID, "LeftModelID")
    Dim importRelatedModels As New clsArray
    For Each item In relatedModels.arr
        importRelatedModels.Add ImportAsRelatedModelBackend(frm, item)
    Next item
    
    relatedModels.arr = Elookups("qrySeqModelRelationships", "LeftModelID = " & SeqModelID, "RightModelID")
    For Each item In relatedModels.arr
        importRelatedModels.Add ImportAsRelatedModelBackend(frm, item)
    Next item
    
    TemplateContent = replace(TemplateContent, "[FindOptions]", findOptions)
    TemplateContent = replace(TemplateContent, "[updateModels]", updateModels)
    TemplateContent = replace(TemplateContent, "[getModels]", getModels)
    TemplateContent = replace(TemplateContent, "[getModel]", getModel)
    TemplateContent = replace(TemplateContent, "[importRelatedModels]", importRelatedModels.JoinArr(vbNewLine))
    
    GenerateControllerForListForm = TemplateContent
    
    GenerateControllerForListForm = GetGeneratedByFunctionSnippet(GenerateControllerForListForm, "GenerateControllerForListForm")
    
    CopyToClipboard GenerateControllerForListForm
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    Dim ProjectPath: ProjectPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ProjectPath")
    If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    
    Dim ControllerFileName: ControllerFileName = rs.fields("ControllerFileName"): If ExitIfTrue(isFalse(ControllerFileName), "ControllerFileName is empty..") Then Exit Function
    Dim filePath: filePath = ProjectPath & "src\controllers\" & ControllerFileName
    
    WriteToFile filePath, GenerateControllerForListForm, SeqModelID
   
End Function

Public Function ImportCompleteControllerFile(frm As Object, Optional SeqModelID = Null)
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''Get the path and file name of the model
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim ControllerFileName: ControllerFileName = rs.fields("ControllerFileName"): If ExitIfTrue(isFalse(ControllerFileName), "ControllerFileName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    
    Dim findOptions: findOptions = GeneratefindOptions(frm, SeqModelID)
    Dim findOptionsCopy: findOptionsCopy = "const findOptionsCopy: FindOptions<typeof " & ModelName & "> = cloneDeep(findOptions);"
    
    Dim lines As New clsArray
    lines.Add "//Generated by ImportCompleteControllerFile"
    lines.Add "import { Request, Response } from ""express"";"
    lines.Add "import { " & ModelName & " } from ""../models/" & ModelName & "Model"";"
    lines.Add "import {genericAdd,genericDelete,genericGetAll,genericGetOne,genericUpdate,genericGetOneBySlug,genericGetAndCountAll} from ""../utils/generic"";"
    lines.Add "import { FindOptions, Op } from ""sequelize"";"
    lines.Add "import handleSequelizeError from ""../utils/errorHandling"";"
    lines.Add "import { convertDateStringToYYYYMMDD, convertStringToFloat, formatSortAsSequelize, getSort, isValidPage, returnJSONResponse } from ""../utils/utils"";"
    lines.Add "import sequelize from ""../config/db"";"
    lines.Add "import { cloneDeep } from ""lodash"";"
    lines.Add "import { " & ModelName & "Body," & ModelName & "BodyForUpdate," & ModelName & "Query } from ""../interfaces/" & ModelName & "Interfaces"";"
    
    ''Relationships where the SeqModelID is a RightModelID
    Dim item, relatedModels As New clsArray: relatedModels.arr = Elookups("qrySeqModelRelationships", "RightModelID = " & SeqModelID, "LeftModelID")
    Dim hasRelationship As Boolean
    For Each item In relatedModels.arr
        hasRelationship = True
        lines.Add ImportAsRelatedModelBackend(frm, item)
    Next item
    
    ''Relationships where the SeqModelID is a LeftModelID
    relatedModels.arr = Elookups("qrySeqModelRelationships", "LeftModelID = " & SeqModelID, "RightModelID")
    For Each item In relatedModels.arr
        hasRelationship = True
        lines.Add ImportAsRelatedModelBackend(frm, item)
    Next item
    
    lines.Add ""
    lines.Add "const ModelObject = " & ModelName & ";"
    lines.Add ""
    lines.Add findOptions
    
    lines.Add GetAddFunctionWithRelationship(frm, SeqModelID)
    
    lines.Add ""
    
    lines.Add GetUpdateFunctionWithRelationship(frm, SeqModelID)
    
    lines.Add ""
    lines.Add Generate_getModelsSimpleFilter(frm, SeqModelID)
    lines.Add ""
    
    If Not isFalse(SlugField) Then
        lines.Add GetGenericController("get", ModelName, "GetOneBySlug", findOptionsCopy)
    Else
        lines.Add GetGenericController("get", ModelName, "GetOne", findOptionsCopy)
    End If
    
    lines.Add ""
    lines.Add GetGenericController("delete", ModelName, "Delete")
    
    Dim ControllerFileContent: ControllerFileContent = lines.JoinArr(vbCrLf)
    
    Dim ProjectPath: ProjectPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ProjectPath")
    Dim filePath: filePath = ProjectPath & "src\controllers\" & ControllerFileName
    
    ImportCompleteControllerFile = ControllerFileContent
    CopyToClipboard ImportCompleteControllerFile
    WriteToFile filePath, ImportCompleteControllerFile, SeqModelID
    
End Function

Private Function GetGenericController(arg1, arg2, arg3, Optional findOptions = "") As String

    Dim lines As New clsArray
    lines.Add "//Generated by GetGenericController"
    lines.Add "export const " & arg1 & arg2 & " = async (req: Request, res: Response) => {"
    
    If Not isFalse(findOptions) Then lines.Add findOptions
    
    Dim genericLine: genericLine = "generic" & arg3 & "(req, res, ModelObject"
    Select Case arg3
        Case "GetAll":
            genericLine = genericLine & ",findOptionsCopy"
        Case "GetOne":
            genericLine = genericLine & ",findOptionsCopy"
        Case "GetOneBySlug":
            genericLine = genericLine & ",findOptionsCopy"
    End Select
    
    lines.Add genericLine & ");"
    lines.Add "};"
    
    GetGenericController = lines.JoinArr(vbCrLf)
    
End Function

Public Function ImportCompleteRouteFile(frm As Object, Optional SeqModelID = Null)
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''Get the path and file name of the model
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim RouteFileName: RouteFileName = rs.fields("RouteFileName"): If ExitIfTrue(isFalse(RouteFileName), "RouteFileName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    
    Dim lines As New clsArray
    lines.Add GetControllerImports(ModelName, PluralizedModelName)
    lines.Add "import { Router } from ""express"";"
    lines.Add ""
    lines.Add "const router = Router();"
    lines.Add ""
    lines.Add "router.route(" & Esc("/" & ModelPath) & ").get(get" & PluralizedModelName & ").post(add" & ModelName & ");"
    lines.Add "router.route(" & Esc("/" & ModelPath & "/:id") & ").get(get" & ModelName & ").put(update" & ModelName & ").delete(delete" & ModelName & ");"
    lines.Add ""
    lines.Add "export default router;"
    
    Dim RouteFileContent: RouteFileContent = lines.JoinArr(vbCrLf)
    
    Dim ProjectPath: ProjectPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ProjectPath")
    Dim filePath: filePath = ProjectPath & "src\routes\" & RouteFileName
    
    ImportCompleteRouteFile = RouteFileContent
    CopyToClipboard ImportCompleteRouteFile
    WriteToFile filePath, ImportCompleteRouteFile, SeqModelID
    
End Function

Private Function GetControllerImports(ModelName, PluralizedModelName)
    
    GetControllerImports = "import { add" & ModelName & ", delete" & ModelName & ", get" & ModelName & "," & _
            "get" & PluralizedModelName & ", update" & ModelName & " } from ""../controllers/" & ModelName & "Controller"""
            
End Function

Public Function ConvertToModelPath(ByVal str As String, Optional dontPluralized As Boolean = False) As String
    Dim output As String
    Dim i As Integer
    
    For i = 1 To Len(str)
        Select Case Mid(str, i, 1)
            Case "A" To "Z"
                If i > 1 Then
                    If Mid(str, i - 1, 1) <> "-" Then
                        output = output & "-"
                    End If
                End If
                output = output & LCase(Mid(str, i, 1))
            Case Else
                output = output & Mid(str, i, 1)
        End Select
    Next i
    
    If Not dontPluralized Then
        If Right(output, 1) = "y" Then
            output = Left(output, Len(output) - 1) & "ies"
        ElseIf Right(output, 1) = "s" Then
            output = output & "es"
        Else
            output = output & "s"
        End If
    End If
    
    ConvertToModelPath = replace(output, " ", "")
End Function

Public Function GenerateIndexRoute(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath
    
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName")
    Dim LModelName: LModelName = LCase(ModelName)
    
    Dim lines As New clsArray
    lines.Add "import " & VariableName & "Route from ""./routes/" & ModelName & "Route"""
    lines.Add "app.use(""/"", " & VariableName & "Route)"
    
    Dim indexRoute: indexRoute = lines.JoinArr(vbCrLf)
    GenerateIndexRoute = indexRoute
    CopyToClipboard indexRoute
    
End Function

Public Function GenerateAttributesOption(frm As Object, Optional SeqModelID = Null, Optional FromRelationship As Boolean = False)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim SlugField: SlugField = ELookup("tblSeqModels", "SeqModelID = " & SeqModelID, "SlugField")
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT IsExpression"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim fieldName: fieldName = rs.fields("FieldName")
        fields.Add Esc(fieldName)
        rs.MoveNext
    Loop
    
    If Not isFalse(SlugField) Then fields.Add Esc("slug")
    
    If Not FromRelationship Then
        Dim ExpressionLiteralFields: ExpressionLiteralFields = GetAllExpressionLiteralFieldBySeqModel(frm, SeqModelID)
        If Not isFalse(ExpressionLiteralFields) Then
            fields.Add ExpressionLiteralFields
        End If
    End If
    
    Dim fieldsStr: fieldsStr = "[" & fields.JoinArr(",") & "]"

    Dim dict As New clsDictionary
    dict.Add "attributes", fieldsStr
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateAttributesOption"
    lines.Add dict.ToFormatString(True)

    GenerateAttributesOption = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateAttributesOption
    
End Function

Public Function GeneratefindAllOptions(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath
    
    ''Combine into one dictionary
    Dim findAllOptions As New clsArray
    
    ''Get the attributes option
    Dim attributes: attributes = GenerateAttributesOption(frm, SeqModelID)
    findAllOptions.Add attributes
    
    
    GeneratefindAllOptions = "const findAllOptions = {" & findAllOptions.JoinArr("," & vbCrLf) & "}"
    CopyToClipboard GeneratefindAllOptions
    
End Function

Public Function GeneratefindOptions(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath
    
    ''Combine into one dictionary
    Dim findOptions As New clsArray
    
    ''Get the include option
    Dim IncludeOption: IncludeOption = GenerateIncludeOption(frm, SeqModelID)
    findOptions.Add IncludeOption
    
    ''Get the attributes option
    Dim attributes: attributes = GenerateAttributesOption(frm, SeqModelID)
    findOptions.Add attributes
    
    Dim lines As New clsArray
    lines.Add "//Generated by GeneratefindOptions"
    lines.Add "const findOptions: FindOptions<typeof " & ModelName & "> = {" & findOptions.JoinArr("," & vbCrLf) & "}"
    
    GeneratefindOptions = lines.JoinArr(vbNewLine)
    CopyToClipboard GeneratefindOptions
    
End Function

Public Function GenerateGetSQLFunction(frm As Object, Optional SeqModelID = Null)
    
    RunCommandSaveRecord
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fieldsAsJS: fieldsAsJS = GenerateSQLFieldList(frm, SeqModelID)
    
    Dim ModelName: ModelName = rs.fields("ModelName")
    Dim TableName: TableName = rs.fields("TableName")
    Dim VariableName: VariableName = rs.fields("VariableName")
    Dim IsMainQuery: IsMainQuery = rs.fields("IsMainQuery")
    Dim MainQuery: MainQuery = IIf(IsMainQuery, "true", "false")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("getSQL")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[TableName]", TableName)
    replacedContent = replace(replacedContent, "[ModelFields]", fieldsAsJS)
    replacedContent = replace(replacedContent, "[VariableName]", ModelName)
    replacedContent = replace(replacedContent, "[Filters]", GenerateSeqModelFilters(frm, SeqModelID))
    replacedContent = replace(replacedContent, "[MainQuery]", MainQuery)
    
    GenerateGetSQLFunction = replacedContent
    GenerateGetSQLFunction = GetGeneratedByFunctionSnippet(GenerateGetSQLFunction, "GenerateGetSQLFunction")
    
    ''ModelName,TableName,ModelFields,VariableName,Filters
    CopyToClipboard GenerateGetSQLFunction
    
End Function

Public Function GenerateSQLFieldList(frm, Optional SeqModelID = "")

    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Not IsExpression  ORDER BY SeqModelFieldID"
    ''TABLE: tblSeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    
    If Not rs.EOF Then
        Dim SlugField: SlugField = rs.fields("SlugField")
        If Not isFalse(SlugField) Then
            fields.Add Esc("slug")
        End If
    End If
    
    Do Until rs.EOF
        Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName")
        Dim fieldName: fieldName = rs.fields("FieldName")
        If fieldName <> DatabaseFieldName Then
            fields.Add "[" & Esc(DatabaseFieldName) & "," & Esc(fieldName) & "]"
        Else
            fields.Add Esc(rs.fields("DatabaseFieldName"))
        End If

        rs.MoveNext
    Loop
    
    GenerateSQLFieldList = "[" & fields.JoinArr(",") & "]"
    GenerateSQLFieldList = GetGeneratedByFunctionSnippet(GenerateSQLFieldList, "GenerateSQLFieldList")
    
    CopyToClipboard GenerateSQLFieldList
    
    
End Function

Public Function GenerateMainGetSQLFunction(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString
    
    Dim fieldsAsJS: fieldsAsJS = GenerateSQLFieldList(frm, SeqModelID)
    
    Dim ModelName: ModelName = rs.fields("ModelName")
    Dim TableName: TableName = rs.fields("TableName")
    Dim VariableName: VariableName = rs.fields("VariableName")
    Dim IsMainQuery: IsMainQuery = rs.fields("IsMainQuery")
    Dim MainQuery: MainQuery = IIf(IsMainQuery, "true", "false")
    Dim SortString: SortString = rs.fields("SortString")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Main getSQL")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[TableName]", TableName)
    replacedContent = replace(replacedContent, "[ModelFields]", fieldsAsJS)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[Filters]", GenerateSeqModelFilters(frm, SeqModelID))
    replacedContent = replace(replacedContent, "[SortString]", SortString)
    
    ''Replace the [COUNT SQL] with the template content from
    replacedContent = replace(replacedContent, "[COUNT SQL]", GetCountSQL(frm, SeqModelID))
    
    ''Replace the [LIMIT AND OFFSET] with the template content from
    replacedContent = replace(replacedContent, "[LIMIT AND OFFSET]", GetTemplateContent("Limit and Offset"))
    
    ''Replace the [DISTINCT SQL] with the template content from
    replacedContent = replace(replacedContent, "[DISTINCT SQL]", GetDistinctSQL(frm, SeqModelID))
    
    GenerateMainGetSQLFunction = replacedContent
    
    ''ModelName,TableName,ModelFields,VariableName,Filters
    CopyToClipboard GenerateMainGetSQLFunction
    
End Function

Public Function GenerateSimpleJoin(frm As Object, Optional SeqModelID = Null) As String
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString
    Dim VariableName: VariableName = rs.fields("VariableName")
    Dim TableName: TableName = rs.fields("TableName")
    Dim LeftKey: LeftKey = rs.fields("LeftKey")
    Dim RightKey: RightKey = rs.fields("RightKey")
    Dim ModelName: ModelName = rs.fields("ModelName")
    
    Dim lines As New clsArray
    
    ''Generate the fields with aliases
    sqlStr = "SELECT * FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    ''TABLE: tblSeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName")
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim field: field = TableName & "." & DatabaseFieldName & " AS `" & ModelName & "." & fieldName & "`"
        fields.Add Esc(field)
        rs.MoveNext
    Loop
    
    lines.Add "sql.fields = sql.fields.concat([" & fields.JoinArr & "])" ''sql.fields = sql.fields.concat(["marvelduel_belongsto.id As `deck.id`", "marvelduel_belongsto.name As `deck.name`"])
    
    ''const deckJoin = new clsJoin("marvelduel_belongsto", "deck_id", "id", null, "LEFT")
    lines.Add "const " & VariableName & "Join = new clsJoin(" & Esc(TableName) & "," & Esc(LeftKey) & "," & Esc(RightKey) & ", """", ""LEFT"")"
    
    GenerateSimpleJoin = lines.JoinArr(vbCrLf)
    CopyToClipboard GenerateSimpleJoin
    
End Function

Public Function GetJoinCancellation(frm As Object, Optional SeqModelID = Null)

    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString
    
    Dim VariableName: VariableName = rs.fields("VariableName")
    Dim ModelName: ModelName = rs.fields("ModelName")
    
    ''cardCardKeyword_SQL = getCardCardKeywordSQL(query, true).sql;
    ''cardCardKeywordJoin.source = cardCardKeyword_SQL.sql();
    ''cardCardKeywordJoin.joinType = "LEFT";
    
    Dim lines As New clsArray
    lines.Add VariableName & "_SQL = get" & ModelName & "SQL(query,true).sql;"
    lines.Add VariableName & "Join.source = " & VariableName & "_SQL.sql();"
    lines.Add VariableName & "Join.joinType = ""LEFT"";"
    
    GetJoinCancellation = lines.JoinArr(vbCrLf)
    CopyToClipboard GetJoinCancellation
    
End Function

Public Function GenerateFieldsAsJSList(frm As Object, Optional SeqModelID = Null)

    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelFieldID"
    ''TABLE: tblSeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    Do Until rs.EOF
        fields.Add Esc(rs.fields("DatabaseFieldName"))
        rs.MoveNext
    Loop
    
    GenerateFieldsAsJSList = "[" & fields.JoinArr(",") & "]"
    
    CopyToClipboard GenerateFieldsAsJSList
    
End Function

Public Function SeqModelDEFormOnCurrent(frm As Form)

    SetFocusOnForm frm, ""
    Dim SeqModelID: SeqModelID = frm("SeqModelID")
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    
    Dim sqlStr: sqlStr = "SELECT SeqModelFieldID,DatabaseFieldName FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " ORDER BY DatabaseFieldName ASC"
    
    frm("subSeqModelFilters").Form("SeqModelFieldID").RowSource = sqlStr
    frm("subSeqModelSorts").Form("SeqModelFieldID").RowSource = sqlStr
    
    sqlStr = "SELECT SeqModelRelationshipID, LeftModelName FROM qrySeqModelRelationships WHERE RightModelID = " & SeqModelID & " ORDER BY RightModelName ASC"
    frm("subSeqModelFilters").Form.controls("SeqModelRelationshipID").RowSource = sqlStr
    frm("subSeqModelEmbeddings").Form("SeqModelRelationshipID").RowSource = sqlStr
    
    sqlStr = "SELECT SeqModelID, ModelName FROM tblSeqModels WHERE BackendProjectID = " & BackendProjectID & " ORDER BY ModelName ASC"
    frm("subSeqModelFields").Form("RelatedModelID").RowSource = sqlStr
    frm("subSeqModelFilters").Form("ModelListID").RowSource = sqlStr
    frm("subSeqModelEmbeddings").Form("RelatedModelID").RowSource = sqlStr
    
    sqlStr = "SELECT SeqModelFieldGroupID, GroupName FROM tblSeqModelFieldGroups WHERE SeqModelID = " & SeqModelID & " ORDER BY GroupOrder"
    frm("subSeqModelFields").Form("SeqModelFieldGroupID").RowSource = sqlStr
'    frm("subSeqModelFields").Form("FieldName").SetFocus
'    DoCmd.RunCommand acCmdFreezeColumn
    
    sqlStr = "Select SeqModelID, ModelName FROm tblSeqModels WHERE BackendProjectID = " & BackendProjectID & " ORDER BY ModelName"
    frm("SameModelD").RowSource = sqlStr
    
End Function

Public Function GenerateSeqModelFilters(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim SeqModelFilters As New clsArray
    
    Dim LikeFilters: LikeFilters = GetLikeFilters(frm, SeqModelID)
    If Not isFalse(LikeFilters) Then SeqModelFilters.Add LikeFilters
    
    Dim MatchFilters: MatchFilters = GetMatchFilters(frm, SeqModelID)
    If Not isFalse(MatchFilters) Then SeqModelFilters.Add MatchFilters
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND NOT " & _
        "FilterOperator = ""LIKE"" ORDER BY SeqModelFilterID"
        
    ''Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND NOT " & _
        "FilterOperator = ""LIKE"" AND NOT SeqModelFieldID IS NULL ORDER BY SeqModelFilterID"
    ''TABLE: tblSeqModelFilters Fields: SeqModelFilterID|SeqModelID|SeqModelFieldID|IsMultiple|FilterQueryName
    ''FilterOperator|Timestamp|CreatedBy|RecordImportID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        SeqModelFilters.Add GenerateModelFilterSnippet(frm, rs.fields("SeqModelFilterID"))
        rs.MoveNext
    Loop
    
    GenerateSeqModelFilters = SeqModelFilters.JoinArr(vbCrLf & vbCrLf)
    GenerateSeqModelFilters = GetGeneratedByFunctionSnippet(GenerateSeqModelFilters, "GenerateSeqModelFilters")
    CopyToClipboard GenerateSeqModelFilters
    
End Function

Public Function GenerateModelJoin(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelID"
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim VariableName: VariableName = rs.fields("VariableName")
    Dim ModelName: ModelName = rs.fields("ModelName")
    Dim LeftKey: LeftKey = rs.fields("LeftKey")
    Dim RightKey: RightKey = rs.fields("RightKey")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Model Join")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[LeftKey]", LeftKey)
    replacedContent = replace(replacedContent, "[RightKey]", RightKey)
     
    GenerateModelJoin = replacedContent
    GenerateModelJoin = GetGeneratedByFunctionSnippet(GenerateModelJoin, "GenerateModelJoin")
    CopyToClipboard GenerateModelJoin
    
End Function

Public Function GenerateModelJoinFromMain(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelID"
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim VariableName: VariableName = rs.fields("VariableName")
    Dim ModelName: ModelName = rs.fields("ModelName")
    Dim LeftKey: LeftKey = rs.fields("LeftKey")
    Dim RightKey: RightKey = rs.fields("RightKey")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Model Join From Main")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[LeftKey]", LeftKey)
    replacedContent = replace(replacedContent, "[RightKey]", RightKey)
     
    GenerateModelJoinFromMain = replacedContent
    CopyToClipboard GenerateModelJoinFromMain
    
End Function

Public Function GenerateGetModelRecords(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelID"
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName")
    Dim ModelName: ModelName = rs.fields("ModelName")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Get Model Records")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[PluralizedModelName]", PluralizedModelName)
     
    GenerateGetModelRecords = replacedContent
    CopyToClipboard GenerateGetModelRecords
    
End Function

Public Function GetCountSQL(frm As Object, Optional SeqModelID = Null)

    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " AND PrimaryKey ORDER BY SeqModelID"
    ''TABLE: tblSeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then Exit Function
    Dim PrimaryKey: PrimaryKey = rs.fields("DatabaseFieldName")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Count SQL")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[PrimaryKey]", PrimaryKey)

    GetCountSQL = replacedContent
    CopyToClipboard GetCountSQL
    
End Function

Public Function GetDistinctSQL(frm As Object, Optional SeqModelID = Null)

    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " AND PrimaryKey ORDER BY SeqModelID"
    ''TABLE: tblSeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then Exit Function
    Dim PrimaryKey: PrimaryKey = rs.fields("DatabaseFieldName")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Distinct SQL")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[PrimaryKey]", PrimaryKey)

    GetDistinctSQL = replacedContent
    CopyToClipboard GetDistinctSQL
    
End Function

Public Function Generate_GetAddFunctionWithRelationshipNext13(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim content
    content = GetReplacedTemplate(rs, "Add With Relationship Next 13")
    content = replace(content, "[RelationshipBodyDeclaration]", getRelationshipBodyDeclarationForAdd(frm, SeqModelID))
    content = replace(content, "[MainModelUniquenessValidation]", getMainModelUniquenessValidationForAdd(frm, SeqModelID))
    content = replace(content, "[EnumFieldsValidation]", GenerateEnumValidation(frm, SeqModelID))
    content = replace(content, "[ChildrenUniquenessValidation]", getChildrenUniquenessValidationForAdd(frm, SeqModelID))
    content = replace(content, "[ChildrenUniquenessValidationWithDatabase]", getChildrenUniquenessValidationWithDatabaseForAdd(frm, SeqModelID))
    content = replace(content, "[ModelFieldsValue]", getModelFieldsValue(frm, SeqModelID))
    content = replace(content, "[ChildrenInserts]", GetChildrenInsertsForAdd(frm, SeqModelID))
    content = replace(content, "[GetBackendModelRequiredSnippets]", GetBackendModelRequiredSnippets(frm, SeqModelID))
    
    Dim lines As New clsArray
    lines.Add "//Generated by Generate_GetAddFunctionWithRelationshipNext13"
    lines.Add content
    
    Generate_GetAddFunctionWithRelationshipNext13 = lines.JoinArr(vbNewLine)
    CopyToClipboard Generate_GetAddFunctionWithRelationshipNext13
    
    
End Function

Public Function GetAddFunctionWithRelationship(frm As Object, Optional SeqModelID = Null)
    
    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim content
    content = GetReplacedTemplate(rs, "Add With Relationship")
    content = replace(content, "[RelationshipBodyDeclaration]", getRelationshipBodyDeclarationForAdd(frm, SeqModelID))
    content = replace(content, "[MainModelUniquenessValidation]", getMainModelUniquenessValidationForAdd(frm, SeqModelID))
    content = replace(content, "[EnumFieldsValidation]", GenerateEnumValidation(frm, SeqModelID))
    content = replace(content, "[ChildrenUniquenessValidation]", getChildrenUniquenessValidationForAdd(frm, SeqModelID))
    content = replace(content, "[ChildrenUniquenessValidationWithDatabase]", getChildrenUniquenessValidationWithDatabaseForAdd(frm, SeqModelID))
    content = replace(content, "[ModelFieldsValue]", getModelFieldsValue(frm, SeqModelID))
    content = replace(content, "[ChildrenInserts]", GetChildrenInsertsForAdd(frm, SeqModelID))
    
    Dim lines As New clsArray
    lines.Add "//Generated by GetAddFunctionWithRelationship"
    lines.Add content
    
    GetAddFunctionWithRelationship = lines.JoinArr(vbNewLine)
    CopyToClipboard GetAddFunctionWithRelationship
    
End Function

Public Function GetUpdateFunctionWithRelationship(frm As Object, Optional SeqModelID = Null)

    If IsNull(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    RunCommandSaveRecord
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName
    Dim ModelName: ModelName = rs.fields("ModelName")
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID")
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND RightModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim RelationshipBodyDeclarations As New clsArray, RelationshipInserts As New clsArray
    Dim RelationshipUniquenessValidations As New clsArray, RelationshipDeletions As New clsArray
    Dim TemplateContent, replacedContent
    Dim fieldName
    Do Until rs.EOF
        Dim IsSimpleRelationship: IsSimpleRelationship = rs.fields("IsSimpleRelationship")
        Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName")
        Dim LeftModelName: LeftModelName = rs.fields("LeftModelName")
        Dim FieldToBeInserted: FieldToBeInserted = rs.fields("FieldToBeInserted")
        Dim LeftForeignKey: LeftForeignKey = rs.fields("LeftForeignKey")
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        If IsSimpleRelationship Then
            ''must produce something like this const CardCardKeywords: string[] = body.CardCardKeywords || [];
            ''must produce -> const deletedCardCardKeywords = body.deletedCardCardKeywords || [];
            ''must produce -> if(CardCardKeywords.length>0){CardCardKeywords.map((cardKeywordId)=>{promises.push(CardCardKeyword.create({cardId:id,cardKeywordId:cardKeywordId},{transaction:t}));});}
            ''must produce -> if(deletedCardCardKeywords.length>0){deletedCardCardKeywords.map((id)=>{promises.push(CardCardKeyword.destroy({where:{id},transaction:t}))})};
            RelationshipBodyDeclarations.Add "const " & LeftPluralizedModelName & ": string[] = body." & LeftPluralizedModelName & " || [];"
            RelationshipBodyDeclarations.Add "const deleted" & LeftPluralizedModelName & " = body.deleted" & LeftPluralizedModelName & " || [];"
            RelationshipInserts.Add "if(" & LeftPluralizedModelName & ".length>0){" & LeftPluralizedModelName & ".map((item)=>" & _
                "{promises.push(" & LeftModelName & ".create({" & LeftForeignKey & ":id," & FieldToBeInserted & ":item},{transaction:t}));});}"
            RelationshipDeletions.Add "if(deleted" & LeftPluralizedModelName & ".length>0){deleted" & LeftPluralizedModelName & ".map((id)=>{promises.push(" & _
                LeftModelName & ".destroy({where:{id},transaction:t}))})};"
        Else
            
            RelationshipBodyDeclarations.Add "const {" & LeftPluralizedModelName & ", deleted" & LeftPluralizedModelName & "} = body"
            
            RelationshipInserts.Add GenerateChildUpdateorInsert(frm, LeftModelID)
            
            RelationshipDeletions.Add GenerateChildDelete(frm, LeftModelID)
            
            RelationshipUniquenessValidations.Add GenerateUniquenessValidation(frm, LeftModelID)
            
        End If
        rs.MoveNext
    Loop
    
    Dim IncludeOption As String, RelationshipBodyDeclaration As String
    Dim RelationshipInsert As String, RelationshipDeletion As String
    Dim RelationshipUniquenessValidation
    
    If RelationshipBodyDeclarations.count > 0 Then
        RelationshipBodyDeclaration = RelationshipBodyDeclarations.JoinArr(vbNewLine)
        RelationshipInsert = RelationshipInserts.JoinArr(vbNewLine)
        RelationshipDeletion = RelationshipDeletions.JoinArr(vbNewLine)
        RelationshipUniquenessValidation = RelationshipUniquenessValidations.JoinArr(vbNewLine)
    End If
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey AND NOT isExpression ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Dim ModelFieldsArr As New clsArray, ModelFieldsValueArr As New clsArray
    Do Until rs.EOF
        fieldName = rs.fields("FieldName")
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID"): If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
        ModelFieldsArr.Add fieldName
        ModelFieldsValueArr.Add GenerateCreateUpdateField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    Dim ModelFields: ModelFields = ModelFieldsArr.JoinArr(",")
    Dim ModelFieldsValue: ModelFieldsValue = ModelFieldsValueArr.JoinArr(", //Generated By GenerateCreateUpdateField" & vbCrLf)
    
    TemplateContent = GetTemplateContent("Update With Relationship")
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[ModelFields]", ModelFields)
    replacedContent = replace(replacedContent, "[ModelFieldsValue]", ModelFieldsValue)
    replacedContent = replace(replacedContent, "[RelationshipBodyDeclaration]", RelationshipBodyDeclaration)
    replacedContent = replace(replacedContent, "[RelationshipInsert]", RelationshipInsert)
    replacedContent = replace(replacedContent, "[RelationshipDeletion]", RelationshipDeletion)
    replacedContent = replace(replacedContent, "[RelationshipUniquenessValidation]", RelationshipUniquenessValidation)
    replacedContent = replace(replacedContent, "[EnumValidation]", GenerateEnumValidation(frm, SeqModelID))
    replacedContent = replace(replacedContent, "[UniquenessValidation]", GenerateMainModelUniqueValidation(frm, SeqModelID))
    
    Dim lines As New clsArray
    lines.Add "//Generated by GetUpdateFunctionWithRelationship"
    lines.Add replacedContent
    
    GetUpdateFunctionWithRelationship = lines.JoinArr(vbNewLine)
    CopyToClipboard GetUpdateFunctionWithRelationship
    
End Function

Public Function GetPrimaryKeyField(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND PrimaryKey"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GetPrimaryKeyField = rs.fields("FieldName")
    
End Function

Public Function CopyModelPage(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ModelName: ModelName = rs.fields("ModelName")
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID")
    Dim VariableName: VariableName = rs.fields("VariableName")
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VerboseModelName: VerboseModelName = rs.fields("VerboseModelName"): If ExitIfTrue(isFalse(VerboseModelName), "VerboseModelName is empty..") Then Exit Function
    Dim SidebarEnabled: SidebarEnabled = isPresent("tblBackendProjects", "BackendProjectID = " & BackendProjectID & " AND SidebarEnabled")
    
    If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent(IIf(SidebarEnabled, "ModelList Page Sidebar", "List Page"))
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[PluralizedModelName]", PluralizedModelName)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[VerboseModelName]", VerboseModelName)
    
    CopyModelPage = replacedContent
    CopyToClipboard CopyModelPage
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\pages\" & ModelPath & "\index.tsx"
    
    WriteToFile filePath, CopyModelPage, SeqModelID
    
End Function

Public Function CopyClientChildInterfaces(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ModelName: ModelName = rs.fields("ModelName")
    If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID")
    If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND (RightModelID = " & SeqModelID & _
        " OR LeftModelID = " & SeqModelID & ") AND Relationship <> ""M:M"""
    Set rs = ReturnRecordset(sqlStr)
    
    ''Each item should produce something like this -> interface CardCardKeyword{id:number;cardId:number;cardKeywordId:number;CardKeyword:BasicModel;}
    Dim interfaceArray As New clsArray
    
    Do Until rs.EOF
        Dim RelatedModelName, RelatedModelID
        If rs.fields("RightModelID") = SeqModelID Then
            RelatedModelName = rs.fields("LeftModelName"): If ExitIfTrue(isFalse(RelatedModelName), "LeftModelName is empty..") Then Exit Function
            RelatedModelID = rs.fields("LeftModelID")
        Else
            RelatedModelName = rs.fields("RightModelName"): If ExitIfTrue(isFalse(RelatedModelName), "RightModelName is empty..") Then Exit Function
            RelatedModelID = rs.fields("RightModelID")
        End If
        
        interfaceArray.Add "interface " & RelatedModelName & GetModelInterfaceItems(RelatedModelID)
        rs.MoveNext
    Loop
    
    CopyClientChildInterfaces = interfaceArray.JoinArr(vbNewLine)
    CopyToClipboard CopyClientChildInterfaces
    
End Function

Private Function GetModelInterfaceItems(SeqModelID) As String
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    
    Dim fieldItems As New clsArray
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID
    ''TABLE: qrySeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName|PluralizedFieldName
    ''AllowedOptions|DataType|DataTypeInterface
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        Dim AllowedOptions: AllowedOptions = rs.fields("AllowedOptions")
        Dim AllowNull: AllowNull = rs.fields("AllowNull")
        
        sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND LeftForeignKey = " & Esc(fieldName)
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        
        If DataType = "ENUM" Then
            Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
            Dim EnumOptionStr: EnumOptionStr = ConvertEnumToArray(DataTypeOption).JoinArr(" | ")
            ''If optional then add a null at the end of the items
            If AllowNull Then
                EnumOptionStr = EnumOptionStr & " | null"
            End If
            
            fieldItems.Add fieldName & ": " & EnumOptionStr
            
        ElseIf Not isFalse(AllowedOptions) Then
            Dim options As New clsArray: options.arr = AllowedOptions
            
            If DataTypeInterface = "string" Then
                options.EscapeItems
            End If
            
            Dim optionsStr: optionsStr = options.JoinArr(" | ")
            If AllowNull Then
                optionsStr = optionsStr & " | null"
            End If
            
            fieldItems.Add fieldName & ": " & optionsStr
            
'        ElseIf Not rs2.EOF Then
'
'            Dim RightVariablePluralName: RightVariablePluralName = rs2.fields("RightVariablePluralName"): If ExitIfTrue(isFalse(RightVariablePluralName), "RightVariablePluralName is empty..") Then Exit Function
'            fieldItems.Add fieldName & ": " & RightVariablePluralName & "[0]"
            
        Else
            If AllowNull Then
                DataTypeInterface = DataTypeInterface & " | null"
            End If
            
            fieldItems.Add fieldName & ": " & DataTypeInterface
        End If
        
        rs.MoveNext
    Loop
    
    If Not IsNull(SlugField) Then fieldItems.Add "slug : string"
    
    ''Any related model of this SeqModelID
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND (RightModelID = " & SeqModelID & _
        " OR LeftModelID = " & SeqModelID & ") AND Relationship <> ""M:M"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim InterfaceFieldDeclaration
        Dim Relationship: Relationship = rs.fields("Relationship"): If ExitIfTrue(isFalse(Relationship), "Relationship is empty..") Then Exit Function
        If rs.fields("RightModelID") = SeqModelID Then
            Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
            Dim LeftModelName: LeftModelName = rs.fields("LeftModelName"): If ExitIfTrue(isFalse(LeftModelName), "LeftModelName is empty..") Then Exit Function
            InterfaceFieldDeclaration = IIf(Relationship = "1:1", LeftModelName, LeftPluralizedModelName) & ": " & LeftModelName & IIf(Relationship = "1:1", "", "[]")
        Else
            Dim RightModelName: RightModelName = rs.fields("RightModelName"): If ExitIfTrue(isFalse(RightModelName), "RightModelName is empty..") Then Exit Function
            InterfaceFieldDeclaration = RightModelName & ": " & RightModelName
        End If
        
        fieldItems.Add InterfaceFieldDeclaration
        
        rs.MoveNext
    Loop
    
    GetModelInterfaceItems = "{" & fieldItems.JoinArr(";" & vbNewLine) & "}"
    
End Function

Public Function CopyClientModelInterface(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ModelName: ModelName = rs.fields("ModelName")
    If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    CopyClientModelInterface = "export interface " & ModelName & "Model " & GetModelInterfaceItems(SeqModelID)
        
    CopyToClipboard CopyClientModelInterface
    
End Function

Public Function CopyClientModelFormInterface(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    Dim modelInterfaceArr As New clsArray
    
    modelInterfaceArr.Add "export interface " & ModelName & "FormModel " & GetModelFormInterfaceItems(frm, SeqModelID, True)
    
    ''Get the related models of left models of this current model
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND RightModelID = " & SeqModelID & " AND Relationship <> ""M:M""" & _
        " AND NOT ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Dim relationships As New clsArray, relationships_2 As New clsArray, modelNames As New clsArray
       
    Do Until rs.EOF
        Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
        Dim IsSimpleRelationship: IsSimpleRelationship = rs.fields("IsSimpleRelationship")
        If IsSimpleRelationship Then
            relationships.Add LeftPluralizedModelName & ": number[]"
            modelNames.Add LeftPluralizedModelName
        End If
        
        relationships.Add "deleted" & LeftPluralizedModelName & ": number[]"
        
        rs.MoveNext
    Loop
    
    If relationships.count > 0 Then
        Dim ExtensionModel
        If modelNames.count > 0 Then
            modelNames.EscapeItems
            ExtensionModel = "Omit<" & ModelName & "FormModel, " & modelNames.JoinArr(" | ") & ">"
        Else
            ExtensionModel = ModelName & "FormModel"
        End If
        modelInterfaceArr.Add "export interface " & ModelName & "FormModelForSubmission extends " & _
        ExtensionModel & " {" & relationships.JoinArr(";" & vbNewLine) & "}"
    Else
        modelInterfaceArr.Add "export interface " & ModelName & "FormModelForSubmission extends " & ModelName & "FormModel {}"
    End If
    
    
    CopyClientModelFormInterface = modelInterfaceArr.JoinArr(vbNewLine & vbNewLine)
    
    CopyToClipboard CopyClientModelFormInterface
    
End Function

Public Function GetModelListFormInterface(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    Dim modelInterfaceArr As New clsArray
    
    modelInterfaceArr.Add "export interface " & ModelName & "FormModel " & GetModelFormInterfaceItems(frm, SeqModelID, True, True)
    modelInterfaceArr.Add "export interface " & ModelName & "FormModelForSubmission extends " & _
        "Record<" & Esc(PluralizedModelName) & "," & ModelName & "FormModel[]>{ deleted" & PluralizedModelName & ": number[]; }"
    
    GetModelListFormInterface = modelInterfaceArr.JoinArr(vbNewLine & vbNewLine)
    
    CopyToClipboard GetModelListFormInterface
    
End Function


Private Function GetModelFormInterfaceItems(frm As Object, SeqModelID, Optional MultipleRelationsOnly As Boolean = False, Optional AddCheckedAndTouched As Boolean = False) As String
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    
    Dim fieldItems As New clsArray
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID
    ''TABLE: qrySeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName|PluralizedFieldName
    ''AllowedOptions|DataType|DataTypeInterface
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        Dim AllowedOptions: AllowedOptions = rs.fields("AllowedOptions")
        Dim AllowNull: AllowNull = rs.fields("AllowNull")
        
        sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND LeftForeignKey = " & Esc(fieldName)
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        
        If DataType = "ENUM" Then
            Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
            Dim EnumOptionStr: EnumOptionStr = ConvertEnumToArray(DataTypeOption).JoinArr(" | ")
            ''If optional then add a null at the end of the items
            If AllowNull Then
                EnumOptionStr = EnumOptionStr & " | """""
            End If
            
            fieldItems.Add fieldName & ": " & EnumOptionStr
            
        ElseIf Not isFalse(AllowedOptions) Then
            Dim options As New clsArray: options.arr = AllowedOptions
            
            ''Force string conversions
            If DataTypeInterface <> "string" Then
                options.EscapeItems
            End If
            
            Dim optionsStr: optionsStr = options.JoinArr(" | ")
            If AllowNull Then
                optionsStr = optionsStr & " | """""
            End If
            
            fieldItems.Add fieldName & ": " & optionsStr
'        ElseIf Not rs2.EOF Then
'
'            Dim RightVariablePluralName: RightVariablePluralName = rs2.fields("RightVariablePluralName"): If ExitIfTrue(isFalse(RightVariablePluralName), "RightVariablePluralName is empty..") Then Exit Function
'            fieldItems.Add FieldName & ": " & RightVariablePluralName & "[0]"
            
        Else
            fieldItems.Add fieldName & ": string"
        End If
        
        rs.MoveNext
    Loop
    
    If AddCheckedAndTouched Then
        fieldItems.Add "checked: boolean"
        fieldItems.Add "touched: boolean"
    End If

    ''Any related model of this SeqModelID
    Dim RelationshipFilter
    If MultipleRelationsOnly Then
        RelationshipFilter = "RightModelID = " & SeqModelID
    Else
        RelationshipFilter = "RightModelID = " & SeqModelID & " OR LeftModelID = " & SeqModelID
    End If
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND (" & _
        RelationshipFilter & ") AND Relationship <> ""M:M"" AND NOT ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim InterfaceFieldDeclaration
        If rs.fields("RightModelID") = SeqModelID Then
            Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
            Dim LeftModelName: LeftModelName = rs.fields("LeftModelName"): If ExitIfTrue(isFalse(LeftModelName), "LeftModelName is empty..") Then Exit Function
            Dim IsSimpleRelationship: IsSimpleRelationship = rs.fields("IsSimpleRelationship")
            If IsSimpleRelationship Then
                InterfaceFieldDeclaration = LeftPluralizedModelName & ": BasicModel[]"
            Else
                Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
                InterfaceFieldDeclaration = LeftPluralizedModelName & ": " & DeclareModelBody(frm, LeftModelID) & "[]"
            End If
        Else
            Dim RightModelName: RightModelName = rs.fields("RightModelName"): If ExitIfTrue(isFalse(RightModelName), "RightModelName is empty..") Then Exit Function
            InterfaceFieldDeclaration = RightModelName & ": " & RightModelName
        End If
        
        fieldItems.Add InterfaceFieldDeclaration
        
        rs.MoveNext
    Loop
    
    GetModelFormInterfaceItems = "{" & fieldItems.JoinArr(";" & vbNewLine) & "}"
    
End Function


Public Function CopyClientModelQueryInterface(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " ORDER BY FilterOrder"
    
    Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray, urlQueries As New clsArray
    Dim uniqueQueries As New clsArray
    Do Until rs.EOF
        Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
        Dim FilterOperator: FilterOperator = rs.fields("FilterOperator"): If ExitIfTrue(isFalse(FilterOperator), "FilterOperator is empty..") Then Exit Function
        Dim DataType: DataType = rs.fields("DataType")
        Dim IsMultiple: IsMultiple = rs.fields("IsMultiple")
        
        If Not uniqueQueries.InArray(FilterQueryName) Then
            uniqueQueries.Add FilterQueryName
            
            If DataType = "DATEONLY" And FilterOperator = "Between" Then
                fields.Add "start_" & FilterQueryName & ": string"
                fields.Add "end_" & FilterQueryName & ": string"
                urlQueries.Add "start_" & FilterQueryName & ": string"
                urlQueries.Add "end_" & FilterQueryName & ": string"
            Else
                Dim FieldType
                If IsMultiple Then
                    FieldType = "BasicModel[]"
                ElseIf FilterOperator = "isPresent" Then
                    FieldType = "boolean"
                Else
                    FieldType = "string"
                End If
                
                fields.Add FilterQueryName & ": " & FieldType
                urlQueries.Add FilterQueryName & ": string"
            End If
            
        End If
        
        rs.MoveNext
    Loop
    
    fields.Add "[key: string]: unknown"
'    urlQueries.Add "page: string"
'    urlQueries.Add "limit: string"
'    urlQueries.Add "sort: string"
    
    Dim lines As New clsArray: lines.Add "export interface " & ModelName & "FilterFormDefaultValue {" & fields.JoinArr(";" & vbNewLine) & vbNewLine & "}"
    lines.Add "export interface " & ModelName & "URLQuery extends ListQuery {" & urlQueries.JoinArr(";" & vbNewLine) & vbNewLine & "}"
    
    CopyClientModelQueryInterface = lines.JoinArr(vbNewLine & vbNewLine)
    
    CopyToClipboard CopyClientModelQueryInterface
    
End Function

Public Function GenerateCompleteClientInterfaceFile(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim fileChunks As New clsArray
    
    fileChunks.Add "import { BasicModel, ListQuery } from ""./GeneralInterfaces"";"
    
    fileChunks.Add CopyClientModelInterface(frm, SeqModelID)
    fileChunks.Add CopyClientChildInterfaces(frm, SeqModelID)
    fileChunks.Add CopyClientModelFormInterface(frm, SeqModelID)
    fileChunks.Add CopyClientModelQueryInterface(frm, SeqModelID)
    
    GenerateCompleteClientInterfaceFile = fileChunks.JoinArr(vbNewLine & vbNewLine)
    CopyToClipboard GenerateCompleteClientInterfaceFile
    
    Dim filePath: filePath = ClientPath & "src\interfaces\" & ModelName & "Interfaces.ts"
    WriteToFile filePath, GenerateCompleteClientInterfaceFile, SeqModelID
    
End Function

'Public Function GenerateClientInterfaceFileForListForm(frm As Object, Optional SeqModelID = "")
'
'    RunCommandSaveRecord
'
'    If isFalse(SeqModelID) Then
'        SeqModelID = frm("SeqModelID")
'        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
'    End If
'
'    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
'    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
'
'    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
'    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
'
'    sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
'    Set rs = ReturnRecordset(sqlStr)
'    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
'
'    Dim fileChunks As New clsArray
'
'    fileChunks.Add "import { ListQuery } from ""./GeneralInterfaces"";"
'    fileChunks.Add CopyClientModelInterface(frm, SeqModelID)
'    fileChunks.Add CopyClientChildInterfaces(frm, SeqModelID)
'    CopyToClipboard Build_requiredList
'
'End Function

Public Function GenerateListHook(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VariablePluralName: VariablePluralName = rs.fields("VariablePluralName"): If ExitIfTrue(isFalse(VariablePluralName), "VariablePluralName is empty..") Then Exit Function
    
    Dim imports: imports = GetReplacedTemplate(rs, "List Hook Imports")
    
    Dim lines As New clsArray
    
    ''Additional import model assertions
    lines.Add imports
    lines.Add GetAdditionalImportBasedOnListVariableName(frm, SeqModelID)
    lines.Add "const use" & PluralizedModelName & " = (query: Partial<" & ModelName & "URLQuery>) => {"
    lines.Add "const router = useRouter();"
    lines.Add CopyAllEnumConstantDeclaration(frm, SeqModelID)
    ''lines.Add CopyUsePromiseAllOfThisModel(frm, SeqModelID)
    ''lines.Add Build_requiredList(frm, SeqModelID)
    lines.Add Generate_filterFormDefaultValue(frm, SeqModelID)
    lines.Add GenerateClientModelSortLimit(frm, SeqModelID)
    lines.Add GenerateMainListObject(frm, SeqModelID)
    lines.Add GenerateFilterFormToggleObject(frm, SeqModelID)
    lines.Add GenerateListBreadcrumbLinks(frm, SeqModelID)
    
    lines.Add "return { requiredListObject, sortAndLimitObject, mainListObject, formikObject, toggleObject, breadCrumbLinks};};"
    lines.Add "export type TMainListObject = {" & VariablePluralName & ": " & ModelName & "Model[]; gridLoading: boolean; recordCount: number;};"
    lines.Add "export type TFormikFilterFormObject = { filterFormDefaultValue: " & ModelName & "FilterFormDefaultValue; filterFormInitialValue: " & ModelName & _
        "FilterFormDefaultValue; handleFilterFormSubmit: (values: " & ModelName & "FilterFormDefaultValue) => void; handleFilterFormReset: (formik: FormikProps<any>) => void; };"
    lines.Add "export type T" & PluralizedModelName & "Hook = ReturnType<typeof use" & PluralizedModelName & ">;"
    lines.Add GenerateTRequiredListDeclaration(frm, SeqModelID)
    lines.Add "export default use" & PluralizedModelName & ";"
    
    GenerateListHook = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateListHook
    
    sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\hooks\" & VariableName & "\use" & PluralizedModelName & ".ts"
    WriteToFile filePath, GenerateListHook, SeqModelID
    
End Function

Public Function CopyAllEnumConstantDeclaration(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''Constants
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND (DataType = ""ENUM"" OR Not isNull(AllowedOptions))"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ConstantsArr As New clsArray
    ConstantsArr.Add "//Generated by CopyAllEnumConstantDeclaration"
    
    Do Until rs.EOF
        ConstantsArr.Add CopyEnumConstantDeclaration(frm, rs.fields("SeqModelFieldID"))
        rs.MoveNext
    Loop
    
    ''The months depending on the filter
    Dim declareMonth: declareMonth = isPresent("qrySeqModelFilters", "SeqModelID = " & SeqModelID & " AND ListVariableName = ""months""")
    If (declareMonth) Then
        ConstantsArr.Add "const months = monthsModel"
    End If
    
    ''Check on the related model if there's a field that require constants
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE RightModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & LeftModelID & " AND (DataType = ""ENUM"" OR Not isNull(AllowedOptions))"
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        Do Until rs2.EOF
            ConstantsArr.Add CopyEnumConstantDeclaration(frm, rs2.fields("SeqModelFieldID"))
            rs2.MoveNext
        Loop
        rs.MoveNext
    Loop
    
    Dim Constants: Constants = ConstantsArr.JoinArr(vbNewLine)
    
    CopyAllEnumConstantDeclaration = Constants
    CopyToClipboard CopyAllEnumConstantDeclaration
    
End Function

Public Function Generate_filterFormDefaultValue(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function

    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    ''TABLE: qrySeqModelFilters Fields: SeqModelFilterID|SeqModelID|SeqModelFieldID|IsMultiple|FilterQueryName
    ''FilterOperator|Timestamp|CreatedBy|RecordImportID|ListVariableName|ControlType|DatabaseFieldName|FieldName
    ''SeqDataTypeID|DataType
    
    ''Available control type -> Option, Text, Checkbox, Switch
    Dim uniqueFilterNames As New clsArray, fields As New clsArray, supplyMissingNames As New clsArray
    Do Until rs.EOF
        Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
        Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
        Dim FilterOperator: FilterOperator = rs.fields("FilterOperator"): If ExitIfTrue(isFalse(FilterOperator), "FilterOperator is empty..") Then Exit Function
        Dim DataType: DataType = rs.fields("DataType")
        Dim IsMultiple: IsMultiple = rs.fields("IsMultiple")
        If Not uniqueFilterNames.InArray(FilterQueryName) Then ''If not present
            uniqueFilterNames.Add FilterQueryName, True
            
            If ControlType = "Option" Then
                fields.Add FilterQueryName & ": ""all"""
            ElseIf ControlType = "Checkbox" Then
                fields.Add FilterQueryName & ": []"
            ElseIf ControlType = "Switch" Then
                fields.Add FilterQueryName & ": false"
            ElseIf FilterOperator = "Between" And DataType = "DATEONLY" Then
                fields.Add "start_" & FilterQueryName & " : """""
                fields.Add "end_" & FilterQueryName & " : """""
            ElseIf ControlType = "Autocomplete" And IsMultiple Then
                fields.Add FilterQueryName & ": []"
            Else
                fields.Add FilterQueryName & ": """""
            End If
            
            If IsMultiple Then
                Dim ListVariableName: ListVariableName = rs.fields("ListVariableName"): If ExitIfTrue(isFalse(ListVariableName), "ListVariableName is empty..") Then Exit Function
                supplyMissingNames.Add "if (" & ListVariableName & ") {supplyMissingNames(" & ListVariableName & " as BasicModel[], filterFormInitialValue." & FilterQueryName & ");}"
            End If
            
        End If
        rs.MoveNext
    Loop
    
    Dim lines As New clsArray
    lines.Add "//Generated by Generate_filterFormDefaultValue"
    lines.Add "const filterFormDefaultValue: " & ModelName & "FilterFormDefaultValue = { " & fields.JoinArr(",") & "}"
    lines.Add "//Reshape the initial values of the filter form depending on the URL queries"
    lines.Add "const filterFormInitialValue: " & ModelName & "FilterFormDefaultValue = getFilterValueFromURL(query,filterFormDefaultValue);"
    lines.Add "//This will supply the name key for each array of objects"
    
    If supplyMissingNames.count > 0 Then
        lines.Add supplyMissingNames.JoinArr(vbNewLine)
    End If
    
    lines.Add GenerateFilterFormikEvents(frm, SeqModelID)
    lines.Add "const formikObject: TFormikFilterFormObject = { filterFormDefaultValue, filterFormInitialValue, handleFilterFormSubmit, handleFilterFormReset}"
    
    Generate_filterFormDefaultValue = lines.JoinArr(vbNewLine)
    CopyToClipboard Generate_filterFormDefaultValue
    
End Function

Private Function SplitSortString(str) As String
    'Check if the string contains the word "DESC"
    If InStr(str, "DESC") > 0 Then
        'Remove the brackets and the word "DESC" from the string
        str = replace(str, "[", "")
        str = replace(str, "]", "")
        str = replace(str, "DESC", "")
        'Split the remaining string into two parts
        Dim arr() As String
        arr = Split(str, " ")
        'Format the output string
        SplitSortString = "[" & arr(0) & """," & Esc("desc") & "]"
    Else
        'Remove the brackets from the string
        str = replace(str, "[", "")
        str = replace(str, "]", "")
        'Format the output string
        SplitSortString = "[" & str & ",""asc""]"
    End If
End Function

Public Function GenerateClientSortOptions(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SortString: SortString = rs.fields("SortString"): If ExitIfTrue(isFalse(SortString), "SortString is empty..") Then Exit Function
    
    Dim lines As New clsArray
    
    lines.Add "//Generated by GenerateClientSortOptions"
    lines.Add "const sortedBy = getSortedBy(query, " & Esc(SortString) & ");"
    
    sqlStr = "SELECT * FROM qrySeqModelSorts WHERE SeqModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    
    Do Until rs.EOF
        Dim ModelFieldCaption: ModelFieldCaption = rs.fields("ModelFieldCaption"): If ExitIfTrue(isFalse(ModelFieldCaption), "ModelFieldCaption is empty..") Then Exit Function
        Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName")
        Dim Asc: Asc = rs.fields("Asc")
        Dim Desc: Desc = rs.fields("Desc")
        Dim SortKey: SortKey = rs.fields("SortKey")
        
        Dim realKey: realKey = IIf(Not isFalse(SortKey), SortKey, DatabaseFieldName)
        Dim realAsc: realAsc = IIf(Not isFalse(Asc), Asc, DatabaseFieldName)
        Dim realDesc: realDesc = IIf(Not isFalse(Desc), Desc, "-" & DatabaseFieldName)
        
        fields.Add realKey & ": { caption: " & Esc(ModelFieldCaption) & ",asc: " & Esc(realAsc) & ", desc: " & Esc(realDesc) & "}"
        rs.MoveNext
    Loop
    
    lines.Add "const sortOptions: SortOptionsAsString = { sortedBy, sortObject: { " & fields.JoinArr & " } };"
    
    GenerateClientSortOptions = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateClientSortOptions
    
End Function

Public Function GenerateClientModelSortLimit(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    Dim lines As New clsArray
    Dim ClientSortOptions: ClientSortOptions = GenerateClientSortOptions(frm, SeqModelID)
    
    lines.Add "//Generated by GenerateClientModelSortLimit"
    lines.Add ClientSortOptions
    lines.Add "const handle" & ModelName & "Sort = (name: string) => modifySortAsString(name, sortOptions, router, query);"
    lines.Add "const handle" & ModelName & "LimitChange = (event: SelectChangeEvent<string>) => modifyLimit(event.target.value, router, query);"
    lines.Add "const sortAndLimitObject = { limit: getFirstItem(query.limit, ""20""),sortOptions, handle" & ModelName & "Sort, handle" & ModelName & "LimitChange };"
    GenerateClientModelSortLimit = lines.JoinArr(vbNewLine & vbNewLine)
    CopyToClipboard GenerateClientModelSortLimit

End Function

Public Function GenerateMainListObject(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim VariablePluralName: VariablePluralName = rs.fields("VariablePluralName"): If ExitIfTrue(isFalse(VariablePluralName), "VariablePluralName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Main List Object")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[VariablePluralName]", VariablePluralName)
    replacedContent = replace(replacedContent, "[PluralizedModelName]", PluralizedModelName)
    replacedContent = replace(replacedContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateMainListObject"
    lines.Add replacedContent
    GenerateMainListObject = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateMainListObject
    
End Function


Public Function GenerateFilterFormikEvents(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Formik events")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateFilterFormikEvents"
    lines.Add replacedContent
    
    GenerateFilterFormikEvents = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateFilterFormikEvents
    
End Function

Public Function GenerateFilterFormToggleObject(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateFilterFormToggleObject"
    lines.Add "const { toggle: isFilterFormShown, handleToggle: toggleFilterForm } = useToggle();"
    lines.Add "const toggleObject = { isFilterFormShown, toggleFilterForm };"
    GenerateFilterFormToggleObject = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateFilterFormToggleObject
    
End Function

Public Function GenerateListBreadcrumbLinks(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim VerboseModelName: VerboseModelName = rs.fields("VerboseModelName"): If ExitIfTrue(isFalse(VerboseModelName), "VerboseModelName is empty..") Then Exit Function
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateListBreadcrumbLinks"
    lines.Add "const breadCrumbLinks = [{href: " & Esc("/" & ModelPath) & ", caption: " & Esc(VerboseModelName & " List") & "}];"
    
    GenerateListBreadcrumbLinks = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateListBreadcrumbLinks
    
End Function

Public Function CreateModelListBodyComponent(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("ModelListBody component")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[PluralizedModelName]", PluralizedModelName)
    replacedContent = replace(replacedContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    
    CreateModelListBodyComponent = replacedContent
    CopyToClipboard CreateModelListBodyComponent
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "ListBody.tsx"
    WriteToFile filePath, CreateModelListBodyComponent, SeqModelID
    
End Function

Public Function CreateModelFilterComponent(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim PluralizedVerboseModelName: PluralizedVerboseModelName = rs.fields("PluralizedVerboseModelName"): If ExitIfTrue(isFalse(PluralizedVerboseModelName), "PluralizedVerboseModelName is empty..") Then Exit Function
    Dim FilterControls As New clsArray, uniqueFields As New clsArray
    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " ORDER BY FilterOrder ASC"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        If Not uniqueFields.InArray(FilterQueryName) Then
            uniqueFields.Add FilterQueryName
            FilterControls.Add GenerateIndividualFilterControl(frm, SeqModelFilterID)
        End If
        rs.MoveNext
    Loop
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("ModelFilter Component")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[PluralizedModelName]", PluralizedModelName)
    replacedContent = replace(replacedContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[PluralizedVerboseModelName]", PluralizedVerboseModelName)
    replacedContent = replace(replacedContent, "[FilterControls]", FilterControls.JoinArr(vbNewLine))
    
    Dim lines As New clsArray
    lines.Add "//Generated by CreateModelFilterComponent"
    lines.Add replacedContent
    
    CreateModelFilterComponent = lines.JoinArr(vbNewLine)
    CopyToClipboard CreateModelFilterComponent
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "Filter.tsx"
    WriteToFile filePath, CreateModelFilterComponent, SeqModelID
    
End Function

Public Function CreateModelListHeaderComponent(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("ModelListHeader component")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    
    CreateModelListHeaderComponent = replacedContent
    
    Dim lines As New clsArray
    lines.Add "//Generated by CreateModelListHeaderComponent"
    lines.Add CreateModelListHeaderComponent
    CreateModelListHeaderComponent = lines.NewLineJoin
    
    CopyToClipboard CreateModelListHeaderComponent
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "ListHeader.tsx"
    WriteToFile filePath, CreateModelListHeaderComponent, SeqModelID
    
End Function

Public Function CreateModelGridComponent(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VariablePluralName: VariablePluralName = rs.fields("VariablePluralName"): If ExitIfTrue(isFalse(VariablePluralName), "VariablePluralName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("ModelGrid component")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[VariablePluralName]", VariablePluralName)
    replacedContent = replace(replacedContent, "[PluralizedModelName]", PluralizedModelName)
    
    CreateModelGridComponent = replacedContent
    CopyToClipboard CreateModelGridComponent
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "Grid.tsx"
    WriteToFile filePath, CreateModelGridComponent, SeqModelID
    
End Function

Public Function CreateModelGridComponentForListForm(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VariablePluralName: VariablePluralName = rs.fields("VariablePluralName"): If ExitIfTrue(isFalse(VariablePluralName), "VariablePluralName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    CreateModelGridComponentForListForm = GetReplacedTemplate(rs, "ModelGrid For List Form")
    CopyToClipboard CreateModelGridComponentForListForm
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "Grid.tsx"
    WriteToFile filePath, CreateModelGridComponentForListForm, SeqModelID
    
End Function

Public Function CreateSingleModelComponent(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    ''TABLE: tblSeqModels Fields: SeqModelID|ModelName|ExportAs|TableName|Timestamps|Timestamp|CreatedBy|RecordImportID
    ''BackendProjectID|ModelFileName|ControllerFileName|RouteFileName|PluralizedModelName|ModelPath|VariableName
    ''IsMainQuery|LeftKey|RightKey|SortString|SlugField|InterfaceFileName|VariablePluralName

    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = IIf(IsNull(SlugField), "id", "slug")
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("SingleModel component")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[SlugOrID]", SlugOrId)
    replacedContent = replace(replacedContent, "[SlugField]", IIf(isFalse(SlugField), "id", SlugField))
    
    CreateSingleModelComponent = replacedContent
    CopyToClipboard CreateSingleModelComponent
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\Single" & ModelName & ".tsx"
    WriteToFile filePath, CreateSingleModelComponent, SeqModelID
    
End Function

Public Function Create_fetchModel(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = IIf(IsNull(SlugField), "id", "slug")
    
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("fetchModel snippet")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[SlugOrID]", SlugOrId)
    
    Dim lines As New clsArray
    lines.Add "//Generated by Create_fetchModel"
    lines.Add replacedContent
    Create_fetchModel = lines.JoinArr(vbNewLine)
    CopyToClipboard Create_fetchModel
    
End Function

Public Function Create_formInitialValues(frm As Object, Optional SeqModelID = "")
    
    ''Run this for the fields
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = IIf(IsNull(SlugField), "id", "slug")
    
    ''Get the field array
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        fields.Add CreateFieldFormInitialValues(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND " & _
        " RightModelID = " & SeqModelID & " AND Relationship <> ""M:M"" AND NOT ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim fixes As New clsArray, touchConditions As New clsArray, LeftPluralizedModelName
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
        Dim IsSimpleRelationship: IsSimpleRelationship = rs.fields("IsSimpleRelationship")
        fields.Add LeftPluralizedModelName & ": []"
        If IsSimpleRelationship Then
            fixes.Add "formInitialValues." & LeftPluralizedModelName & " = " & VariableName & "." & LeftPluralizedModelName & ".map((item) => ({id: """", name: """"}));"
        Else
            fixes.Add "formInitialValues." & LeftPluralizedModelName & " = " & VariableName & "." & LeftPluralizedModelName & ".map((item) => (" & GetFromRecordToFormInitAssignment(frm, LeftModelID) & "));"
            
        End If
        rs.MoveNext
    Loop
    
    Dim lines As New clsArray
    lines.Add "//Generated by Create_formInitialValues"
    lines.Add "const formInitialValues: " & ModelName & "FormModel = {" & fields.JoinArr(", //Generated by CreateFieldFormInitialValues" & vbNewLine) & "};"
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Fix Form Initial Value")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[VariableName]", VariableName)
    
    Dim FixRelated: FixRelated = "//Edit the id and name to match the id and name of the list" & vbNewLine & fixes.JoinArr(vbNewLine & vbNewLine)
    replacedContent = replace(replacedContent, "[FixRelated]", IIf(fixes.count > 0, FixRelated, ""))
    lines.Add replacedContent
    
    ''Add at least one row for the non simple relationships
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND " & _
        " RightModelID = " & SeqModelID & " AND Relationship <> ""M:M"" AND NOT ExcludeInForm AND NOT IsSimpleRelationship"
        
    Set rs = ReturnRecordset(sqlStr)
    Dim deletedStates As New clsArray
    Do Until rs.EOF
        LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add PushAtLeastOneRow(frm, SeqModelRelationshipID)
        deletedStates.Add "const [deleted" & LeftPluralizedModelName & ", setDeleted" & LeftPluralizedModelName & "] = useState<number[]>([]);"
        rs.MoveNext
    Loop
    
    If deletedStates.count > 0 Then
        Dim item
        For Each item In deletedStates.arr
            lines.Add item
        Next item
    End If
    
    Create_formInitialValues = lines.JoinArr(vbNewLine)
    CopyToClipboard Create_formInitialValues
    
End Function

Public Function GetFromRecordToFormInitAssignment(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " ORDER BY FieldOrder"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''id: item.id
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim PrimaryKey: PrimaryKey = rs.fields("PrimaryKey")
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        
        Dim ValueString: ValueString = fieldName
        
        If PrimaryKey Or DataTypeInterface = "number" Then ValueString = ValueString & ".toString()"
            
        fields.Add fieldName & ": item." & ValueString
        rs.MoveNext
    Loop
    
    fields.Add "checked : false"
    fields.Add "touched : false"
    
    Dim lines As New clsArray
    lines.Add "//Generated by GetFromRecordToFormInitAssignment"
    lines.Add "{" & fields.JoinArr("," & vbNewLine) & "}"
    
    GetFromRecordToFormInitAssignment = lines.JoinArr(vbNewLine)
    
    CopyToClipboard GetFromRecordToFormInitAssignment
    
End Function

Public Function GetInitialValues_forListForm(frm As Object, Optional SeqModelID = "")
    
    ''Run this for the fields
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        
    Dim VariablePluralName: VariablePluralName = rs.fields("VariablePluralName"): If ExitIfTrue(isFalse(VariablePluralName), "VariablePluralName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    
    ''Fields, FieldsForPush
    Dim fields As New clsArray, FieldsForPush As New clsArray
    Do Until rs.EOF
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        
        Dim fieldItem: fieldItem = fieldName & ": item." & fieldName
        If DataTypeInterface = "number" Then fieldItem = fieldItem & ".toString()"
        fields.Add fieldItem
        
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption")
        Dim AllowNull: AllowNull = rs.fields("AllowNull")
        Dim AllowedOptions: AllowedOptions = rs.fields("AllowedOptions")
        Dim options As New clsArray
        
        sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND LeftForeignKey = " & Esc(fieldName)
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        
        If DataType = "ENUM" Then
            If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
            Set options = ConvertEnumToArray(DataTypeOption)
            FieldsForPush.Add fieldName & ": " & IIf(AllowNull, """""", options.arr(0))
        ElseIf DataType = "DECIMAL" Then
            FieldsForPush.Add fieldName & ": " & IIf(AllowNull, """""", """0.00""")
        ElseIf Not IsNull(AllowedOptions) Then
            options.arr = AllowedOptions
            FieldsForPush.Add fieldName & ": " & Esc(options.arr(0))
        ElseIf Not rs2.EOF Then
            Dim RightVariablePluralName: RightVariablePluralName = rs2.fields("RightVariablePluralName"): If ExitIfTrue(isFalse(RightVariablePluralName), "RightVariablePluralName is empty..") Then Exit Function
            FieldsForPush.Add fieldName & ": " & RightVariablePluralName & " && " & RightVariablePluralName & ".length > 0 ? " & RightVariablePluralName & "[0].id.toString() : """""
        Else
            FieldsForPush.Add fieldName & ": """""
        End If
            
        rs.MoveNext
    Loop
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Initial Values for List Form")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[VariablePluralName]", VariablePluralName)
    replacedContent = replace(replacedContent, "[PluralizedModelName]", PluralizedModelName)
    replacedContent = replace(replacedContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[Fields]", fields.JoinArr("," & vbNewLine))
    replacedContent = replace(replacedContent, "[FieldsForPush]", FieldsForPush.JoinArr("," & vbNewLine))
    
    GetInitialValues_forListForm = replacedContent
    CopyToClipboard GetInitialValues_forListForm
    
End Function


Public Function CreateModelForm_breadcrumbLinks(frm As Object, Optional SeqModelID = "")
    
    ''Run this for the fields
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VerboseModelName: VerboseModelName = rs.fields("VerboseModelName"): If ExitIfTrue(isFalse(VerboseModelName), "VerboseModelName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = IIf(IsNull(SlugField), "id", "slug")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Form breadcrumbLinks")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    replacedContent = replace(replacedContent, "[VerboseModelName]", VerboseModelName)
    replacedContent = replace(replacedContent, "[SlugOrID]", SlugOrId)
    replacedContent = replace(replacedContent, "[SlugField]", IIf(IsNull(SlugField), "id", SlugField))
    
    Dim lines As New clsArray
    lines.Add "//Generated by CreateModelForm_breadcrumbLinks"
    lines.Add replacedContent
    
    CreateModelForm_breadcrumbLinks = lines.JoinArr(vbNewLine)
    CopyToClipboard CreateModelForm_breadcrumbLinks
    
End Function

Public Function Create_handleSubmit(frm As Object, Optional SeqModelID = "")
    
    ''Run this for the fields
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VerboseModelName: VerboseModelName = rs.fields("VerboseModelName"): If ExitIfTrue(isFalse(VerboseModelName), "VerboseModelName is empty..") Then Exit Function
    
    ''ConversionToIDs
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND RightModelID = " & SeqModelID & " AND Relationship <> ""M:M""" & _
        " AND NOT ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Dim lines As New clsArray
    lines.Add "//Generated by Create_handleSubmit"
    Dim TemplateContent, replacedContent
    Dim relationshipUpdates As New clsArray, touchConditions As New clsArray
    Do Until rs.EOF
        Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
        Dim IsSimpleRelationship: IsSimpleRelationship = rs.fields("IsSimpleRelationship")
        If IsSimpleRelationship Then
            TemplateContent = GetTemplateContent("Modify Relationship for Update")
            lines.Add "newValues[" & Esc(LeftPluralizedModelName) & "] = values." & LeftPluralizedModelName & ".map((item) => parseInt(item.id as string));"
        Else
            touchConditions.Add "newValues." & LeftPluralizedModelName & " = newValues." & LeftPluralizedModelName & ".filter((item) => item.touched);"
            TemplateContent = GetTemplateContent("Modify Relationship for Update Non-simple")
        End If
        replacedContent = replace(TemplateContent, "[LeftPluralizedModelName]", LeftPluralizedModelName)
        replacedContent = replace(replacedContent, "[VariableName]", VariableName)
        relationshipUpdates.Add replacedContent
        rs.MoveNext
    Loop
    Dim ConversionToIDs: ConversionToIDs = lines.JoinArr(vbNewLine)
    Dim ModifyList: ModifyList = relationshipUpdates.JoinArr(vbNewLine)
    
    TemplateContent = GetTemplateContent("Form Handle Submit")
    replacedContent = replace(TemplateContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    replacedContent = replace(replacedContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[ConversionToIDs]", ConversionToIDs)
    replacedContent = replace(replacedContent, "[VerboseModelName]", VerboseModelName)
    replacedContent = replace(replacedContent, "[ModifyList]", ModifyList)
    replacedContent = replace(replacedContent, "[TouchConditions]", touchConditions.JoinArr(vbNewLine))
    
    lines.Add replacedContent
    
    Create_handleSubmit = lines.JoinArr(vbNewLine)
    CopyToClipboard Create_handleSubmit
    
End Function

Public Function Get_handleSubmitForListForm(frm As Object, Optional SeqModelID = "")
    
    ''Run this for the fields
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VerboseModelName: VerboseModelName = rs.fields("VerboseModelName"): If ExitIfTrue(isFalse(VerboseModelName), "VerboseModelName is empty..") Then Exit Function
    Dim PluralizedVerboseModelName: PluralizedVerboseModelName = rs.fields("PluralizedVerboseModelName"): If ExitIfTrue(isFalse(PluralizedVerboseModelName), "PluralizedVerboseModelName is empty..") Then Exit Function
    Dim VariablePluralName: VariablePluralName = rs.fields("VariablePluralName"): If ExitIfTrue(isFalse(VariablePluralName), "VariablePluralName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    Dim TemplateContent, replacedContent
    TemplateContent = GetTemplateContent("handleSubmit for List Form")
    replacedContent = replace(TemplateContent, "[PluralizedVerboseModelName]", PluralizedVerboseModelName)
    replacedContent = replace(replacedContent, "[VariablePluralName]", VariablePluralName)
    replacedContent = replace(replacedContent, "[PluralizedModelName]", PluralizedModelName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    replacedContent = replace(replacedContent, "[ModelName]", ModelName)
    
    Get_handleSubmitForListForm = replacedContent
    CopyToClipboard Get_handleSubmitForListForm
    
End Function

Public Function Create_handleDelete(frm As Object, Optional SeqModelID = "")
    
    ''Run this for the fields
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VerboseModelName: VerboseModelName = rs.fields("VerboseModelName"): If ExitIfTrue(isFalse(VerboseModelName), "VerboseModelName is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("handleDelete")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    replacedContent = replace(replacedContent, "[VerboseModelName]", VerboseModelName)
    
    Dim lines As New clsArray
    lines.Add "//Generated by Create_handleDelete"
    lines.Add replacedContent
    
    Create_handleDelete = lines.JoinArr(vbNewLine)
    CopyToClipboard Create_handleDelete
    
End Function

Public Function Create_useModelFormHook(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VariablePluralName: VariablePluralName = rs.fields("VariablePluralName"): If ExitIfTrue(isFalse(VariablePluralName), "VariablePluralName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = IIf(IsNull(SlugField), "id", "slug")
    
    Dim lines As New clsArray
    lines.Add CopyAllEnumConstantDeclaration(frm, SeqModelID)
    ''lines.Add CopyUsePromiseAllOfThisModel(frm, SeqModelID)
    ''lines.Add Build_requiredList(frm, SeqModelID)
    Dim RequiredLists: RequiredLists = lines.JoinArr(vbNewLine)
    Dim FetchModel: FetchModel = Create_fetchModel(frm, SeqModelID)
    Dim InitialValues: InitialValues = Create_formInitialValues(frm, SeqModelID)
    Dim breadCrumbLinks: breadCrumbLinks = CreateModelForm_breadcrumbLinks(frm, SeqModelID)
    Dim HandleSubmit: HandleSubmit = Create_handleSubmit(frm, SeqModelID)
    Dim HandleDelete: HandleDelete = Create_handleDelete(frm, SeqModelID)
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("useModelForm hook")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[RequiredLists]", RequiredLists)
    replacedContent = replace(replacedContent, "[FetchModel]", FetchModel)
    replacedContent = replace(replacedContent, "[InitialValues]", InitialValues)
    replacedContent = replace(replacedContent, "[BreadcrumbLinks]", breadCrumbLinks)
    replacedContent = replace(replacedContent, "[HandleSubmit]", HandleSubmit)
    replacedContent = replace(replacedContent, "[HandleDelete]", HandleDelete)
    replacedContent = replace(replacedContent, "[SlugOrID]", SlugOrId)
    replacedContent = replace(replacedContent, "[TRequiredList]", GenerateTRequiredListDeclaration(frm, SeqModelID))
    replacedContent = replace(replacedContent, "[setDeleteds]", GenerateSetDeleteds(frm, SeqModelID))
    replacedContent = replace(replacedContent, "[RelatedModelImports]", CreateRelatedModelFormHookImports(frm, SeqModelID))
    
    
    Create_useModelFormHook = "//Generated by Create_useModelFormHook" & vbNewLine & replacedContent
    CopyToClipboard Create_useModelFormHook
    
    sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\hooks\" & VariableName & "\use" & ModelName & "Form.ts"
    WriteToFile filePath, Create_useModelFormHook, SeqModelID
    
End Function

Public Function GenerateSetDeleteds(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE RightModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim items As New clsArray
    Do Until rs.EOF
        Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
        items.Add "setDeleted" & LeftPluralizedModelName
        rs.MoveNext
    Loop
    
    If items.count > 0 Then
        GenerateSetDeleteds = "," & items.JoinArr(",")
    End If
    
    CopyToClipboard GenerateSetDeleteds
    
End Function

Public Function Get_useModelFormHook_forListForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VariablePluralName: VariablePluralName = rs.fields("VariablePluralName"): If ExitIfTrue(isFalse(VariablePluralName), "VariablePluralName is empty..") Then Exit Function

    
    Dim lines As New clsArray
    lines.Add CopyAllEnumConstantDeclaration(frm, SeqModelID)
    ''lines.Add CopyUsePromiseAllOfThisModel(frm, SeqModelID)
    ''lines.Add Build_requiredList(frm, SeqModelID)
    Dim RequiredLists: RequiredLists = lines.JoinArr(vbNewLine)
    Dim InitialValues: InitialValues = GetInitialValues_forListForm(frm, SeqModelID)
    Dim HandleSubmit: HandleSubmit = Get_handleSubmitForListForm(frm, SeqModelID)
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("useModelForm hook For List Form")
    
    Dim replacedContent
    replacedContent = GetReplacedTemplate(rs, "useModelForm hook For List Form")
    replacedContent = replace(replacedContent, "[RequiredLists]", RequiredLists)
    replacedContent = replace(replacedContent, "[InitialValues]", InitialValues)
    replacedContent = replace(replacedContent, "[HandleSubmit]", HandleSubmit)
    replacedContent = replace(replacedContent, "[TRequiredList]", GenerateTRequiredListDeclaration(frm, SeqModelID))
    
    Get_useModelFormHook_forListForm = replacedContent
    CopyToClipboard Get_useModelFormHook_forListForm
    
    sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\hooks\" & VariableName & "\use" & ModelName & "Form.ts"
    WriteToFile filePath, Get_useModelFormHook_forListForm, SeqModelID
    
End Function

Public Function CreateModelSchema(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    
    ''Generate the Fields variable -> Run CreateFieldSchema on each SeqModelFields
    Dim items As New clsArray
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        items.Add CreateFieldSchema(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    ''The Right model is SeqModelID
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND " & _
        " RightModelID = " & SeqModelID & " AND Relationship <> ""M:M"" AND NOT ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)

    Do Until rs.EOF
        Dim IsSimpleRelationship: IsSimpleRelationship = rs.fields("IsSimpleRelationship")
        Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        Dim LeftForeignKey: LeftForeignKey = rs.fields("LeftForeignKey"): If ExitIfTrue(isFalse(LeftForeignKey), "LeftForeignKey is empty..") Then Exit Function
        If IsSimpleRelationship Then
            items.Add LeftPluralizedModelName & ": Yup.array()"
        Else
            items.Add CreateModelSchemaArrayValidation(frm, LeftModelID, LeftForeignKey)
        End If
        rs.MoveNext
    Loop
    
    Dim fields: fields = items.JoinArr("," & vbNewLine)

    Dim TemplateContent: TemplateContent = GetTemplateContent("ModelSchema")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[Fields]", fields)
    
    CreateModelSchema = replacedContent
    CreateModelSchema = GetGeneratedByFunctionSnippet(CreateModelSchema, "CreateModelSchema")
    CopyToClipboard CreateModelSchema
    
    sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\schema\" & ModelName & "Schema.ts"
    WriteToFile filePath, CreateModelSchema, SeqModelID
    
End Function

Public Function CreateModelDetailPage(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim VerboseModelName: VerboseModelName = rs.fields("VerboseModelName"): If ExitIfTrue(isFalse(VerboseModelName), "VerboseModelName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = IIf(IsNull(SlugField), "id", "slug")
    Dim SidebarEnabled: SidebarEnabled = isPresent("tblBackendProjects", "BackendProjectID = " & BackendProjectID & " AND SidebarEnabled")
    
    Dim TemplateContent: TemplateContent = GetTemplateContent(IIf(SidebarEnabled, "Detail Page with Sidebar", "[id].tsx"))
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[ModelPath]", ModelPath)
    replacedContent = replace(replacedContent, "[SlugOrID]", SlugOrId)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[VerboseModelName]", VerboseModelName)
    
    CreateModelDetailPage = replacedContent
    CopyToClipboard CreateModelDetailPage
    
    sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\pages\" & ModelPath & "\[" & SlugOrId & "].tsx"
    WriteToFile filePath, CreateModelDetailPage, SeqModelID
    
End Function

Public Function CreateModelFormBody(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("ModelFormBody")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    
    CreateModelFormBody = replacedContent
    CopyToClipboard CreateModelFormBody
    
    sqlStr = "SELECT * FROM tblBackendProjects WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "FormBody.tsx"
    WriteToFile filePath, CreateModelFormBody, SeqModelID
    
End Function

Public Function GenerateModelFormControls(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    
    ''GenerateIndividualFormControl, GenerateFormControlFromRelationship
    Dim controls As New clsArray
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey ORDER BY FieldOrder ASC"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        controls.Add GenerateIndividualFormControl(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & " AND " & _
        " RightModelID = " & SeqModelID & " AND Relationship <> ""M:M"" AND NOT ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim IsSimpleRelationship: IsSimpleRelationship = rs.fields("IsSimpleRelationship")
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        If IsSimpleRelationship Then
            controls.Add GenerateFormControlFromRelationship(frm, SeqModelRelationshipID)
        Else
            controls.Add GenerateFieldArrayComponent(frm, SeqModelRelationshipID)
            GenerateFieldArrayFormFromRelationship frm, SeqModelRelationshipID
        End If
        rs.MoveNext
    Loop
    
    GenerateModelFormControls = controls.JoinArr(vbNewLine)
    GenerateModelFormControls = GetGeneratedByFunctionSnippet(GenerateModelFormControls, "GenerateModelFormControls", True)
    
    CopyToClipboard GenerateModelFormControls
    
End Function

Public Function CreateModelForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    
    Dim FormControls: FormControls = GenerateModelFormControls(frm, SeqModelID)
    
    Dim item, LeftModelNames As New clsArray, LeftPluralizedModelNames As New clsArray
    LeftModelNames.arr = Elookups("qrySeqModelRelationships", "RightModelID = " & SeqModelID, "LeftModelName")
    LeftPluralizedModelNames.arr = Elookups("qrySeqModelRelationships", "RightModelID = " & SeqModelID, "LeftPluralizedModelName")
    
    Dim RelatedImports As New clsArray
    For Each item In LeftModelNames.arr
        RelatedImports.Add "import " & item & "Form from ""./" & item & "Form"";"
    Next item
    
    
    Dim setDeleteds As New clsArray
    For Each item In LeftPluralizedModelNames.arr
        setDeleteds.Add "setDeleted" & item & ": Dispatch<SetStateAction<number[]>>;"
    Next item
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("ModelForm")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[ModelName]", ModelName)
    replacedContent = replace(replacedContent, "[VariableName]", VariableName)
    replacedContent = replace(replacedContent, "[FormControls]", FormControls)
    replacedContent = replace(replacedContent, "[RelatedImports]", RelatedImports.JoinArr(vbNewLine))
    replacedContent = replace(replacedContent, "[setDeleteds]", setDeleteds.JoinArr(";" & vbNewLine))
    
    CreateModelForm = replacedContent
    
    CreateModelForm = GetGeneratedByFunctionSnippet(CreateModelForm, "CreateModelForm")
    
    CopyToClipboard CreateModelForm
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "Form.tsx"
    
    WriteToFile filePath, CreateModelForm, SeqModelID
    
    
End Function

Public Function CreateModelFormForListForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    
    CreateModelFormForListForm = GetReplacedTemplate(rs, "ModelForm for List Form")
    
    Dim handleAddRowButtonClick: handleAddRowButtonClick = Generate_handleAddRowButtonClickForListForm(frm, SeqModelID)
    CreateModelFormForListForm = replace(CreateModelFormForListForm, "[handleAddRowButtonClick]", handleAddRowButtonClick)
    
    CopyToClipboard CreateModelFormForListForm
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "Form.tsx"
    
    WriteToFile filePath, CreateModelFormForListForm, SeqModelID
    
    
End Function

Public Function GenerateOnKeyDownForListForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GenerateOnKeyDownForListForm = GetReplacedTemplate(rs, "OnKeyDown For List Form")
    CopyToClipboard GenerateOnKeyDownForListForm

End Function

Public Function Generate_getModelsSimpleFilter(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetReplacedTemplate(rs, "getModels no complex filter")
    
    Dim findOptions: findOptions = Generate_findOptionsCopy(frm, SeqModelID)
    Dim filterSnippets As New clsArray
    
    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " ORDER BY FilterOrder"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim uniqueNames As New clsArray
    
    Do Until rs.EOF
        Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID"): If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
        If Not uniqueNames.InArray(FilterQueryName) Then
            uniqueNames.Add FilterQueryName, True
            filterSnippets.Add GenerateSimpleFilterFieldSnippet(frm, SeqModelFilterID)
        End If
        rs.MoveNext
    Loop
    
    Dim filters: filters = filterSnippets.JoinArr(vbNewLine)
    
    TemplateContent = replace(TemplateContent, "[Filters]", filters)
    TemplateContent = replace(TemplateContent, "[FindOptions]", findOptions)
    
    Generate_getModelsSimpleFilter = TemplateContent
    
    Dim lines As New clsArray
    lines.Add "//Generated by Generate_getModelsSimpleFilter"
    lines.Add Generate_getModelsSimpleFilter
    Generate_getModelsSimpleFilter = lines.NewLineJoin
    
    CopyToClipboard Generate_getModelsSimpleFilter
    
End Function

Public Function GenerateTRequiredListDeclaration(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    Dim requiredList As New clsArray
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND (DataType = ""ENUM"" OR Not AllowedOptions IS NULL)"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim PluralizedFieldName: PluralizedFieldName = rs.fields("PluralizedFieldName"): If ExitIfTrue(isFalse(PluralizedFieldName), "PluralizedFieldName is empty..") Then Exit Function
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        If DataTypeInterface = "number" And DataType <> "ENUM" Then
            requiredList.Add PluralizedFieldName & ": number[]"
        Else
            requiredList.Add PluralizedFieldName & ": BasicModel[]"
        End If
        rs.MoveNext
    Loop
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE BackendProjectID = " & BackendProjectID & _
        " AND LeftModelID = " & SeqModelID & " AND NOT ExcludeInRequiredList"
    Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim RightModelID: RightModelID = rs.fields("RightModelID"): If ExitIfTrue(isFalse(RightModelID), "RightModelID is empty..") Then Exit Function
        
        sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & RightModelID
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        
        Dim VariablePluralName: VariablePluralName = rs2.fields("VariablePluralName"): If ExitIfTrue(isFalse(VariablePluralName), "VariablePluralName is empty..") Then Exit Function
        requiredList.Add VariablePluralName & ": BasicModel[]"
        rs.MoveNext
    Loop
    
    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND Not ListVariableName IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim ListVariableName: ListVariableName = rs.fields("ListVariableName")
        requiredList.Add ListVariableName & ": BasicModel[]"
        rs.MoveNext
    Loop
    
    ''Related Model enum lists
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE RightModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & LeftModelID & " AND (DataType = ""ENUM"" OR Not isNull(AllowedOptions))"
        Set rs2 = ReturnRecordset(sqlStr)
        Do Until rs2.EOF
            PluralizedFieldName = rs2.fields("PluralizedFieldName"): If ExitIfTrue(isFalse(PluralizedFieldName), "PluralizedFieldName is empty..") Then Exit Function
            requiredList.Add PluralizedFieldName & ": BasicModel[]"
            rs2.MoveNext
        Loop
        rs.MoveNext
    Loop
    
    requiredList.Add "isRequiredListLoading : boolean"

    Dim lines As New clsArray
    lines.Add "//Generated by GenerateTRequiredListDeclaration"
    lines.Add "export type TRequiredList = { " & requiredList.JoinArr(";" & vbNewLine) & " };"
    
    GenerateTRequiredListDeclaration = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateTRequiredListDeclaration
        
End Function

Public Function CreateModelSchemaArrayValidation(frm As Object, Optional SeqModelID = "", Optional ForeignKeyName = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    Dim fields As New clsArray
    Dim filterStr: filterStr = "SeqModelID = " & SeqModelID & " AND NOT PrimaryKey"
    
    If Not isFalse(ForeignKeyName) Then
        filterStr = filterStr & " AND FieldName <> " & Esc(ForeignKeyName)
    End If
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE " & filterStr & " ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID"): If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
        fields.Add CreateFieldSchema(frm, SeqModelFieldID, True)
        rs.MoveNext
    Loop
    
    CreateModelSchemaArrayValidation = PluralizedModelName & ": Yup.array().of(Yup.object().shape({" & fields.JoinArr("," & vbNewLine) & "}))"
    CopyToClipboard CreateModelSchemaArrayValidation
    
End Function

Public Function GetModelSchemaForListForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    
    Dim fields As New clsArray
    Dim filterStr: filterStr = "SeqModelID = " & SeqModelID & " AND NOT PrimaryKey AND NOT (FieldName = ""createdAt"" OR FieldName = ""updatedAt"")"
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE " & filterStr & " ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID"): If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
        fields.Add CreateFieldSchema(frm, SeqModelFieldID, True)
        rs.MoveNext
    Loop
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("ModelSchema for List Form")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[Fields]", fields.JoinArr("," & vbNewLine))
    replacedContent = replace(replacedContent, "[PluralizedModelName]", PluralizedModelName)
    replacedContent = replace(replacedContent, "[ModelName]", ModelName)
    
    GetModelSchemaForListForm = replacedContent
    CopyToClipboard GetModelSchemaForListForm
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\schema\" & ModelName & "Schema.ts"
    WriteToFile filePath, GetModelSchemaForListForm, SeqModelID
    
End Function
    
Public Function GenerateFormikArrayTableHead(frm As Object, Optional SeqModelID = "", Optional ForeignKey = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim filters As New clsArray
    filters.Add "SeqModelID = " & SeqModelID
    filters.Add "ControlType <> ""Hidden"""
    filters.Add "NOT PrimaryKey"
    
    If Not isFalse(ForeignKey) Then
        filters.Add "FieldName <> " & Esc(ForeignKey)
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE " & filters.JoinArr(" AND ") & " ORDER BY FieldOrder"
    
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        fields.Add GenerateFieldTableCellHeader(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Form Array TableHead")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[Fields]", fields.JoinArr(vbNewLine))
    
    GenerateFormikArrayTableHead = replacedContent
    
    GenerateFormikArrayTableHead = GetGeneratedByFunctionSnippet(GenerateFormikArrayTableHead, "GenerateFormikArrayTableHead", True)
    
    CopyToClipboard GenerateFormikArrayTableHead
    
End Function

Public Function GenerateFormikArrayControls(frm As Object, Optional SeqModelID = "", Optional ForeignKey = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    Dim filters As New clsArray
    filters.Add "SeqModelID = " & SeqModelID
    filters.Add "ControlType <> ""Hidden"""
    filters.Add "NOT PrimaryKey"
    
    If Not isFalse(ForeignKey) Then
        filters.Add "FieldName <> " & Esc(ForeignKey)
    End If
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE " & filters.JoinArr(" AND ") & " ORDER BY FieldOrder"
    
    Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        fields.Add "<TableCell>" & GenerateIndividualFormControlArray(frm, SeqModelFieldID) & "</TableCell>"
        rs.MoveNext
    Loop
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("Form Array Form Controls")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[Fields]", fields.JoinArr(vbNewLine))
    replacedContent = replace(replacedContent, "[PluralizedModelName]", PluralizedModelName)
    
    GenerateFormikArrayControls = replacedContent
    GenerateFormikArrayControls = GetGeneratedByFunctionSnippet(GenerateFormikArrayControls, "GenerateFormikArrayControls", True)
    CopyToClipboard GenerateFormikArrayControls
    
End Function

Public Function Generate_handleAddRowButtonClick(frm As Object, Optional SeqModelID = "", Optional RightModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        Dim AllowNull: AllowNull = rs.fields("AllowNull")
        Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption")
        Dim options As New clsArray
        
        If Not isFalse(RightModelID) Then
            sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftForeignKey = " & Esc(fieldName) & " AND LeftModelID = " & SeqModelID & _
            " AND RightModelID = " & RightModelID
            Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
            If rs2.EOF Then
                If DataType = "DECIMAL" Then
                    fields.Add fieldName & ": " & IIf(AllowNull, """""", """0.00""")
                Else
                    fields.Add fieldName & ": """""
                End If
            Else
                Dim RightVariableName: RightVariableName = rs2.fields("RightVariableName"): If ExitIfTrue(isFalse(RightVariableName), "RightVariableName is empty..") Then Exit Function
                fields.Add fieldName & ":" & RightVariableName & " ? " & RightVariableName & ".id : """""
            End If
        Else
            If DataType = "ENUM" Then
                If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
                Set options = ConvertEnumToArray(DataTypeOption)
                fields.Add fieldName & ": " & IIf(AllowNull, """""", options.arr(0))
            ElseIf DataType = "DECIMAL" Then
                fields.Add fieldName & ": " & IIf(AllowNull, """""", """0.00""")
            Else
                fields.Add fieldName & ": """""
            End If
            
        End If
        rs.MoveNext
    Loop
    
    Dim TemplateContent: TemplateContent = GetTemplateContent("handleAddRowButtonClick")
    
    Dim replacedContent
    replacedContent = replace(TemplateContent, "[Fields]", fields.JoinArr("," & vbNewLine))
    
    Generate_handleAddRowButtonClick = replacedContent
    Generate_handleAddRowButtonClick = GetGeneratedByFunctionSnippet(Generate_handleAddRowButtonClick, "Generate_handleAddRowButtonClick")
    CopyToClipboard Generate_handleAddRowButtonClick
    
End Function


Public Function CreateModelMultilineForm(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    
    Dim TemplateContent: TemplateContent = GetReplacedTemplate(rs, "FieldArray Form For List Form")
    
    Dim TableHead: TableHead = GenerateFormikArrayTableHead(frm, SeqModelID)
    Dim TableBody: TableBody = GenerateFormikArrayControls(frm, SeqModelID)
    
    TemplateContent = replace(TemplateContent, "[TableHead]", TableHead)
    TemplateContent = replace(TemplateContent, "[TableBody]", TableBody)
    
    CreateModelMultilineForm = TemplateContent
    CopyToClipboard CreateModelMultilineForm
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "MultiLineForm.tsx"
    
    WriteToFile filePath, CreateModelMultilineForm, SeqModelID
    
End Function

Public Function Generate_handleAddRowButtonClickForListForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr, rs As Recordset
    sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    
    Dim TemplateContent: TemplateContent = GetReplacedTemplate(rs, "handleAddRowButtonClick for List Form")
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " ORDER BY FIeldOrder"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption")
        Dim AllowNull: AllowNull = rs.fields("AllowNull")
        Dim AllowedOptions: AllowedOptions = rs.fields("AllowedOptions")
        Dim options As New clsArray
        
        sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND LeftForeignKey = " & Esc(fieldName)
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        
        If DataType = "ENUM" Then
            If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
            Set options = ConvertEnumToArray(DataTypeOption)
            fields.Add fieldName & ": " & IIf(AllowNull, """""", options.arr(0))
        ElseIf DataType = "DECIMAL" Then
            fields.Add fieldName & ": " & IIf(AllowNull, """""", """0.00""")
        ElseIf Not IsNull(AllowedOptions) Then
            options.arr = AllowedOptions
            fields.Add fieldName & ": " & Esc(options.arr(0))
        ElseIf Not rs2.EOF Then
            Dim RightVariablePluralName: RightVariablePluralName = rs2.fields("RightVariablePluralName"): If ExitIfTrue(isFalse(RightVariablePluralName), "RightVariablePluralName is empty..") Then Exit Function
            fields.Add fieldName & ": requiredListObject." & RightVariablePluralName & " && " & _
                "requiredListObject." & RightVariablePluralName & ".length > 0 ? " & _
                "requiredListObject." & RightVariablePluralName & "[0].id.toString() : """""
        Else
            fields.Add fieldName & ": """""
        End If
            
        rs.MoveNext
    Loop
    
    Generate_handleAddRowButtonClickForListForm = replace(TemplateContent, "[Fields]", fields.JoinArr("," & vbNewLine))
    CopyToClipboard Generate_handleAddRowButtonClickForListForm
    
End Function

Public Function GenerateModelRouteForListForm(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim RouteFileName: RouteFileName = rs.fields("RouteFileName"): If ExitIfTrue(isFalse(RouteFileName), "RouteFileName is empty..") Then Exit Function
    
    GenerateModelRouteForListForm = GetReplacedTemplate(rs, "ModelRoute for List Form")
    CopyToClipboard GenerateModelRouteForListForm
    
    Dim ProjectPath: ProjectPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ProjectPath")
    Dim filePath: filePath = ProjectPath & "src\routes\" & RouteFileName
    
    WriteToFile filePath, GenerateModelRouteForListForm, SeqModelID
    
End Function

Public Function Generate_updateModelsFunction(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = IIf(IsNull(SlugField), "id", "slug")
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    Dim replacedContent
    replacedContent = GetReplacedTemplate(rs, "updateModels For List Form")
    
    Dim ModelInsertUpdate, ModelDelete
    ModelInsertUpdate = GetReplacedTemplate(rs, "Backend Related Table Update or Insert")
    ModelInsertUpdate = replace(ModelInsertUpdate, "[SlugOrID]", SlugOrId)
    
    ModelDelete = GetReplacedTemplate(rs, "Backend Related Table Delete")
    
    Dim fields As New clsArray, enumValidations As New clsArray, requiredFields As New clsArray
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim PrimaryKey: PrimaryKey = rs.fields("PrimaryKey")
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        Dim VerboseFieldName: VerboseFieldName = rs.fields("VerboseFieldName"): If ExitIfTrue(isFalse(VerboseFieldName), "VerboseFieldName is empty..") Then Exit Function
        Dim DataType: DataType = rs.fields("DataType")
        Dim AllowNull: AllowNull = rs.fields("AllowNull")
        
        Dim ValueString: ValueString = "item." & fieldName
        
        If PrimaryKey Or DataTypeInterface = "number" Then ValueString = "parseInt(" & ValueString & ")"
        
        ''Add Enum validation then replace the
        If DataType = "ENUM" Then
        
            ValueString = ValueString & " as T" & ModelName & "[" & Esc(fieldName) & "]"
            
            Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
            Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
            enumValidations.Add Get_updateModelsEnumValidation(frm, SeqModelFieldID)
            
        End If
        
        If DataType = "DECIMAL" Then
            
            ValueString = "convertStringToFloat(" & ValueString & ").toString()"
            
        End If
        
        ''Required validations
        If Not AllowNull Then
            requiredFields.Add fieldName & ":" & Esc(VerboseFieldName)
        End If
            
        fields.Add fieldName & ": " & ValueString
        rs.MoveNext
    Loop
    
    ''Validation portion of the code
    Dim lines As New clsArray
    lines.Add "//Validate each item and should not be null"
    lines.Add "const validationMessage = validateFieldIfBlank(item,{" & requiredFields.JoinArr(",") & "});"
    lines.Add "if (validationMessage) {throw Error(validationMessage);}"
    
    ModelInsertUpdate = replace(ModelInsertUpdate, "[Fields]", fields.JoinArr("," & vbNewLine))
    ModelInsertUpdate = replace(ModelInsertUpdate, "[updateModelsEnumValidation]", enumValidations.JoinArr(vbNewLine))
    ModelInsertUpdate = replace(ModelInsertUpdate, "[updateModelsCheckTruthinessOfFields]", lines.JoinArr(vbNewLine))
    
    Generate_updateModelsFunction = replace(replacedContent, "[ModelInsertUpdate]", ModelInsertUpdate)
    Generate_updateModelsFunction = replace(Generate_updateModelsFunction, "[ModelDelete]", ModelDelete)
    Generate_updateModelsFunction = replace(Generate_updateModelsFunction, "[ValidateUniqueness]", GenerateUniquenessValidation(frm, SeqModelID))
    
    Generate_updateModelsFunction = GetGeneratedByFunctionSnippet(Generate_updateModelsFunction, "Generate_updateModelsFunction")
    
    CopyToClipboard Generate_updateModelsFunction
    
End Function

Public Function GenerateModelListBodyForListForm(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    GenerateModelListBodyForListForm = GetReplacedTemplate(rs, "ModelListBody for List Form")
    
    CopyToClipboard GenerateModelListBodyForListForm
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\components\" & VariableName & "\" & ModelName & "ListBody.tsx"
    
    WriteToFile filePath, GenerateModelListBodyForListForm, SeqModelID
    
End Function


Public Function Generate_addSlugs(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    Generate_addSlugs = GetReplacedTemplate(rs, "addSlugs")
    CopyToClipboard Generate_addSlugs
    
    Dim ProjectPath: ProjectPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ProjectPath")
    Dim filePath: filePath = ProjectPath & "src\standalone\addSlugs.ts"
    WriteToFile filePath, Generate_addSlugs, SeqModelID
    
End Function

Public Function GenerateSyncModel(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    GenerateSyncModel = GetReplacedTemplate(rs, "Model Sync")
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateSyncModel"
    lines.Add GenerateSyncModel
    
    GenerateSyncModel = lines.NewLineJoin
    
    CopyToClipboard GenerateSyncModel
    
    
End Function

Public Function GenerateModelNavbarItem(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GenerateModelNavbarItem = GetReplacedTemplate(rs, "Model Navbar Item")
    CopyToClipboard GenerateModelNavbarItem
    
End Function

Public Function GenerateSyncCodeForIndex(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GenerateSyncCodeForIndex = GetReplacedTemplate(rs, "Import Sync To Index")
    CopyToClipboard GenerateSyncCodeForIndex
    
End Function

Public Function GenerateNext13APISetup(frm As Object, Optional BackendProjectID = "")
    
    RunCommandSaveRecord
        
    If isFalse(BackendProjectID) Then
        BackendProjectID = frm("BackEndProjectID")
        If ExitIfTrue(isFalse(BackendProjectID), "BackEndProjectID is empty..") Then Exit Function
    End If
    
    Dim response: response = MsgBox("This will overwrite all the files and be replaced by the newer file if it exists" & vbNewLine & _
        "Do you want to proceed?", vbYesNo)
    
    If vbYes Then
        NoHasWriteToFilePrompt = True
        Dim sqlStr: sqlStr = "SELECT * FROM qryFunctionChainItems WHERE FunctionChainName = ""Next 13 API setup"" ORDER BY FunctionOrder ASC"
        Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        Do Until rs.EOF
            Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
            Run FunctionName, frm
            rs.MoveNext
        Loop
        NoHasWriteToFilePrompt = False
    End If
    
    MsgBox "Complete files for Next 13 API Setup generated.", vbOKOnly
    
End Function

Public Function GenerateAllModelFilesComplete(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
        
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim response: response = MsgBox("This will overwrite all the files and be replaced by the newer file if it exists" & vbNewLine & _
        "Do you want to proceed?", vbYesNo)
    
    If vbYes Then
        NoHasWriteToFilePrompt = True
        Dim sqlStr: sqlStr = "SELECT * FROM qryFunctionChainItems WHERE FunctionChainName = ""Specific Model with Detail Form"" ORDER BY FunctionOrder ASC"
        Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        Do Until rs.EOF
            Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
            Run FunctionName, frm, SeqModelID
            rs.MoveNext
        Loop
        NoHasWriteToFilePrompt = False
    End If
    
    MsgBox "Complete model files were generated.", vbOKOnly
    
End Function

Public Function GenerateAllModelFilesForListForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim response: response = MsgBox("This will overwrite all the files and be replaced by the newer file if it exists" & vbNewLine & _
        "Do you want to proceed?", vbYesNo)

    If response = vbYes Then
        NoHasWriteToFilePrompt = True
        Dim sqlStr: sqlStr = "SELECT * FROM qryFunctionChainItems WHERE FunctionChainName = ""Specific Model With List Form"" ORDER BY FunctionOrder ASC"
        Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        Do Until rs.EOF
            Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
            Run FunctionName, frm, SeqModelID
            rs.MoveNext
        Loop
        NoHasWriteToFilePrompt = False
    End If
    
    MsgBox "Complete model list files were generated.", vbOKOnly
End Function

Public Function ImportAsRelatedModelBackend(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim UseApp: UseApp = rs.fields("UseApp")

    ImportAsRelatedModelBackend = GetReplacedTemplate(rs, IIf(UseApp, "Import as Related Model Next 13", "Import as Related Model")) & "//Generated by ImportAsRelatedModelBackend"
    CopyToClipboard ImportAsRelatedModelBackend
    
End Function

Public Function GenerateEnumValidation(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateEnumValidation"
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND DataType = ""ENUM"" ORDER BY FieldOrder"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID"): If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
        lines.Add GenerateEnumValidationForSingleField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GenerateEnumValidation = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateEnumValidation
    
End Function

Public Function GenerateUniquenessValidation(frm As Object, Optional SeqModelID = "") As String
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Unique"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim VerboseFieldName: VerboseFieldName = rs.fields("VerboseFieldName"): If ExitIfTrue(isFalse(VerboseFieldName), "VerboseFieldName is empty..") Then Exit Function
        fields.Add fieldName & ": " & Esc(VerboseFieldName)
        rs.MoveNext
    Loop
    
    sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID & " AND NOT SlugField IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    
    GenerateUniquenessValidation = GetReplacedTemplate(rs, "Validate Uniqueness for submitted list with relationship")
    GenerateUniquenessValidation = replace(GenerateUniquenessValidation, "[Fields]", fields.JoinArr)
    
    Dim lines As New clsArray
    
    If fields.count > 0 Then
        lines.Add "//Generated by GenerateUniquenessValidation"
        lines.Add GenerateUniquenessValidation
    End If
    
    GenerateUniquenessValidation = lines.NewLineJoin
    CopyToClipboard GenerateUniquenessValidation
    
End Function

Public Function GenerateChildUpdateorInsert(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim templateName As String: templateName = IIf(IsNull(SlugField), "Backend Related Table Update or Insert No CreatedIds no slug", "Backend Related Table Update or Insert No CreatedIds")
    
    GenerateChildUpdateorInsert = GetReplacedTemplate(rs, templateName)
    GenerateChildUpdateorInsert = replace(GenerateChildUpdateorInsert, "[Fields]", GenerateCreateUpdateFieldsAsChild(frm, SeqModelID))
    GenerateChildUpdateorInsert = replace(GenerateChildUpdateorInsert, "[ChildEnumValidation]", GenerateChildrenEnumValidation(frm, SeqModelID))
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateChildUpdateorInsert"
    lines.Add GenerateChildUpdateorInsert
    
    GenerateChildUpdateorInsert = lines.NewLineJoin
    
    CopyToClipboard GenerateChildUpdateorInsert
    
End Function

Public Function GenerateChildDelete(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GenerateChildDelete = GetReplacedTemplate(rs, "Backend Related Table Delete")
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateChildDelete"
    lines.Add GenerateChildDelete
    GenerateChildDelete = lines.NewLineJoin
    CopyToClipboard GenerateChildDelete
    
End Function

Public Function GenerateCreateUpdateFieldsAsChild(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim lines As New clsArray
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey ORDER BY FieldOrder"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID"): If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
        lines.Add GenerateCreateUpdateFieldAsChild(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GenerateCreateUpdateFieldsAsChild = "//Generated by GenerateCreateUpdateFieldsAsChild" & vbNewLine & lines.JoinArr("," & "//Generated by GenerateCreateUpdateFieldAsChild" & vbNewLine)
    CopyToClipboard GenerateCreateUpdateFieldsAsChild
    
End Function

Public Function GenerateMainModelUniqueValidation(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GenerateMainModelUniqueValidation = GetReplacedTemplate(rs, "Main Model Uniqueness Validation")
    CopyToClipboard GenerateMainModelUniqueValidation
    
End Function

Public Function getRelationshipBodyDeclarationForAdd(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim lines As New clsArray
    lines.Add "//Generated by getRelationshipBodyDeclarationForAdd"
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey ORDER BY FieldOrder"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim UseApp: UseApp = rs.fields("UseApp")
    
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        fields.Add fieldName
        rs.MoveNext
    Loop
    
    lines.Add "const { " & fields.JoinArr & " } = " & IIf(UseApp, "res", "body") & ";"
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE RightModelID = " & SeqModelID: Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
        lines.Add "const " & LeftPluralizedModelName & " = body." & LeftPluralizedModelName & ";"
        rs.MoveNext
    Loop
    
    getRelationshipBodyDeclarationForAdd = lines.JoinArr(vbNewLine)
    CopyToClipboard getRelationshipBodyDeclarationForAdd
    
End Function


Public Function getMainModelUniquenessValidationForAdd(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Unique ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    
    If Not rs.EOF Then
    
        sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
        Set rs = ReturnRecordset(sqlStr)
        
        getMainModelUniquenessValidationForAdd = GetReplacedTemplate(rs, "Main Model Uniqueness Validation for Add")
        getMainModelUniquenessValidationForAdd = replace(getMainModelUniquenessValidationForAdd, "[Fields]", GetUniquenessWhereFields(frm, SeqModelID))
        getMainModelUniquenessValidationForAdd = replace(getMainModelUniquenessValidationForAdd, "[FieldCombination]", GetUniqueCombinationCaption(frm, SeqModelID))
        
        Dim lines As New clsArray
        lines.Add "//Generated by getMainModelUniquenessValidationForAdd"
        lines.Add getMainModelUniquenessValidationForAdd
        getMainModelUniquenessValidationForAdd = lines.NewLineJoin
    Else
        getMainModelUniquenessValidationForAdd = ""
    End If
    
    CopyToClipboard getMainModelUniquenessValidationForAdd
    
End Function

Public Function getChildrenUniquenessValidationForAdd(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim lines As New clsArray
    
    lines.Add "//Generated by getChildrenUniquenessValidationForAdd"
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE RightModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        lines.Add GenerateUniquenessValidation(frm, LeftModelID)
        rs.MoveNext
    Loop
    
    getChildrenUniquenessValidationForAdd = lines.JoinArr(vbNewLine)
    CopyToClipboard getChildrenUniquenessValidationForAdd
    
End Function

Public Function getChildrenUniquenessValidationWithDatabaseForAdd(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim lines As New clsArray
    lines.Add "//Generated by getChildrenUniquenessValidationWithDatabaseForAdd"
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE RightModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & LeftModelID
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        lines.Add GetReplacedTemplate(rs2, "Uniqueness Validation For Add With Database as a Child")
        rs.MoveNext
    Loop
    
    getChildrenUniquenessValidationWithDatabaseForAdd = lines.JoinArr(vbNewLine)
    CopyToClipboard getChildrenUniquenessValidationWithDatabaseForAdd
    
End Function

Public Function getModelFieldsValue(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey ORDER BY FieldOrder"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim ModelFieldsValueArr As New clsArray
    Do Until rs.EOF
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID"): If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
        ModelFieldsValueArr.Add GenerateCreateUpdateField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    getModelFieldsValue = ModelFieldsValueArr.JoinArr(", //Generated by GenerateCreateUpdateField" & vbNewLine)
    
    Dim lines As New clsArray
    lines.Add "//Generated by getModelFieldsValue"
    lines.Add getModelFieldsValue
    getModelFieldsValue = lines.NewLineJoin
    
    CopyToClipboard getModelFieldsValue
    
End Function



Public Function GetChildrenInsertsForAdd(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE RightModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim lines As New clsArray
    lines.Add "//Generated by GetChildrenInsertsForAdd"
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        Dim RightModelID: RightModelID = rs.fields("RightModelID"): If ExitIfTrue(isFalse(RightModelID), "RightModelID is empty..") Then Exit Function
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset("SELECT * FROM tblSeqModels WHERE SeqModelID = " & LeftModelID)
        
        Dim TemplateContent: TemplateContent = GetReplacedTemplate(rs2, "Children Inserts For Add")
        TemplateContent = replace(TemplateContent, "[Fields]", GetModelFieldsForAdd(frm, LeftModelID, RightModelID))
        TemplateContent = replace(TemplateContent, "[ChildrenEnumValidation]", GenerateChildrenEnumValidation(frm, LeftModelID))
        lines.Add TemplateContent
        
        rs.MoveNext
    Loop
    
    GetChildrenInsertsForAdd = lines.JoinArr(vbNewLine)
    CopyToClipboard GetChildrenInsertsForAdd
    
End Function

Public Function GetModelFieldsForAdd(frm As Object, Optional SeqModelID = "", Optional RightModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND RightModelID = " & RightModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim LeftForeignKey: LeftForeignKey = rs.fields("LeftForeignKey"): If ExitIfTrue(isFalse(LeftForeignKey), "LeftForeignKey is empty..") Then Exit Function
    Dim RightVariableName: RightVariableName = rs.fields("RightVariableName"): If ExitIfTrue(isFalse(RightVariableName), "RightVariableName is empty..") Then Exit Function
    Dim LeftModelName: LeftModelName = rs.fields("LeftModelName"): If ExitIfTrue(isFalse(LeftModelName), "LeftModelName is empty..") Then Exit Function
    
    Set rs = ReturnRecordset("SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey ORDER BY FieldOrder")
    
    Dim fields As New clsArray
    Do Until rs.EOF
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        Dim AllowNull: AllowNull = rs.fields("AllowNull")
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface")
        
        Dim fieldValue
        fieldValue = "item." & fieldName
        
        If AllowNull Then
            fieldValue = fieldValue & " ? null : " & fieldValue
        End If
        
        If DataTypeInterface = "number" Then
            fieldValue = "parseInt(" & fieldValue & ")"
        End If
        
'        If DataType = "DECIMAL" Then
'            fieldValue = "convertStringToFloat(" & fieldValue & ")"
'        End If
        
        If DataType = "DATEONLY" Then
            fieldValue = "convertDateStringToYYYYMMDD(" & fieldValue & ")"
        End If
        
        If DataType = "ENUM" Then
            fieldValue = fieldValue & " as T" & LeftModelName & "[" & Esc(fieldName) & "]"
        End If
        
        If LeftForeignKey = fieldName Then
            fields.Add fieldName & ": " & RightVariableName & ".id"
        Else
            fields.Add fieldName & ": " & fieldValue
        End If
        
        rs.MoveNext
    Loop
    
    GetModelFieldsForAdd = fields.JoinArr("," & vbNewLine)
    CopyToClipboard GetModelFieldsForAdd
    
End Function

Public Function GetAdditionalImportBasedOnListVariableName(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    lines.Add "//Generated by GetAdditionalImportBasedOnListVariableName"
    
    ''Also peek at the filter used by this model and look for list variable name
    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND Not ListVariableName IS NULL AND SeqModelFieldID IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim ListVariableName: ListVariableName = rs.fields("ListVariableName")
        sqlStr = "SELECT * FROM tblSeqModels WHERE VariablePluralName = " & Esc(ListVariableName)
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        lines.Add GetReplacedTemplate(rs2, "Import Based on ListVariableName")
        rs.MoveNext
    Loop
    
    '''import { SubAccountTitleModel } from "../../interfaces/SubAccountTitleInterfaces";
    GetAdditionalImportBasedOnListVariableName = lines.JoinArr(vbNewLine)
    CopyToClipboard GetAdditionalImportBasedOnListVariableName
    
End Function

Public Function CreateRelatedModelFormHookImports(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    ''Also peek at the filter used by this model and look for list variable name
    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND Not ListVariableName IS NULL AND SeqModelFieldID IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim ListVariableName: ListVariableName = rs.fields("ListVariableName")
        sqlStr = "SELECT * FROM tblSeqModels WHERE VariablePluralName = " & Esc(ListVariableName)
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        lines.Add GetReplacedTemplate(rs2, "Import Based on ListVariableName")
        rs.MoveNext
    Loop
    
    '''import { SubAccountTitleModel } from "../../interfaces/SubAccountTitleInterfaces";

    CreateRelatedModelFormHookImports = lines.JoinArr(vbNewLine)
    CopyToClipboard CreateRelatedModelFormHookImports
    
End Function

Public Function GetFormikArrayTotalComputation(frm As Object, Optional SeqModelID = "", Optional ForeignKey = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    Dim filters As New clsArray
    filters.Add "SeqModelID = " & SeqModelID
    filters.Add "ControlType <> ""Hidden"""
    filters.Add "NOT PrimaryKey"
    
    If Not isFalse(ForeignKey) Then
        filters.Add "FieldName <> " & Esc(ForeignKey)
    End If
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE " & filters.JoinArr(" AND ") & " ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Set rs = ReturnRecordset(sqlStr)
    
    Dim totalFields As New clsArray, incrementals As New clsArray
    
    Do Until rs.EOF
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        
        If (DataType = "DECIMAL" Or DataTypeInterface = "number") And Not fieldName Like "*id*" Then
            totalFields.Add "total_" & fieldName & " = 0"
            
            If DataType = "DECIMAL" Then
                incrementals.Add "total_" & fieldName & " += parseFloat(item." & fieldName & ");"
            Else
                incrementals.Add "total_" & fieldName & " += parseInt(item." & fieldName & ");"
            End If
            
        End If
        
        rs.MoveNext
    Loop
    
    Dim letDeclaration: letDeclaration = "let " & totalFields.JoinArr(",")
    Dim formikLoop: formikLoop = "formik.values." & PluralizedModelName & ".forEach((item) => {" & incrementals.JoinArr(vbNewLine) & "});"
    
    lines.Add "//Compute field totals for this record"
    lines.Add letDeclaration
    lines.Add formikLoop

    GetFormikArrayTotalComputation = lines.JoinArr(vbNewLine)
    GetFormikArrayTotalComputation = GetGeneratedByFunctionSnippet(GetFormikArrayTotalComputation, "GetFormikArrayTotalComputation")
    CopyToClipboard GetFormikArrayTotalComputation
    
End Function

Public Function GetFormikArrayTableFooterForTotal(frm As Object, Optional SeqModelID = "", Optional ForeignKey = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    
    Dim filters As New clsArray
    filters.Add "SeqModelID = " & SeqModelID
    filters.Add "ControlType <> ""Hidden"""
    filters.Add "NOT PrimaryKey"
    
    If Not isFalse(ForeignKey) Then
        filters.Add "FieldName <> " & Esc(ForeignKey)
    End If
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE " & filters.JoinArr(" AND ") & " ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim fields As New clsArray
    
    Do Until rs.EOF
        Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        
        If (DataType = "DECIMAL" Or DataTypeInterface = "number") And Not fieldName Like "*id*" Then
            
            
            If DataType = "DECIMAL" Then
                fields.Add "<TableCell align=""right""><Box sx={{ pr: ""14px"" }}>{formatCurrency(total_" & fieldName & ")}</Box></TableCell>"
            Else
                fields.Add "<TableCell align=""right""><Box sx={{ pr: ""14px"" }}>{total_" & fieldName & "}</Box></TableCell>"
            End If
        Else
            fields.Add "<TableCell></TableCell>"
        End If
        
        rs.MoveNext
    Loop
    
    GetFormikArrayTableFooterForTotal = GetReplacedTemplate(rs, "FieldArray Table Footer")
    GetFormikArrayTableFooterForTotal = replace(GetFormikArrayTableFooterForTotal, "[Fields]", fields.JoinArr(vbNewLine))
    GetFormikArrayTableFooterForTotal = GetGeneratedByFunctionSnippet(GetFormikArrayTableFooterForTotal, "GetFormikArrayTableFooterForTotal", True)
    
    CopyToClipboard GetFormikArrayTableFooterForTotal
    
End Function

Public Function GetSidebarLink(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSidebarLink = GetReplacedTemplate(rs, "SidebarLink")
    CopyToClipboard GetSidebarLink
    
End Function

Public Function GetUniquenessWhereFields(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Unique ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
        Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
        
        Dim fieldValue: fieldValue = fieldName
        
        If DataTypeInterface = "number" Then
            fieldValue = "parseInt(" & fieldValue & ")"
        End If
        
        lines.Add fieldName & ":" & fieldValue
        rs.MoveNext
    Loop
    
    GetUniquenessWhereFields = lines.JoinArr(",")
    CopyToClipboard GetUniquenessWhereFields
    
End Function

Public Function GetUniqueCombinationCaption(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Unique ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim VerboseFieldName: VerboseFieldName = rs.fields("VerboseFieldName"): If ExitIfTrue(isFalse(VerboseFieldName), "VerboseFieldName is empty..") Then Exit Function
        lines.Add Esc(VerboseFieldName)
        rs.MoveNext
    Loop
    
    GetUniqueCombinationCaption = lines.JoinArr(",")
    CopyToClipboard GetUniqueCombinationCaption
    
End Function

Public Function GenerateChildrenEnumValidation(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord
    
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim lines As New clsArray
    lines.Add "//Generated by GenerateChildrenEnumValidation"
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND DataType = ""ENUM"" ORDER BY FieldOrder"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        
        
        Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
        Dim VerboseFieldName: VerboseFieldName = rs.fields("VerboseFieldName"): If ExitIfTrue(isFalse(VerboseFieldName), "VerboseFieldName is empty..") Then Exit Function
        
        Dim options As New clsArray: Set options = ConvertEnumToArray(DataTypeOption)
        
        Dim fieldName: fieldName = "item." & rs.fields("FieldName")
        
        Dim TemplateContent
        TemplateContent = GetTemplateContent("Enum Validation")
        TemplateContent = replace(TemplateContent, "[Enums]", "[" & options.JoinArr(",") & "]")
        TemplateContent = replace(TemplateContent, "[FieldName]", fieldName)
        TemplateContent = replace(TemplateContent, "[VerboseFieldName]", VerboseFieldName)
        lines.Add TemplateContent
        rs.MoveNext
        
    Loop
    
    GenerateChildrenEnumValidation = lines.JoinArr(vbNewLine)
    CopyToClipboard GenerateChildrenEnumValidation

End Function

Public Function Generate_findOptionsCopy(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Generate_findOptionsCopy = GetReplacedTemplate(rs, "findOptionsCopy")
    Generate_findOptionsCopy = GetGeneratedByFunctionSnippet(Generate_findOptionsCopy, "Generate_findOptionsCopy")
    
    CopyToClipboard Generate_findOptionsCopy
    
End Function

Public Function GenerateModelSyncRoute(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID")
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    
    GenerateModelSyncRoute = GetReplacedTemplate(rs, "Model Sync File")
    GenerateModelSyncRoute = GetGeneratedByFunctionSnippet(GenerateModelSyncRoute, "GenerateModelSyncRoute")
    CopyToClipboard GenerateModelSyncRoute
    
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\app\api\" & ModelPath & "\sync\route.ts"
    
    WriteToFile filePath, GenerateModelSyncRoute, SeqModelID
    
End Function

Public Function RunModelSync(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    ''ModelPath
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    
    ''Port|Host
    sqlStr = "SELECT * FROM tblBackendDatabaseConfigs WHERE BackendProjectID = " & BackendProjectID
    Set rs = ReturnRecordset(sqlStr)
    Dim Port: Port = rs.fields("Port"): If ExitIfTrue(isFalse(Port), "Port is empty..") Then Exit Function
    Dim Host: Host = rs.fields("Host"): If ExitIfTrue(isFalse(Host), "Host is empty..") Then Exit Function
    ''http://localhost:8000/api/users/sync
    Dim endpointUrl As String: endpointUrl = "http://" & Host & ":" & Port & "/api/" & ModelPath & "/sync"
    If isFalse(endpointUrl) Then Exit Function
    
    endpointUrl = endpointUrl & "?rand=" & Rnd()
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", endpointUrl, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.setRequestHeader "Pragma", "no-cache"
    http.send
    
    Dim StatusCode: StatusCode = http.status
    
    Dim json As Object
    Set json = JsonConverter.ParseJson(http.responseText)
    
    Dim prettyJson As String
    prettyJson = JsonConverter.ConvertToJson(json)
    
    prettyJson = UnescapeJson(prettyJson)
    
    Dim TestApi
    TestApi = PrettifyJson(prettyJson)
    TestApi = Left(TestApi, 500)
    
    MsgBox "Returned: " & TestApi & vbNewLine & "Status:" & StatusCode
        
End Function

Public Function Generate_getModelsSimpleFilterNext13(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Generate_getModelsSimpleFilterNext13 = GetReplacedTemplate(rs, "getModels no complex filter next-13")
    
    Dim findOptions: findOptions = Generate_findOptionsCopy(frm, SeqModelID)
    Dim filterSnippets As New clsArray
    
    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " ORDER BY FilterOrder"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim uniqueNames As New clsArray
    
    Do Until rs.EOF
        Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID"): If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
        If Not uniqueNames.InArray(FilterQueryName) Then
            uniqueNames.Add FilterQueryName, True
            filterSnippets.Add GenerateSimpleFilterFieldSnippet(frm, SeqModelFilterID)
        End If
        rs.MoveNext
    Loop
    
    Dim filters: filters = filterSnippets.JoinArr(vbNewLine)
    
    Generate_getModelsSimpleFilterNext13 = replace(Generate_getModelsSimpleFilterNext13, "[Filters]", filters)
    Generate_getModelsSimpleFilterNext13 = replace(Generate_getModelsSimpleFilterNext13, "[FindOptions]", findOptions)
    
    
    Generate_getModelsSimpleFilterNext13 = GetGeneratedByFunctionSnippet(Generate_getModelsSimpleFilterNext13, "Generate_getModelsSimpleFilterNext13")
    CopyToClipboard Generate_getModelsSimpleFilterNext13
    
End Function

Public Function Generate_getModelsAPIRouteNext13(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function

    Generate_getModelsAPIRouteNext13 = GetReplacedTemplate(rs, "getModels API Route Next 13")
    Generate_getModelsAPIRouteNext13 = replace(Generate_getModelsAPIRouteNext13, "[GeneratefindOptions]", GeneratefindOptions(frm, SeqModelID))
    Generate_getModelsAPIRouteNext13 = replace(Generate_getModelsAPIRouteNext13, "[Generate_getModelsSimpleFilterNext13]", Generate_getModelsSimpleFilterNext13(frm, SeqModelID))
    Generate_getModelsAPIRouteNext13 = replace(Generate_getModelsAPIRouteNext13, "[GenerateImportRelatedModels]", GenerateImportRelatedModels(frm, SeqModelID))
    Generate_getModelsAPIRouteNext13 = replace(Generate_getModelsAPIRouteNext13, "[Generate_GetAddFunctionWithRelationshipNext13]", Generate_GetAddFunctionWithRelationshipNext13(frm, SeqModelID))
    
    ''Generate_GetAddFunctionWithRelationshipNext13
    
    Generate_getModelsAPIRouteNext13 = GetGeneratedByFunctionSnippet(Generate_getModelsAPIRouteNext13, "Generate_getModelsAPIRouteNext13")
    CopyToClipboard Generate_getModelsAPIRouteNext13
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID")
    Dim ClientPath: ClientPath = ELookup("tblBackendProjects", "BackendProjectID = " & BackendProjectID, "ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    Dim filePath: filePath = ClientPath & "src\app\api\" & ModelPath & "\route.ts"
    
    WriteToFile filePath, Generate_getModelsAPIRouteNext13, SeqModelID
    
End Function

Public Function GenerateImportRelatedModels(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim lines As New clsArray
    
    ''Relationships where the SeqModelID is a RightModelID
    Dim item, relatedModels As New clsArray: relatedModels.arr = Elookups("qrySeqModelRelationships", "RightModelID = " & SeqModelID, "LeftModelID")
    Dim hasRelationship As Boolean
    For Each item In relatedModels.arr
        hasRelationship = True
        lines.Add ImportAsRelatedModelBackend(frm, item)
    Next item
    
    ''Relationships where the SeqModelID is a LeftModelID
    relatedModels.arr = Elookups("qrySeqModelRelationships", "LeftModelID = " & SeqModelID, "RightModelID")
    For Each item In relatedModels.arr
        hasRelationship = True
        lines.Add ImportAsRelatedModelBackend(frm, item)
    Next item
    
    GenerateImportRelatedModels = lines.JoinArr(vbNewLine)
    GenerateImportRelatedModels = GetGeneratedByFunctionSnippet(GenerateImportRelatedModels, "GenerateImportRelatedModels")
    CopyToClipboard GenerateImportRelatedModels
    
End Function

Public Function Generate_getModelAPIRouteNext13(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = IIf(IsNull(SlugField), "id", "slug")
    Dim GetGenericFunction: GetGenericFunction = IIf(IsNull(SlugField), "genericGetOne", "genericGetOneBySlug")
    
    Generate_getModelAPIRouteNext13 = GetReplacedTemplate(rs, "getModel API Route Next 13")
    
    Generate_getModelAPIRouteNext13 = replace(Generate_getModelAPIRouteNext13, "[GetGenericFunction]", GetGenericFunction)
    Generate_getModelAPIRouteNext13 = replace(Generate_getModelAPIRouteNext13, "[GenerateImportRelatedModels]", GenerateImportRelatedModels(frm, SeqModelID))
    Generate_getModelAPIRouteNext13 = replace(Generate_getModelAPIRouteNext13, "[SlugOrId]", SlugOrId)
    Generate_getModelAPIRouteNext13 = replace(Generate_getModelAPIRouteNext13, "[GeneratefindOptions]", GeneratefindOptions(frm, SeqModelID))
    Generate_getModelAPIRouteNext13 = replace(Generate_getModelAPIRouteNext13, "[GetUpdateFunctionWithRelationshipNext13]", GetUpdateFunctionWithRelationshipNext13(frm, SeqModelID))
    Generate_getModelAPIRouteNext13 = replace(Generate_getModelAPIRouteNext13, "[GetAllAPIRelatedLeftModelImportBySeqModel]", GetAllAPIRelatedLeftModelImportBySeqModel(frm, SeqModelID))
    Generate_getModelAPIRouteNext13 = replace(Generate_getModelAPIRouteNext13, "[GetAllAPIRelatedRightModelImportBySeqModel]", GetAllAPIRelatedRightModelImportBySeqModel(frm, SeqModelID))
    Generate_getModelAPIRouteNext13 = GetGeneratedByFunctionSnippet(Generate_getModelAPIRouteNext13, "Generate_getModelAPIRouteNext13", "getModel API Route Next 13")
    CopyToClipboard Generate_getModelAPIRouteNext13
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\app\api\" & ModelPath & "\[id]\route.ts"
    WriteToFile filePath, Generate_getModelAPIRouteNext13, SeqModelID
    
End Function

Public Function GenerateNext13APIModelRelatedBackendFiles(frm As Object, Optional SeqModelID = "")
    
    ''Next 13 API Model related backend files
    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim response: response = MsgBox("This will overwrite all the files and be replaced by the newer file if it exists" & vbNewLine & _
        "Do you want to proceed?", vbYesNo)
    
    If vbYes Then
        NoHasWriteToFilePrompt = True
        Dim sqlStr: sqlStr = "SELECT * FROM qryFunctionChainItems WHERE FunctionChainName = ""Next 13 API Model related backend files"" ORDER BY FunctionOrder ASC"
        Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        Do Until rs.EOF
            Dim FunctionName: FunctionName = rs.fields("FunctionName"): If ExitIfTrue(isFalse(FunctionName), "FunctionName is empty..") Then Exit Function
            Run FunctionName, frm, SeqModelID
            rs.MoveNext
        Loop
        NoHasWriteToFilePrompt = False
    End If
    
    MsgBox "Complete next 13 backend model files were generated.", vbOKOnly
    
End Function

Public Function CreateSequelizeModelCreateMigration(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    CreateSequelizeModelCreateMigration = GetReplacedTemplate(rs, "sequelize create model migration")
    CreateSequelizeModelCreateMigration = replace(CreateSequelizeModelCreateMigration, "[GenerateFieldsForModelMigration]", GenerateFieldsForModelMigration(frm, SeqModelID))
    CreateSequelizeModelCreateMigration = replace(CreateSequelizeModelCreateMigration, "[GetAllUniqueKeysOption]", GetAllUniqueKeysOption(frm, SeqModelID))
    CreateSequelizeModelCreateMigration = GetGeneratedByFunctionSnippet(CreateSequelizeModelCreateMigration, "CreateSequelizeModelCreateMigration", "sequelize create model migration")
    
    CopyToClipboard CreateSequelizeModelCreateMigration
    
    Dim TableName: TableName = rs.fields("tableName"): If ExitIfTrue(isFalse(TableName), "tableName is empty..") Then Exit Function
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim fileName: fileName = ConvertToCustomTimestamp & "-create_" & TableName & ".js"
    Dim filePath: filePath = ProjectPath & "src\migrations\" & fileName
    WriteToFile filePath, CreateSequelizeModelCreateMigration, SeqModelID

End Function

Public Function GenerateFieldsForModelMigration(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GenerateFieldsForModelMigration = GetModelFieldsDictionary(SeqModelID)
    GenerateFieldsForModelMigration = replace(GenerateFieldsForModelMigration, "DataTypes.", "Sequelize.")
    ''GenerateFieldsForModelMigration = GetReplacedTemplate(rs, "")
    GenerateFieldsForModelMigration = GetGeneratedByFunctionSnippet(GenerateFieldsForModelMigration, "GenerateFieldsForModelMigration")
    CopyToClipboard GenerateFieldsForModelMigration
    
End Function

Public Function GenerateMigrationForUniqueFieldCombination(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Not UniqueWith IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    
    If rs.EOF Then
        Exit Function
    End If
    
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    Dim UniqueWith: UniqueWith = rs.fields("UniqueWith"): If ExitIfTrue(isFalse(UniqueWith), "UniqueWith is empty..") Then Exit Function
    Dim uniqueFields: uniqueFields = Esc(DatabaseFieldName) & "," & Esc(UniqueWith)
    GenerateMigrationForUniqueFieldCombination = GetReplacedTemplate(rs, "Sequelize Migrate - Unique Field Combination")
    GenerateMigrationForUniqueFieldCombination = replace(GenerateMigrationForUniqueFieldCombination, "[UniqueFields]", uniqueFields)
    GenerateMigrationForUniqueFieldCombination = GetGeneratedByFunctionSnippet(GenerateMigrationForUniqueFieldCombination, "GenerateMigrationForUniqueFieldCombination")
    CopyToClipboard GenerateMigrationForUniqueFieldCombination
    
    
    Dim fileName: fileName = ConvertToCustomTimestamp & "-add_unique_" & DatabaseFieldName & "_" & UniqueWith & ".js"
    Dim filePath: filePath = ProjectPath & "src\migrations\" & fileName
    WriteToFile filePath, GenerateMigrationForUniqueFieldCombination, SeqModelID
        
End Function

Public Function GenerateTimestampMigrationFile(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GenerateTimestampMigrationFile = GetReplacedTemplate(rs, "Timestamp Migration File")
    GenerateTimestampMigrationFile = GetGeneratedByFunctionSnippet(GenerateTimestampMigrationFile, "GenerateTimestampMigrationFile")
    CopyToClipboard GenerateTimestampMigrationFile
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim fileName: fileName = ConvertToCustomTimestamp & "-add_timestamp_fields.js"
    Dim filePath: filePath = ProjectPath & "src\migrations\" & fileName
    WriteToFile filePath, GenerateTimestampMigrationFile, SeqModelID
    
End Function

Public Function GetCompleteNextAuthModelFile(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    Dim ModelFields: ModelFields = GetModelFieldsDictionary(SeqModelID, False, False)
    
    
    GetCompleteNextAuthModelFile = GetReplacedTemplate(rs, "NextAuth Model")
    GetCompleteNextAuthModelFile = replace(GetCompleteNextAuthModelFile, "[ModelFields]", ModelFields)
    GetCompleteNextAuthModelFile = replace(GetCompleteNextAuthModelFile, "[ImportRelationships]", GetRelationshipImports(frm, SeqModelID))
    GetCompleteNextAuthModelFile = replace(GetCompleteNextAuthModelFile, "[RelationshipDeclarations]", GetRelationshipDeclarations(frm, SeqModelID))
    GetCompleteNextAuthModelFile = GetGeneratedByFunctionSnippet(GetCompleteNextAuthModelFile, "GetCompleteNextAuthModelFile")
    CopyToClipboard GetCompleteNextAuthModelFile
    
    ''C:\Users\User\Desktop\Web Development\next-13-tutorial\src\models\UserModel.ts
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim filePath: filePath = ProjectPath & "src\models\" & ModelName & "Model.ts"
    WriteToFile filePath, GetCompleteNextAuthModelFile, SeqModelID
    
End Function

Private Function GetRelationshipDeclarations(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If
    
    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE DeclareInModel = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GenerateModelRelationship(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    GetRelationshipDeclarations = lines.NewLineJoin
    GetRelationshipDeclarations = GetGeneratedByFunctionSnippet(GetRelationshipDeclarations, "GetRelationshipDeclarations")
    
End Function

Private Function GetRelationshipImports(frm As Object, Optional SeqModelID = "")
    
    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE DeclareInModel = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim ModelToUse: ModelToUse = rs.fields("LeftModelID")
        If SeqModelID = ModelToUse Then
            ModelToUse = rs.fields("RightModelID")
        End If
        
        Dim ModelName: ModelName = ELookup("tblSeqModels", "SeqModelID = " & ModelToUse, "ModelName")
        ''import { Hero } from "./HeroModel";
        lines.Add "import { " & ModelName & " } from ""./" & ModelName & "Model"";"
        
        Dim ThroughModelName: ThroughModelName = rs.fields("ThroughModelName")
        If Not isFalse(ThroughModelName) Then
            lines.Add "import { " & ThroughModelName & " } from ""./" & ThroughModelName & "Model"";"
        End If
        rs.MoveNext
    Loop
    
    GetRelationshipImports = lines.NewLineJoin
    GetRelationshipImports = GetGeneratedByFunctionSnippet(GetRelationshipImports, "GetRelationshipImports")
    
End Function

Public Function RunSequelizeMigrateFromModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim lines As New clsArray
    
    lines.Add "cmd.exe /k cd /d " & Esc(ProjectPath)
    lines.Add "npx sequelize db:migrate"
    Call Shell(lines.JoinArr(" & "))
    
End Function

Public Function GetBackendModelRequiredSnippets(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetBackendModelRequiredSnippets = GetReplacedTemplate(rs, "GetBackendModelRequiredSnippets")
    GetBackendModelRequiredSnippets = GetGeneratedByFunctionSnippet(GetBackendModelRequiredSnippets, "GetBackendModelRequiredSnippets")
    CopyToClipboard GetBackendModelRequiredSnippets
    
End Function

Public Function GetModelHeaderLinkItem(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelHeaderLinkItem = GetReplacedTemplate(rs, "header link item")
    GetModelHeaderLinkItem = GetGeneratedByFunctionSnippet(GetModelHeaderLinkItem, "GetModelHeaderLinkItem", "header link item")
    CopyToClipboard GetModelHeaderLinkItem
    
End Function


Public Function GetAllFilterInterfaceBySeqmodel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If isPresent("qrySeqModelFilters", "SeqModelID = " & SeqModelID & " AND FilterQueryName = ""q""") Then
        lines.Add "q: string"
    End If
    
    sqlStr = "SELECT SeqModelFilterID FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterQueryName <> ""q"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetThisFilterInterface(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    GetAllFilterInterfaceBySeqmodel = lines.JoinArr(vbNewLine)
    GetAllFilterInterfaceBySeqmodel = GetGeneratedByFunctionSnippet(GetAllFilterInterfaceBySeqmodel, "GetAllFilterInterfaceBySeqmodel")
    CopyToClipboard GetAllFilterInterfaceBySeqmodel
    
End Function

Public Function WriteToModelinterface_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath")
    If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    WriteToModelinterface_ts = GetReplacedTemplate(rs, "ModelInterface.ts Next 13")
    
    Dim isDropZone: isDropZone = isPresent("tblSeqModelRelationships", "LeftModelID = " & SeqModelID & " AND IncludeAsDropzone")
    ''Dim GetModelFilesUpdatePayload: GetModelFilesUpdatePayload = IIf(isDropZone, GetReplacedTemplate(rs, "GetModelFilesUpdatePayload"), "")
    ''Dim GetModelFilesFormUpdatePayload: GetModelFilesFormUpdatePayload = IIf(isDropZone, GetReplacedTemplate(rs, "GetModelFilesFormUpdatePayload"), "")
    
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllFilterInterfaceBySeqmodel]", GetAllFilterInterfaceBySeqmodel(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllModelFieldTypeBySeqModel]", GetAllModelFieldTypeBySeqModel(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetSlugType]", GetSlugType(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetTimestampTypes]", GetTimestampTypes(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllNonStringFilterNames]", GetAllNonStringFilterNames(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllNonStringFilterTypes]", GetAllNonStringFilterTypes(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllNonStringFieldNamesBySeqModel]", GetAllNonStringFieldNamesBySeqModel(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllNonStringFieldTypesBySeqModel]", GetAllNonStringFieldTypesBySeqModel(frm, SeqModelID))
    ''WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllRelatedUpdatePayloadBySeqModel]", GetAllRelatedUpdatePayloadBySeqModel(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllRelatedInterfaceImportBySeqModel]", GetAllRelatedInterfaceImportBySeqModel(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllChildModelInterfaceBySeqModel]", GetAllChildModelInterfaceBySeqModel(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllParentModelInterfaceBySeqModel]", GetAllParentModelInterfaceBySeqModel(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllRelatedModelNameBySeqModel]", GetAllRelatedModelNameBySeqModel(frm, SeqModelID))
    ''WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllOmittedModelFormKeys]", GetAllOmittedModelFormKeys(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllRelatedModelUpdatePayload]", GetAllRelatedModelUpdatePayload(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllRelatedFormikInitialValues]", GetAllRelatedFormikInitialValues(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetModelUpdatePayloadExtension]", GetModelUpdatePayloadExtension(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllRelatedRightModelImport]", GetAllRelatedRightModelImport(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllRelatedRightModelInterface]", GetAllRelatedRightModelInterface(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllSimpleRelatedKey]", GetAllSimpleRelatedKey(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllSimplePluralizedFieldName]", GetAllSimplePluralizedFieldName(frm, SeqModelID))
    WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetAllSimpleRelatedKeyPayload]", GetAllSimpleRelatedKeyPayload(frm, SeqModelID))
    ''WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetModelFilesUpdatePayload]", GetModelFilesUpdatePayload)
    ''WriteToModelinterface_ts = replace(WriteToModelinterface_ts, "[GetModelFilesFormUpdatePayload]", GetModelFilesFormUpdatePayload)
    WriteToModelinterface_ts = GetGeneratedByFunctionSnippet(WriteToModelinterface_ts, "WriteToModelinterface_ts", "ModelInterface.ts Next 13")
    
    CopyToClipboard WriteToModelinterface_ts
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\interfaces\DeckInterfaces.ts
    Dim filePath: filePath = ClientPath & "src\interfaces\" & ModelName & "Interfaces.ts"
    WriteToFile filePath, WriteToModelinterface_ts, SeqModelID, "WriteToModelinterface_ts"
    
End Function

Public Function GetAllModelFieldTypeBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim Timestamps: Timestamps = rs.fields("Timestamps")
    
'    Dim ExceptedFieldsArr As New clsArray
'    ExceptedFieldsArr.arr = "slug,createdAt,updatedAt"
'    ExceptedFieldsArr.EscapeItems
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND ((ViewName IS NULL AND Not IsGeneratedField)" & _
        " OR (Not ViewName IS NULL))"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetModelFieldType(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If Not isFalse(SlugField) Then
        lines.Add "slug: string;"
    End If
    
    If Timestamps Then
        lines.Add "createdAt: string;"
        lines.Add "updatedAt: string;"
    End If
    
    GetAllModelFieldTypeBySeqModel = lines.JoinArr(vbNewLine)
    GetAllModelFieldTypeBySeqModel = GetGeneratedByFunctionSnippet(GetAllModelFieldTypeBySeqModel, "GetAllModelFieldTypeBySeqModel")
    CopyToClipboard GetAllModelFieldTypeBySeqModel
    
End Function

Public Function GetSlugType(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    
    GetSlugType = ""
    If Not isFalse(SlugField) Then
        GetSlugType = "slug : string"
    End If
    
End Function

Public Function GetTimestampTypes(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim Timestamps: Timestamps = rs.fields("Timestamps")
    
    If Timestamps Then
        GetTimestampTypes = "createdAt: string;updatedAt: string;"
    End If

End Function

Public Function WriteToModelconstants_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
'    GetUniqueFields
'    GetRequiredFields

    WriteToModelconstants_ts = GetReplacedTemplate(rs, "ModelConstants.ts")
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetAllModelFilterDefaultBySeqModel]", GetAllModelFilterDefaultBySeqModel(frm, SeqModelID))
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetFirstFieldInForm]", GetFirstFieldInForm(frm, SeqModelID))
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetLastFieldInForm]", GetLastFieldInForm(frm, SeqModelID))
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetDefaultQFilter]", GetDefaultQFilter(frm, SeqModelID))
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetAllFormDefaultValueBySeqModel]", GetAllFormDefaultValueBySeqModel(frm, SeqModelID))
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetModelPrimaryKey]", GetModelPrimaryKey(frm, SeqModelID))
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetAllUniqueFieldsBySeqModel]", GetAllUniqueFieldsBySeqModel(frm, SeqModelID))
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetAllRequiredFieldsBySeqModel]", GetAllRequiredFieldsBySeqModel(frm, SeqModelID))
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetControlOptionsBySeqModel]", GetControlOptionsBySeqModel(frm, SeqModelID))
    WriteToModelconstants_ts = replace(WriteToModelconstants_ts, "[GetCOLUMNSObject]", GetCOLUMNSObject(frm, SeqModelID))
    WriteToModelconstants_ts = GetGeneratedByFunctionSnippet(WriteToModelconstants_ts, "WriteToModelconstants_ts", "ModelConstants.ts")
    CopyToClipboard WriteToModelconstants_ts
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\utils\constants\DeckConstants.ts
    Dim filePath: filePath = ClientPath & "src\utils\constants\" & ModelName & "Constants.ts"
    WriteToFile filePath, WriteToModelconstants_ts, SeqModelID
        
End Function

Public Function GetAllModelFilterDefaultBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If isPresent("qrySeqModelFilters", "SeqModelID = " & SeqModelID & " AND FilterQueryName = ""q""") Then
        lines.Add "q: """","
    End If
    
    sqlStr = "SElECT SeqModelFilterID FROm qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterQueryName <> ""q"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetModelFilterDefault(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    GetAllModelFilterDefaultBySeqModel = lines.JoinArr(vbNewLine)
    GetAllModelFilterDefaultBySeqModel = GetGeneratedByFunctionSnippet(GetAllModelFilterDefaultBySeqModel, "GetAllModelFilterDefaultBySeqModel")
    CopyToClipboard GetAllModelFilterDefaultBySeqModel
    
End Function

Public Function GetDefaultQFilter(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterQueryName = " & Esc("q")
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If Not rs.EOF Then
        GetDefaultQFilter = "q: """""
    End If
    
    GetDefaultQFilter = GetGeneratedByFunctionSnippet(GetDefaultQFilter, "GetDefaultQFilter")
    CopyToClipboard GetDefaultQFilter
    
End Function

Public Function GetFirstFieldInForm(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey ORDER BY FieldOrder"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    
    GetFirstFieldInForm = fieldName
    CopyToClipboard GetFirstFieldInForm
    
End Function

Public Function GetLastFieldInForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey And ControlType <> ""Hidden"" ORDER BY FieldOrder DESC"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    
    GetLastFieldInForm = fieldName
    CopyToClipboard GetLastFieldInForm
    
End Function

Public Function WriteToModelcolumns_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim IsTable: IsTable = rs.fields("IsTable")
    
    Dim GetFormikShapeOrModel: GetFormikShapeOrModel = IIf(IsTable, "Model", "FormikShape")
    ''GetAllTableFieldCellInputBySeqModel
    
    WriteToModelcolumns_tsx = GetReplacedTemplate(rs, "ModelColumns.tsx")
    WriteToModelcolumns_tsx = replace(WriteToModelcolumns_tsx, "[GetAllTableFieldCellInputBySeqModel]", GetAllTableFieldCellInputBySeqModel(frm, SeqModelID))
    WriteToModelcolumns_tsx = replace(WriteToModelcolumns_tsx, "[GetControlOptionsImportLine]", GetControlOptionsImportLine(frm, SeqModelID))
    WriteToModelcolumns_tsx = replace(WriteToModelcolumns_tsx, "[GetActionCell]", GetActionCell(frm, SeqModelID))
    WriteToModelcolumns_tsx = replace(WriteToModelcolumns_tsx, "[GetModelRowActionsImport]", GetModelRowActionsImport(frm, SeqModelID))
    WriteToModelcolumns_tsx = replace(WriteToModelcolumns_tsx, "[GetFormikShapeOrModel]", GetFormikShapeOrModel)
    WriteToModelcolumns_tsx = replace(WriteToModelcolumns_tsx, "[GetAllRightModelListImportForColumn]", GetAllRightModelListImportForColumn(frm, SeqModelID))
    WriteToModelcolumns_tsx = replace(WriteToModelcolumns_tsx, "[GetAllUseRightModelListForColumn]", GetAllUseRightModelListForColumn(frm, SeqModelID))
    WriteToModelcolumns_tsx = GetGeneratedByFunctionSnippet(WriteToModelcolumns_tsx, "WriteToModelcolumns_tsx", "ModelColumns.tsx")
    CopyToClipboard WriteToModelcolumns_tsx
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\components\deck\DeckColumns.tsx
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "Columns.tsx"
    WriteToFile filePath, WriteToModelcolumns_tsx, SeqModelID

End Function

Public Function GetAllTableFieldCellInputBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey AND ControlType <> ""Hidden"" ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetTableFieldCellInput(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GetAllTableFieldCellInputBySeqModel = lines.JoinArr(",")
    GetAllTableFieldCellInputBySeqModel = GetGeneratedByFunctionSnippet(GetAllTableFieldCellInputBySeqModel, "GetAllTableFieldCellInputBySeqModel")
    CopyToClipboard GetAllTableFieldCellInputBySeqModel
    
End Function

Public Function WriteToModeldeletedialog_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim IsTable: IsTable = rs.fields("IsTable")
    
    Dim GetFormikArgument: GetFormikArgument = IIf(IsTable, "", "{ formik }: { formik: any }")

    Dim GetOnSuccessFormikDeleteDialog: GetOnSuccessFormikDeleteDialog = IIf(IsTable, GetReplacedTemplate(rs, "GetOnSuccessFormikDeleteDialogTable"), GetReplacedTemplate(rs, "GetOnSuccessFormikDeleteDialog"))
    
    WriteToModeldeletedialog_tsx = GetReplacedTemplate(rs, "ModelDeleteDialog.tsx")
    WriteToModeldeletedialog_tsx = replace(WriteToModeldeletedialog_tsx, "[GetFormikArgument]", GetFormikArgument)
    WriteToModeldeletedialog_tsx = replace(WriteToModeldeletedialog_tsx, "[GetOnSuccessFormikDeleteDialog]", GetOnSuccessFormikDeleteDialog)
    WriteToModeldeletedialog_tsx = GetGeneratedByFunctionSnippet(WriteToModeldeletedialog_tsx, "WriteToModeldeletedialog_tsx", "ModelDeleteDialog.tsx")
    CopyToClipboard WriteToModeldeletedialog_tsx
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\components\decks\DeckDeleteDialog.tsx
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "DeleteDialog.tsx"
    WriteToFile filePath, WriteToModeldeletedialog_tsx, SeqModelID

    
End Function

Public Function WriteToModelfilterform_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    ''GetAllFormikFilterControlBySeqModel
    ''GetFormikFilterQControl
    WriteToModelfilterform_tsx = GetReplacedTemplate(rs, "ModelFilterForm.tsx")
    ''WriteToModelfilterform_tsx = replace(WriteToModelfilterform_tsx, "[GetFormikFilterQControl]", GetFormikFilterQControl(frm, SeqModelID))
    ''WriteToModelfilterform_tsx = replace(WriteToModelfilterform_tsx, "[GetAllFormikFilterControlBySeqModel]", GetAllFormikFilterControlBySeqModel(frm, SeqModelID))
    WriteToModelfilterform_tsx = replace(WriteToModelfilterform_tsx, "[GetRequiredQueryFromTanstackBySeqModel]", GetRequiredQueryFromTanstackBySeqModel(frm, SeqModelID))
    WriteToModelfilterform_tsx = replace(WriteToModelfilterform_tsx, "[ImportControCONTROL_OPTIONS]", ImportControCONTROL_OPTIONS(frm, SeqModelID))
    ''WriteToModelfilterform_tsx = replace(WriteToModelfilterform_tsx, "[ImportAllUseModelListHookBySeqModel]", ImportAllUseModelListHookBySeqModel(frm, SeqModelID))
    WriteToModelfilterform_tsx = GetGeneratedByFunctionSnippet(WriteToModelfilterform_tsx, "WriteToModelfilterform_tsx", "ModelFilterForm.tsx")
    CopyToClipboard WriteToModelfilterform_tsx
    
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "FilterForm.tsx"
    WriteToFile filePath, WriteToModelfilterform_tsx, SeqModelID, "WriteToModelfilterform_tsx"
    
End Function

Public Function GetFormikFilterQControl(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If isPresent("qrySeqModelFilters", "SeqModelID = " & SeqModelID & " AND FilterQueryName = ""q""") Then
        GetFormikFilterQControl = GetReplacedTemplate(rs, "GetFormikFilterQControl")
        GetFormikFilterQControl = GetGeneratedByFunctionSnippet(GetFormikFilterQControl, "GetFormikFilterQControl", "GetFormikFilterQControl", True)
    End If
    
    CopyToClipboard GetFormikFilterQControl
    
End Function

Public Function WriteToModelformarray_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToModelformarray_tsx = GetReplacedTemplate(rs, "ModelFormArray.tsx")
    WriteToModelformarray_tsx = replace(WriteToModelformarray_tsx, "[GetAllRequiredListForTableForm]", GetAllRequiredListForTableForm(frm, SeqModelID))
    WriteToModelformarray_tsx = replace(WriteToModelformarray_tsx, "[GetAllRightModelListImportForColumn]", GetAllRightModelListImportForColumn(frm, SeqModelID))
    WriteToModelformarray_tsx = replace(WriteToModelformarray_tsx, "[GetAllUseRightModelListForColumn]", GetAllUseRightModelListForColumn(frm, SeqModelID))
    WriteToModelformarray_tsx = GetGeneratedByFunctionSnippet(WriteToModelformarray_tsx, "WriteToModelformarray_tsx", "ModelFormArray.tsx")
    CopyToClipboard WriteToModelformarray_tsx
    
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "FormArray.tsx"
    WriteToFile filePath, WriteToModelformarray_tsx, SeqModelID
    
End Function

Public Function GetAllFormDefaultValueBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetFormDefaultValue(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GetAllFormDefaultValueBySeqModel = lines.JoinArr(vbNewLine)
    GetAllFormDefaultValueBySeqModel = GetGeneratedByFunctionSnippet(GetAllFormDefaultValueBySeqModel, "GetAllFormDefaultValueBySeqModel")
    CopyToClipboard GetAllFormDefaultValueBySeqModel
    
End Function

Public Function WriteToModeltable_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim IsTable: IsTable = rs.fields("IsTable")
    
    ''Dim TemplateName As String: TemplateName = IIf(IsTable, "ModelTable.tsx for Table 9/18", "ModelTable.tsx")
    Dim templateName As String: templateName = "ModelTable.tsx for Table 9/18"
    
    WriteToModeltable_tsx = GetReplacedTemplate(rs, templateName)
    'WriteToModeltable_tsx = replace(WriteToModeltable_tsx, "[GetAllSearchParamsBySeqModel]", GetAllSearchParamsBySeqModel(frm, SeqModelID))
    'WriteToModeltable_tsx = replace(WriteToModeltable_tsx, "[GetAllFilterQueryNameBySeqModel]", GetAllFilterQueryNameBySeqModel(frm, SeqModelID))
    'WriteToModeltable_tsx = replace(WriteToModeltable_tsx, "[GetMutationSnippets]", GetMutationSnippets(frm, SeqModelID))
    WriteToModeltable_tsx = GetGeneratedByFunctionSnippet(WriteToModeltable_tsx, "WriteToModeltable_tsx", templateName)
    CopyToClipboard WriteToModeltable_tsx
    
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "Table.tsx"
    WriteToFile filePath, WriteToModeltable_tsx, SeqModelID, "WriteToModeltable_tsx"
    
End Function

Public Function WriteToUsemodelstore_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim IsTable: IsTable = rs.fields("IsTable")
    Dim GetFormikShapeOrModel: GetFormikShapeOrModel = IIf(IsTable, "Model", "FormikShape")
    
    WriteToUsemodelstore_ts = GetReplacedTemplate(rs, "useModelStore.ts")
    WriteToUsemodelstore_ts = replace(WriteToUsemodelstore_ts, "[GetFormikShapeOrModel]", GetFormikShapeOrModel)
    WriteToUsemodelstore_ts = GetGeneratedByFunctionSnippet(WriteToUsemodelstore_ts, "WriteToUsemodelstore_ts", "useModelStore.ts")
    CopyToClipboard WriteToUsemodelstore_ts
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\hooks\deck\useDeckStore.ts
    Dim filePath: filePath = ClientPath & "src\hooks\" & ModelPath & "\use" & ModelName & "Store.tsx"
    WriteToFile filePath, WriteToUsemodelstore_ts, SeqModelID
    
End Function

Public Function WriteToUsemodeldeletedialog_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToUsemodeldeletedialog_ts = GetReplacedTemplate(rs, "useModelDeleteDialog.ts")
    WriteToUsemodeldeletedialog_ts = GetGeneratedByFunctionSnippet(WriteToUsemodeldeletedialog_ts, "WriteToUsemodeldeletedialog_ts", "useModelDeleteDialog.ts")
    CopyToClipboard WriteToUsemodeldeletedialog_ts
    
    Dim filePath: filePath = ClientPath & "src\hooks\" & ModelPath & "\use" & ModelName & "DeleteDialog.tsx"
    WriteToFile filePath, WriteToUsemodeldeletedialog_ts, SeqModelID
End Function

Public Function WriteToModelschema_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
'    GetAllFieldValidationBySeqModel
'    GetAllArrayFieldValidationBySeqModel

    WriteToModelschema_ts = GetReplacedTemplate(rs, "ModelSchema.ts")
    WriteToModelschema_ts = replace(WriteToModelschema_ts, "[GetAllFieldValidationBySeqModel]", GetAllFieldValidationBySeqModel(frm, SeqModelID))
    WriteToModelschema_ts = replace(WriteToModelschema_ts, "[GetAllArrayFieldValidationBySeqModel]", GetAllArrayFieldValidationBySeqModel(frm, SeqModelID))
    WriteToModelschema_ts = replace(WriteToModelschema_ts, "[GetAllRelatedLeftArrayValidation]", GetAllRelatedLeftArrayValidation(frm, SeqModelID))
    WriteToModelschema_ts = GetGeneratedByFunctionSnippet(WriteToModelschema_ts, "WriteToModelschema_ts", "ModelSchema.ts")
    CopyToClipboard WriteToModelschema_ts
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\schema\DeckSchema.ts
    Dim filePath: filePath = ClientPath & "src\schema\" & ModelName & "Schema.ts"
    WriteToFile filePath, WriteToModelschema_ts, SeqModelID
    
End Function

Public Function GetAllFieldValidationBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add CreateFieldSchema(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GetAllFieldValidationBySeqModel = lines.JoinArr(vbNewLine)
    GetAllFieldValidationBySeqModel = GetGeneratedByFunctionSnippet(GetAllFieldValidationBySeqModel, "GetAllFieldValidationBySeqModel")
    CopyToClipboard GetAllFieldValidationBySeqModel
    
End Function

Public Function GetAllArrayFieldValidationBySeqModel(frm As Object, Optional SeqModelID = "", Optional ExludeField = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey"
    If Not isFalse(ExludeField) Then
        sqlStr = sqlStr & " AND DatabaseFieldName <> " & Esc(ExludeField)
    End If
    
    sqlStr = sqlStr & " ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add CreateFieldSchema(frm, SeqModelFieldID, True)
        rs.MoveNext
    Loop
    
    GetAllArrayFieldValidationBySeqModel = lines.JoinArr(",")
    GetAllArrayFieldValidationBySeqModel = GetGeneratedByFunctionSnippet(GetAllArrayFieldValidationBySeqModel, "GetAllArrayFieldValidationBySeqModel")
    CopyToClipboard GetAllArrayFieldValidationBySeqModel
    
End Function

Public Function GetModelPrimaryKey(frm As Object, Optional SeqModelID = "", Optional UseFieldName As Boolean = False)

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GetModelPrimaryKey = ELookup("tblSeqModelFields", "PrimaryKey AND SeqModelID = " & SeqModelID, IIf(UseFieldName, "FieldName", "DatabaseFieldName"))
    
End Function

Public Function GetAllQFilterFieldBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFilterID FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterQueryName = ""q"" AND NOT SeqModelFieldID IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GenerateAQFilterField(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    GetAllQFilterFieldBySeqModel = lines.JoinArr(",")
    GetAllQFilterFieldBySeqModel = GetGeneratedByFunctionSnippet(GetAllQFilterFieldBySeqModel, "GetAllQFilterFieldBySeqModel")
    CopyToClipboard GetAllQFilterFieldBySeqModel
    
End Function

Public Function GetAllModelAttributesBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim Timestamps: Timestamps = rs.fields("Timestamps")
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT IsExpression"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetModelAttribute(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If Not isFalse(SlugField) Then lines.Add Esc("slug")
    If Timestamps Then
        lines.Add Esc("createdAt")
        lines.Add Esc("updatedAt")
    End If
    
    GetAllModelAttributesBySeqModel = lines.JoinArr(",")
    GetAllModelAttributesBySeqModel = GetGeneratedByFunctionSnippet(GetAllModelAttributesBySeqModel, "GetAllModelAttributesBySeqModel")
    CopyToClipboard GetAllModelAttributesBySeqModel
    
End Function

Public Function GetAllFieldsToUpdateBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey AND NOT isExpression"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetFieldsToUpdate(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GetAllFieldsToUpdateBySeqModel = lines.JoinArr(",")
    GetAllFieldsToUpdateBySeqModel = GetGeneratedByFunctionSnippet(GetAllFieldsToUpdateBySeqModel, "GetAllFieldsToUpdateBySeqModel")
    CopyToClipboard GetAllFieldsToUpdateBySeqModel
    
End Function

Public Function WriteToModelsRouteApi(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim IsMainQuery: IsMainQuery = rs.fields("IsMainQuery")
    Dim IsTable: IsTable = rs.fields("IsTable")
    Dim templateName As String: templateName = IIf(IsMainQuery, "models route next 13 with SQL", "models route next 13")
    
    WriteToModelsRouteApi = GetReplacedTemplate(rs, templateName)
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetAllModelAttributesBySeqModel]", GetAllModelAttributesBySeqModel(frm, SeqModelID))
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetAllQFilterFieldBySeqModel]", GetAllQFilterFieldBySeqModel(frm, SeqModelID))
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetAllFieldsToUpdateBySeqModel]", GetAllFieldsToUpdateBySeqModel(frm, SeqModelID))
    ''WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GenerateImportRelatedModels]", GenerateImportRelatedModels(frm, SeqModelID))
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetAllBackendFiltersBySeqModel]", GetAllBackendFiltersBySeqModel(frm, SeqModelID))
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetGetmodelsqlNext13]", GetGetmodelsqlNext13(frm, SeqModelID))
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetSqlModelsGetRoute]", GetSqlModelsGetRoute(frm, SeqModelID))
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetAllGetmodelsqlChildNext13]", GetAllGetmodelsqlChildNext13(frm, SeqModelID))
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetAllGetmodelsqlLeftModelChildNext13]", GetAllGetmodelsqlLeftModelChildNext13(frm, SeqModelID))
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetAllCreateSimpleModelFromRoute]", GetAllCreateSimpleModelFromRoute(frm, SeqModelID))
    
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetMultiCreateModelPOSTRoute]", IIf(IsTable, "", GetMultiCreateModelPOSTRoute(frm, SeqModelID)))
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetSingleCreateModelPOSTRoute]", IIf(Not IsTable, "", GetSingleCreateModelPOSTRoute(frm, SeqModelID)))
    
    WriteToModelsRouteApi = replace(WriteToModelsRouteApi, "[GetAllRelatedLeftModelImportRoute]", GetAllRelatedLeftModelImportRoute(frm, SeqModelID))
    
    WriteToModelsRouteApi = GetGeneratedByFunctionSnippet(WriteToModelsRouteApi, "WriteToModelsRouteApi", templateName)
    
    CopyToClipboard WriteToModelsRouteApi
        
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\app\api\decks\route.ts
    Dim filePath: filePath = ClientPath & "src\app\api\" & ModelPath & "\route.ts"
    WriteToFile filePath, WriteToModelsRouteApi, SeqModelID
    
End Function

Public Function WriteToModelsPage(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim APIOnly: APIOnly = rs.fields("APIOnly")
    
    If APIOnly Then Exit Function
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim SidebarEnabled: SidebarEnabled = rs.fields("SidebarEnabled")
    Dim templateName: templateName = IIf(SidebarEnabled, "Model Page Sidebar", "Model Page")
    Dim ContainerWidth: ContainerWidth = rs.fields("ContainerWidth"): ContainerWidth = IIf(isFalse(ContainerWidth), "", ContainerWidth)
    
    WriteToModelsPage = GetReplacedTemplate(rs, templateName)
    WriteToModelsPage = replace(WriteToModelsPage, "[ContainerWidthStr]", ContainerWidth)
    WriteToModelsPage = GetGeneratedByFunctionSnippet(WriteToModelsPage, "WriteToModelsPage", templateName)
    CopyToClipboard WriteToModelsPage
    
    ''C:\Users\User\Desktop\Web Development\panda-realty\src\app\(protected)\buyer-status\page.tsx
    Dim filePath: filePath = ClientPath & "src\app\(protected)\" & ModelPath & "\page.tsx"
    WriteToFile filePath, WriteToModelsPage, SeqModelID, "WriteToModelsPage"
    
End Function

Public Function GetAllUniqueFieldsBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Unique"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetUniqueField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GetAllUniqueFieldsBySeqModel = lines.JoinArr(",")
    GetAllUniqueFieldsBySeqModel = GetGeneratedByFunctionSnippet(GetAllUniqueFieldsBySeqModel, "GetAllUniqueFieldsBySeqModel")
    CopyToClipboard GetAllUniqueFieldsBySeqModel
    
End Function

Public Function GetAllRequiredFieldsBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Not AllowNull AND NOT PrimaryKey AND Not isExpression" & _
        " AND DataTypeInterface <> " & Esc("boolean")
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetRequiredField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GetAllRequiredFieldsBySeqModel = lines.JoinArr(vbNewLine)
    GetAllRequiredFieldsBySeqModel = GetGeneratedByFunctionSnippet(GetAllRequiredFieldsBySeqModel, "GetAllRequiredFieldsBySeqModel")
    CopyToClipboard GetAllRequiredFieldsBySeqModel
    
End Function

Public Function GetGetmodelsqlNext13(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

'GenerateSQLFieldList
'GenerateSeqModelFilters
    GetGetmodelsqlNext13 = GetReplacedTemplate(rs, "getModelSQL Next 13")
    GetGetmodelsqlNext13 = replace(GetGetmodelsqlNext13, "[GenerateSQLFieldList]", GenerateSQLFieldList(frm, SeqModelID))
    GetGetmodelsqlNext13 = replace(GetGetmodelsqlNext13, "[GenerateSeqModelFilters]", GenerateSeqModelFilters(frm, SeqModelID))
    GetGetmodelsqlNext13 = replace(GetGetmodelsqlNext13, "[GetAllSQLRightJoinSnippets]", GetAllSQLRightJoinSnippets(frm, SeqModelID))
    GetGetmodelsqlNext13 = replace(GetGetmodelsqlNext13, "[GetAllRightModelJoinCancellationSnippet]", GetAllRightModelJoinCancellationSnippet(frm, SeqModelID))
    GetGetmodelsqlNext13 = replace(GetGetmodelsqlNext13, "[GetAllRightJoinName]", GetAllRightJoinName(frm, SeqModelID))
    GetGetmodelsqlNext13 = replace(GetGetmodelsqlNext13, "[GetAllSQLLeftJoinSnippets]", GetAllSQLLeftJoinSnippets(frm, SeqModelID))
    GetGetmodelsqlNext13 = replace(GetGetmodelsqlNext13, "[GetAllLeftModelJoinCancellationSnippet]", GetAllLeftModelJoinCancellationSnippet(frm, SeqModelID))
    GetGetmodelsqlNext13 = replace(GetGetmodelsqlNext13, "[GetAllLeftJoinName]", GetAllLeftJoinName(frm, SeqModelID))
    GetGetmodelsqlNext13 = GetGeneratedByFunctionSnippet(GetGetmodelsqlNext13, "GetGetmodelsqlNext13", "getModelSQL Next 13")
    CopyToClipboard GetGetmodelsqlNext13
    
End Function

Public Function GetSqlModelsGetRoute(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSqlModelsGetRoute = GetReplacedTemplate(rs, "GET Models route")
    GetSqlModelsGetRoute = replace(GetSqlModelsGetRoute, "[GetAllLeftModelReduceResultAndRemoveDuplicates]", GetAllLeftModelReduceResultAndRemoveDuplicates(frm, SeqModelID))
    GetSqlModelsGetRoute = replace(GetSqlModelsGetRoute, "[GetAllLeftModelsToReduce]", GetAllLeftModelsToReduce(frm, SeqModelID))
    GetSqlModelsGetRoute = GetGeneratedByFunctionSnippet(GetSqlModelsGetRoute, "GetSqlModelsGetRoute", "GET Models route")
    CopyToClipboard GetSqlModelsGetRoute
    
    
End Function

Public Function GetGetmodelsqlChildNext13(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetGetmodelsqlChildNext13 = GetReplacedTemplate(rs, "getModelSQL child next 13")
    GetGetmodelsqlChildNext13 = replace(GetGetmodelsqlChildNext13, "[GenerateSQLFieldList]", GenerateSQLFieldList(frm, SeqModelID))
    GetGetmodelsqlChildNext13 = replace(GetGetmodelsqlChildNext13, "[GenerateSeqModelFilters]", GenerateSeqModelFilters(frm, SeqModelID))
    GetGetmodelsqlChildNext13 = GetGeneratedByFunctionSnippet(GetGetmodelsqlChildNext13, "GetGetmodelsqlChildNext13", "getModelSQL child next 13")
    CopyToClipboard GetGetmodelsqlChildNext13
    
End Function

Public Function GetAllNonStringFilterNames(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        
    Dim nonStringControls As New clsArray: nonStringControls.arr = "Switch,FacetedControl,CheckBoxGroup"
    nonStringControls.EscapeItems
    
    sqlStr = "SELECT SeqModelFilterID FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND ControlType In(" & nonStringControls.JoinArr & ")"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        Dim FilterQueryName: FilterQueryName = GetFilterQueryName(frm, SeqModelFilterID)
        lines.Add Left(FilterQueryName, Len(FilterQueryName) - 34)
        rs.MoveNext
    Loop
    If lines.count > 0 Then
        lines.EscapeItems
        GetAllNonStringFilterNames = lines.JoinArr(" | ")
        GetAllNonStringFilterNames = GetGeneratedByFunctionSnippet(GetAllNonStringFilterNames, "GetAllNonStringFilterNames")
    Else
        GetAllNonStringFilterNames = Esc("")
    End If
    CopyToClipboard GetAllNonStringFilterNames
    
End Function

Public Function GetAllNonStringFilterTypes(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim nonStringControls As New clsArray: nonStringControls.arr = "Switch,FacetedControl,CheckBoxGroup"
    nonStringControls.EscapeItems
    sqlStr = "SElECT SeqModelFilterID FROm qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND ControlType In(" & nonStringControls.JoinArr & ")"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetThisFilterInterface(frm, SeqModelFilterID, True)
        rs.MoveNext
    Loop
    
    GetAllNonStringFilterTypes = lines.JoinArr(vbNewLine)
    GetAllNonStringFilterTypes = GetGeneratedByFunctionSnippet(GetAllNonStringFilterTypes, "GetAllNonStringFilterTypes")
    CopyToClipboard GetAllNonStringFilterTypes
    
End Function

Public Function GetAllSimpleJoinFieldsBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Not isExpression"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetSimpleJoinFields(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GetAllSimpleJoinFieldsBySeqModel = lines.JoinArr(",")
    GetAllSimpleJoinFieldsBySeqModel = GetGeneratedByFunctionSnippet(GetAllSimpleJoinFieldsBySeqModel, "GetAllSimpleJoinFieldsBySeqModel")
    CopyToClipboard GetAllSimpleJoinFieldsBySeqModel
    
End Function

Public Function GetAllFormikFilterControlBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SElECT SeqModelFilterID FROm qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterQueryName <> ""q"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetFormikFilterControl(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    GetAllFormikFilterControlBySeqModel = lines.JoinArr(vbNewLine)
    GetAllFormikFilterControlBySeqModel = "{/* Generated by GetAllFormikFilterControlBySeqModel */}" & GetAllFormikFilterControlBySeqModel
    CopyToClipboard GetAllFormikFilterControlBySeqModel
    
End Function

Public Function GetAllSearchParamsBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If isPresent("qrySeqModelFilters", "FilterQueryName = ""q"" AND SeqModelID = " & SeqModelID) Then
        lines.Add "const q = query[""q""] || """""
    End If
    
    sqlStr = "SElECT SeqModelFilterID FROm qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterQueryName <> ""q"" ORDER BY SeqModelFilterID"
    
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetSearchParamVariable(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    GetAllSearchParamsBySeqModel = lines.JoinArr(vbNewLine)
    GetAllSearchParamsBySeqModel = GetGeneratedByFunctionSnippet(GetAllSearchParamsBySeqModel, "GetAllSearchParamsBySeqModel")
    CopyToClipboard GetAllSearchParamsBySeqModel
    
End Function

Public Function GetAllFilterQueryNameBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If isPresent("qrySeqModelFilters", "SeqModelID = " & SeqModelID & " And FilterQueryName = ""q""") Then
        lines.Add "q,"
    End If
    
    sqlStr = "SElECT SeqModelFilterID FROm qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " And FilterQueryName <> ""q"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetFilterQueryName(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    GetAllFilterQueryNameBySeqModel = lines.JoinArr(vbNewLine)
    GetAllFilterQueryNameBySeqModel = GetGeneratedByFunctionSnippet(GetAllFilterQueryNameBySeqModel, "GetAllFilterQueryNameBySeqModel")
    CopyToClipboard GetAllFilterQueryNameBySeqModel
    
End Function

Public Function GetAllNonStringFieldNamesBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND DataTypeInterface = ""number"" AND Not PrimaryKey"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetModelFieldName(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        lines.EscapeItems
        GetAllNonStringFieldNamesBySeqModel = lines.JoinArr(" | ")
        GetAllNonStringFieldNamesBySeqModel = GetGeneratedByFunctionSnippet(GetAllNonStringFieldNamesBySeqModel, "GetAllNonStringFieldNamesBySeqModel")
        CopyToClipboard GetAllNonStringFieldNamesBySeqModel
    Else
        GetAllNonStringFieldNamesBySeqModel = Esc("")
    End If
    
End Function

Public Function GetAllNonStringFieldTypesBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND DataTypeInterface = ""number"" AND NOT PrimaryKey"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetModelFieldType(frm, SeqModelFieldID, True)
        rs.MoveNext
    Loop
    
    GetAllNonStringFieldTypesBySeqModel = lines.JoinArr(";")
    GetAllNonStringFieldTypesBySeqModel = GetGeneratedByFunctionSnippet(GetAllNonStringFieldTypesBySeqModel, "GetAllNonStringFieldTypesBySeqModel")
    CopyToClipboard GetAllNonStringFieldTypesBySeqModel
    
End Function

Public Function GetControlOptionsBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    sqlStr = "SELECT FieldName, SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND (DataType = ""ENUM"" Or " & _
        "Not AllowedOptions IS NULL) ORDER BY FieldOrder"
    
    Dim uqFieldNames As New clsArray
    
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        Dim fieldName: fieldName = rs.fields("FieldName")
        uqFieldNames.Add fieldName
        lines.Add GetFieldOptions(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    sqlStr = "SELECT SeqModelFilterID, FilterQueryName FROM qrySeqModelFilterOptions GROUP BY SeqModelFilterID, FilterQueryName ORDER BY SeqModelFilterID "
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID"): If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
        Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
        If Not uqFieldNames.InArray(FilterQueryName) Then
            uqFieldNames.Add FilterQueryName
            lines.Add GetAllFilterManualOption(frm, SeqModelFilterID)
        Else
            MsgBox "Filter Query: " & FilterQueryName & " clashes with a Field Name."
        End If
        rs.MoveNext
    Loop
    
    GetControlOptionsBySeqModel = lines.JoinArr(vbNewLine)
    GetControlOptionsBySeqModel = "export const CONTROL_OPTIONS = {" & GetControlOptionsBySeqModel & vbNewLine & "}"
    GetControlOptionsBySeqModel = GetGeneratedByFunctionSnippet(GetControlOptionsBySeqModel, "GetControlOptionsBySeqModel")
   
    
    CopyToClipboard GetControlOptionsBySeqModel
    
End Function

Public Function GetControlOptionsImportLine(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If isPresent("qrySeqModelFields", "DataType = ""ENUM"" AND SeqModelID = " & SeqModelID) Then
        GetControlOptionsImportLine = GetReplacedTemplate(rs, "GetControlOptionsImportLine")
        GetControlOptionsImportLine = GetGeneratedByFunctionSnippet(GetControlOptionsImportLine, "GetControlOptionsImportLine", "GetControlOptionsImportLine")
    End If
    CopyToClipboard GetControlOptionsImportLine
    
End Function

Public Function GetuseModelListts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function

    GetuseModelListts = GetReplacedTemplate(rs, "useModelList.ts")
    GetuseModelListts = replace(GetuseModelListts, "[GetModelPrimaryKeyName]", GetModelPrimaryKey(frm, SeqModelID, True))
    GetuseModelListts = replace(GetuseModelListts, "[GetModelUniqueName]", GetModelUniqueName(frm, SeqModelID))
    GetuseModelListts = GetGeneratedByFunctionSnippet(GetuseModelListts, "GetuseModelListts", "useModelList.ts")
    CopyToClipboard GetuseModelListts
    
    
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\hooks\heroes\useHeroList.ts
    Dim filePath: filePath = ClientPath & "src\hooks\" & ModelPath & "\use" & ModelName & "List.ts"
    WriteToFile filePath, GetuseModelListts, SeqModelID

End Function

Public Function GetModelUniqueName(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GetModelUniqueName = ELookup("tblSeqModelFields", "SeqModelID = " & SeqModelID & " AND Unique", "FieldName", "FieldOrder")
    
    If isFalse(GetModelUniqueName) Then
        GetModelUniqueName = ELookup("tblSeqModelFields", "SeqModelID = " & SeqModelID & " AND PrimaryKey", "FieldName") & ".toString()"
    End If

    CopyToClipboard GetModelUniqueName
    
End Function

Public Function GetRequiredQueryFromTanstackBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SElECT SeqModelFilterID FROm qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & _
        "AND Not ModelListID IS NULL Order By FilterOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetRequiredQueryFromTanstack(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    GetRequiredQueryFromTanstackBySeqModel = lines.JoinArr(vbNewLine)
    GetRequiredQueryFromTanstackBySeqModel = GetGeneratedByFunctionSnippet(GetRequiredQueryFromTanstackBySeqModel, "GetRequiredQueryFromTanstackBySeqModel")
    CopyToClipboard GetRequiredQueryFromTanstackBySeqModel
    
End Function

Public Function ImportControCONTROL_OPTIONS(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim list As New clsArray: list.arr = "FacetedControl,Combobox,OptionGroup,Select,CheckboxGroup"
    list.EscapeItems
    If isPresent("qrySeqModelFilters", "ControlType in (" & list.JoinArr & ") AND ListVariableName IS NULL ") Then
        ImportControCONTROL_OPTIONS = "CONTROL_OPTIONS,"
    End If
    
End Function

Public Function ImportAllUseModelListHookBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim controls As New clsArray: controls.arr = "FacetedControl,Combobox,OptionGroup,Select,CheckboxGroup"
    controls.EscapeItems
    
    sqlStr = "SElECT SeqModelFilterID FROm qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & _
        " AND Not ModelListID IS NULL ORDER BY FilterOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add ImportUseModelListHook(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    ImportAllUseModelListHookBySeqModel = lines.JoinArr(vbNewLine)
    ImportAllUseModelListHookBySeqModel = GetGeneratedByFunctionSnippet(ImportAllUseModelListHookBySeqModel, "ImportAllUseModelListHookBySeqModel")
    CopyToClipboard ImportAllUseModelListHookBySeqModel
    
End Function


Public Function GetAllBackendFiltersBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFilterID FROM tblSeqModelFilters WHERE SeqModelID = " & SeqModelID & _
        " And FilterQueryName <> ""q"" AND NOT SeqModelFieldID IS NULL ORDER BY FilterOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetBackendFilter(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    GetAllBackendFiltersBySeqModel = lines.JoinArr(vbNewLine)
    GetAllBackendFiltersBySeqModel = GetGeneratedByFunctionSnippet(GetAllBackendFiltersBySeqModel, "GetAllBackendFiltersBySeqModel")
    CopyToClipboard GetAllBackendFiltersBySeqModel
    
End Function

Public Function GetAllRelatedInterfaceImportBySeqModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID,LeftModelID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    Dim LeftModelIDs As New clsArray
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = Trim(rs.fields("LeftModelID"))
        Debug.Print LeftModelIDs.JoinArr, LeftModelID
        If Not LeftModelIDs.InArray(LeftModelID) Then
            Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
            lines.Add GetRelatedInterfaceImport(frm, SeqModelRelationshipID)
            LeftModelIDs.Add LeftModelID
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedInterfaceImportBySeqModel = lines.JoinArr(vbNewLine)
        GetAllRelatedInterfaceImportBySeqModel = GetGeneratedByFunctionSnippet(GetAllRelatedInterfaceImportBySeqModel, "GetAllRelatedInterfaceImportBySeqModel")
        CopyToClipboard GetAllRelatedInterfaceImportBySeqModel
    End If
    
End Function

Public Function GetAllRelatedUpdatePayloadBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID, Relationship FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        Dim Relationship: Relationship = rs.fields("Relationship"): If ExitIfTrue(isFalse(Relationship), "Relationship is empty..") Then Exit Function
        If Relationship = "1:1" Then
            lines.Add GetRelatedUpdatePayloadParent(frm, SeqModelRelationshipID)
        Else
            lines.Add GetRelatedUpdatePayload(frm, SeqModelRelationshipID)
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedUpdatePayloadBySeqModel = "," & lines.JoinArr("," & vbNewLine)
        GetAllRelatedUpdatePayloadBySeqModel = GetGeneratedByFunctionSnippet(GetAllRelatedUpdatePayloadBySeqModel, "GetAllRelatedUpdatePayloadBySeqModel")
        CopyToClipboard GetAllRelatedUpdatePayloadBySeqModel
    End If
    
    
End Function

Public Function GetAllAPIRelatedImportBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetAPIRelatedImport(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    GetAllAPIRelatedImportBySeqModel = lines.JoinArr(vbNewLine)
    GetAllAPIRelatedImportBySeqModel = GetGeneratedByFunctionSnippet(GetAllAPIRelatedImportBySeqModel, "GetAllAPIRelatedImportBySeqModel")
    CopyToClipboard GetAllAPIRelatedImportBySeqModel
    
End Function

Public Function GetAllExpressionLiteralFieldBySeqModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT Expression IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetExpressionLiteralField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllExpressionLiteralFieldBySeqModel = lines.JoinArr(",")
        GetAllExpressionLiteralFieldBySeqModel = GetGeneratedByFunctionSnippet(GetAllExpressionLiteralFieldBySeqModel, "GetAllExpressionLiteralFieldBySeqModel")
        CopyToClipboard GetAllExpressionLiteralFieldBySeqModel
    End If

End Function


Public Function GetAllChildModelInterfaceBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"" And Not ExcludeInTable"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetChildModelInterface(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllChildModelInterfaceBySeqModel = lines.JoinArr(vbNewLine)
        GetAllChildModelInterfaceBySeqModel = GetGeneratedByFunctionSnippet(GetAllChildModelInterfaceBySeqModel, "GetAllChildModelInterfaceBySeqModel")
        CopyToClipboard GetAllChildModelInterfaceBySeqModel
    End If
    
End Function

Public Function GetAllParentModelInterfaceBySeqModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship  = ""1:1"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetParentModelInterface(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllParentModelInterfaceBySeqModel = lines.JoinArr(vbNewLine)
        GetAllParentModelInterfaceBySeqModel = GetGeneratedByFunctionSnippet(GetAllParentModelInterfaceBySeqModel, "GetAllParentModelInterfaceBySeqModel", "")
        CopyToClipboard GetAllParentModelInterfaceBySeqModel
    End If

End Function

Public Function GetAllRelatedModelNameBySeqModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"""
    Set rs = ReturnRecordset(sqlStr)

    Dim SeqModelRelationshipID
    
    Do Until rs.EOF
        SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedPluralizedModelName(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
     sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship = ""1:1"""
    Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedModelName(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedRightModelName(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelNameBySeqModel = lines.JoinArr(vbNewLine)
        CopyToClipboard GetAllRelatedModelNameBySeqModel
    End If

End Function

Public Function GetAllOmittedModelFormKeys(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetEscapedPluralizedModelName(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllOmittedModelFormKeys = lines.JoinArr("|")
        CopyToClipboard GetAllOmittedModelFormKeys
    Else
        GetAllOmittedModelFormKeys = Esc("")
    End If

End Function

Public Function GetAllRelatedModelUpdatePayload(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"" AND NOT IsSimpleRelationship AND NOT ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetModelUpdatePayload(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelUpdatePayload = "," & lines.JoinArr("," & vbNewLine)
        GetAllRelatedModelUpdatePayload = GetGeneratedByFunctionSnippet(GetAllRelatedModelUpdatePayload, "GetAllRelatedModelUpdatePayload", "")
        CopyToClipboard GetAllRelatedModelUpdatePayload
    End If

End Function

Public Function GetAllRelatedFormikInitialValues(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"" AND Not IsSimpleRelationship " & _
        " AND NOT ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetModelFormikInitialValue(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedFormikInitialValues = "," & lines.JoinArr(",")
        CopyToClipboard GetAllRelatedFormikInitialValues
    End If

End Function

Public Function GetCOLUMNSObject(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetConstantFieldDictionary(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    GetCOLUMNSObject = "export const COLUMNS:Record<string,ColumnAttrs> = {"
    If lines.count > 0 Then
        GetCOLUMNSObject = GetCOLUMNSObject & lines.JoinArr(vbNewLine)
    End If
    
    GetCOLUMNSObject = GetCOLUMNSObject & vbNewLine & "}"
    GetCOLUMNSObject = GetGeneratedByFunctionSnippet(GetCOLUMNSObject, "GetCOLUMNSObject", "")
    CopyToClipboard GetCOLUMNSObject

End Function

Public Function GetMutationSnippets(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetMutationSnippets = GetReplacedTemplate(rs, "Mutation Snippet")
    GetMutationSnippets = GetGeneratedByFunctionSnippet(GetMutationSnippets, "GetMutationSnippets")
    CopyToClipboard GetMutationSnippets
    
End Function

Public Function GetModelUpdatePayloadExtension(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"" AND Not ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedPartialPayload(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetModelUpdatePayloadExtension = "extends " & lines.JoinArr(",")
        CopyToClipboard GetModelUpdatePayloadExtension
    End If

End Function

Public Function WriteToDetailPageNext13(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim APIOnly: APIOnly = rs.fields("APIOnly")
    
    If APIOnly Then Exit Function
    
    Dim IsSupabase: IsSupabase = rs.fields("IsSupabase")
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = "id"
    
    Dim SidebarEnabled: SidebarEnabled = rs.fields("SidebarEnabled")
    Dim templateName: templateName = IIf(SidebarEnabled, IIf(IsSupabase, "Detail Page Sidebar supabase", "Detail Page Sidebar"), "Detail Page")
    
    WriteToDetailPageNext13 = GetReplacedTemplate(rs, templateName)
    WriteToDetailPageNext13 = replace(WriteToDetailPageNext13, "[SlugOrID]", SlugOrId)
    WriteToDetailPageNext13 = GetGeneratedByFunctionSnippet(WriteToDetailPageNext13, "WriteToDetailPageNext13", templateName)
    
    ''C:\Users\User\Desktop\Web Development\panda-realty\src\app\(protected)\buyer-status\[id]\page.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\app\(protected)\" & ModelPath & "\[id]\page.tsx"
    WriteToFile filePath, WriteToDetailPageNext13, SeqModelID, "WriteToDetailPageNext13"

    CopyToClipboard WriteToDetailPageNext13
    
End Function

Public Function GenerateAllRelatedModelSubform(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add WriteToLeftModelSubform_tsx(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
End Function

Public Function WriteToUsemodelquery_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToUsemodelquery_ts = GetReplacedTemplate(rs, "useModelQuery.ts")
    WriteToUsemodelquery_ts = replace(WriteToUsemodelquery_ts, "[GetAllRelatedIndexAndID]", GetAllRelatedIndexAndID(frm, SeqModelID))
    WriteToUsemodelquery_ts = replace(WriteToUsemodelquery_ts, "[GetAllRelatedIDSimple]", GetAllRelatedIDSimple(frm, SeqModelID))
    WriteToUsemodelquery_ts = replace(WriteToUsemodelquery_ts, "[GetCodeOriginallyFromModelTable]", GetCodeOriginallyFromModelTable(frm, SeqModelID))
    WriteToUsemodelquery_ts = GetGeneratedByFunctionSnippet(WriteToUsemodelquery_ts, "WriteToUsemodelquery_ts", "useModelQuery.ts")
    
    CopyToClipboard WriteToUsemodelquery_ts
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\hooks\decks\useDeckQuery.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\hooks\" & ModelPath & "\use" & ModelName & "Query.ts"
    WriteToFile filePath, WriteToUsemodelquery_ts, SeqModelID

End Function

Public Function GetAllRelatedIndexAndID(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & "  AND Relationship <> ""1:1"" AND NOT IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetIndexAndID(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedIndexAndID = lines.JoinArr(";")
        GetAllRelatedIndexAndID = GetGeneratedByFunctionSnippet(GetAllRelatedIndexAndID, "GetAllRelatedIndexAndID")
        CopyToClipboard GetAllRelatedIndexAndID
    End If

End Function

Public Function HasFieldOrder(SeqModelID) As Boolean
    
    Dim sqlStr: sqlStr = "SELECT LeftModelID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Not ExcludeInForm" & _
        " ORDER BY LeftModelID"
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        Dim OrderField: OrderField = ELookup("qrySeqModelFields", "SeqModelID = " & LeftModelID & " AND OrderField", "FieldName")
        If Not isFalse(OrderField) Then
            HasFieldOrder = True
            Exit Function
        End If
        rs.MoveNext
    Loop
    
End Function

Public Function WriteToModelform_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    Dim UniqueField: UniqueField = ELookup("tblSeqModelFields", "Unique AND SeqModelID = " & SeqModelID, "FieldName", "FieldOrder")
    Dim GetSetRecordName: GetSetRecordName = "values." & UniqueField
    If isFalse(UniqueField) Then
        Dim PrimaryKey: PrimaryKey = ELookup("tblSeqModelFields", "PrimaryKey AND SeqModelID = " & SeqModelID, "FieldName", "FieldOrder")
        UniqueField = PrimaryKey & ".toString()"
        GetSetRecordName = "data." & PrimaryKey & " ? data." & UniqueField & " : values." & UniqueField
    End If
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim SlugOrId: SlugOrId = IIf(isFalse(SlugField), "id", "slug")
    
    Dim SidebarEnabled: SidebarEnabled = rs.fields("SidebarEnabled")
    Dim AddFormClassIfSidebar: AddFormClassIfSidebar = IIf(SidebarEnabled, "flex-1 h-full", "")
    
    Dim OrderField, GetFormOrderFieldModuleImport: OrderField = HasFieldOrder(SeqModelID)
    If OrderField Then
        GetFormOrderFieldModuleImport = GetReplacedTemplate(rs, "GetFormOrderFieldModuleImport")
    End If
    
    WriteToModelform_tsx = GetReplacedTemplate(rs, "ModelForm.tsx")
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllFormikControls]", GetAllFormikControls(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedSubforms]", GetAllRelatedSubforms(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllModelFormRelatedConstantsImport]", GetAllModelFormRelatedConstantsImport(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedModelEmptyArray]", GetAllRelatedModelEmptyArray(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedModelMapToInitialValue]", GetAllRelatedModelMapToInitialValue(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedModelSortInitialValue]", GetAllRelatedModelSortInitialValue(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedPayloadAssignment]", GetAllRelatedPayloadAssignment(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllReplaceEmptyRelatedModel]", GetAllReplaceEmptyRelatedModel(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedModelSortFromStore]", GetAllRelatedModelSortFromStore(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedListFromRelatedModel]", GetAllRelatedListFromRelatedModel(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllModelFormRequiredListImport]", GetAllModelFormRequiredListImport(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedRightModelListFromRelatedModel]", GetAllRelatedRightModelListFromRelatedModel(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllModelFormRequiredRightModelListImport]", GetAllModelFormRequiredRightModelListImport(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedSimpleModelMapToInitialValue]", GetAllRelatedSimpleModelMapToInitialValue(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllAddedAndDeletedSimpleRelationship]", GetAllAddedAndDeletedSimpleRelationship(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedSimplePayloadAssignment]", GetAllRelatedSimplePayloadAssignment(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedSimpleFacetedControl]", GetAllRelatedSimpleFacetedControl(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllUpdateOriginalRelatedSimpleModels]", GetAllUpdateOriginalRelatedSimpleModels(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[UniqueField]", UniqueField)
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllSimpleOriginalModelState]", GetAllSimpleOriginalModelState(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRightModelDefaultList]", GetAllRightModelDefaultList(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllSetOriginalSimpleModel]", GetAllSetOriginalSimpleModel(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[SlugOrId]", SlugOrId)
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedListFromRightRelatedModel]", GetAllRelatedListFromRightRelatedModel(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedModelEmptyArraySimpleOnly]", GetAllRelatedModelEmptyArraySimpleOnly(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[AddFormClassIfSidebar]", AddFormClassIfSidebar)
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetSetRecordName]", GetSetRecordName)
    
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllLeftModelDropzoneImport]", GetAllLeftModelDropzoneImport(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllModelFilesInitial]", GetAllModelFilesInitial(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllModelFilesInitialModification]", GetAllModelFilesInitialModification(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllModelDropzoneComponent]", GetAllModelDropzoneComponent(frm, SeqModelID))
    
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRequiredListImportForRelatedModelFromField]", GetAllRequiredListImportForRelatedModelFromField(frm, SeqModelID))
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetAllRelatedModelListFromField]", GetAllRelatedModelListFromField(frm, SeqModelID))
    
    WriteToModelform_tsx = replace(WriteToModelform_tsx, "[GetFormOrderFieldModuleImport]", GetFormOrderFieldModuleImport)
    
    WriteToModelform_tsx = GetGeneratedByFunctionSnippet(WriteToModelform_tsx, "WriteToModelform_tsx", "ModelForm.tsx")
    CopyToClipboard WriteToModelform_tsx
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\components\decks\DeckForm.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "Form.tsx"
    WriteToFile filePath, WriteToModelform_tsx, SeqModelID
    
End Function

Public Function GetAllFormikControls(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND NOT PrimaryKey AND NOT ControlType = ""Hidden"" ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetFormikControl(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllFormikControls = lines.JoinArr(vbNewLine)
        GetAllFormikControls = GetGeneratedByFunctionSnippet(GetAllFormikControls, "GetAllFormikControls", "", True)
        CopyToClipboard GetAllFormikControls
    End If

End Function

Public Function GetAllRelatedSubforms(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND NOT IsSimpleRelationship AND Not ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedSubform(frm, SeqModelRelationshipID)
        WriteToLeftModelSubform_tsx frm, SeqModelRelationshipID
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedSubforms = lines.JoinArr(vbNewLine)
        GetAllRelatedSubforms = GetGeneratedByFunctionSnippet(GetAllRelatedSubforms, "GetAllRelatedSubforms", "", True)
        CopyToClipboard GetAllRelatedSubforms
    End If

End Function

Public Function GetAllModelFormRelatedConstantsImport(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"" AND NOT IsSimpleRelationship" & _
        " AND NOT ExcludeInForm"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetModelFormRelatedConstantsImport(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFormRelatedConstantsImport = lines.JoinArr(vbNewLine)
        GetAllModelFormRelatedConstantsImport = GetGeneratedByFunctionSnippet(GetAllModelFormRelatedConstantsImport, "GetAllModelFormRelatedConstantsImport")
        CopyToClipboard GetAllModelFormRelatedConstantsImport
    End If

End Function

Public Function GetAllRelatedModelEmptyArray(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"" AND Not IsSimpleRelationship" & _
        " AND NOT ExcludeInForm"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedModelEmptyArray(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelEmptyArray = lines.JoinArr(vbNewLine)
        GetAllRelatedModelEmptyArray = GetGeneratedByFunctionSnippet(GetAllRelatedModelEmptyArray, "GetAllRelatedModelEmptyArray", "")
        CopyToClipboard GetAllRelatedModelEmptyArray
    End If

End Function

Public Function GetAllRelatedModelMapToInitialValue(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"" AND NOT IsSimpleRelationship" & _
        " AND NOT ExcludeInForm AND NOT IncludeAsDropzone"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedModelMapToInitialValue(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelMapToInitialValue = lines.JoinArr(vbNewLine)
        GetAllRelatedModelMapToInitialValue = GetGeneratedByFunctionSnippet(GetAllRelatedModelMapToInitialValue, "GetAllRelatedModelMapToInitialValue")
        CopyToClipboard GetAllRelatedModelMapToInitialValue
    End If

End Function

Public Function GetAllRelatedModelSortInitialValue(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND NOT IsSimpleRelationship AND Not ExcludeInForm"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedModelSortInitialValue(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelSortInitialValue = lines.JoinArr(vbNewLine)
        GetAllRelatedModelSortInitialValue = GetGeneratedByFunctionSnippet(GetAllRelatedModelSortInitialValue, "GetAllRelatedModelSortInitialValue")
        CopyToClipboard GetAllRelatedModelSortInitialValue
    End If

End Function

Public Function GetAllRelatedPayloadAssignment(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID

    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"" AND NOT IsSimpleRelationship" & _
        " AND Not ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedPayloadAssignment(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedPayloadAssignment = lines.JoinArr(vbNewLine)
        GetAllRelatedPayloadAssignment = GetGeneratedByFunctionSnippet(GetAllRelatedPayloadAssignment, "GetAllRelatedPayloadAssignment")
        CopyToClipboard GetAllRelatedPayloadAssignment
    End If

End Function

Public Function GetAllReplaceEmptyRelatedModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND NOT IsSimpleRelationship AND Not ExcludeInForm"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetReplaceEmptyRelatedModel(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllReplaceEmptyRelatedModel = lines.JoinArr(vbNewLine)
        GetAllReplaceEmptyRelatedModel = GetGeneratedByFunctionSnippet(GetAllReplaceEmptyRelatedModel, "GetAllReplaceEmptyRelatedModel")
        CopyToClipboard GetAllReplaceEmptyRelatedModel
    End If

End Function

Public Function GetAllRelatedModelSortFromStore(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND Not IsSimpleRelationship AND Not ExcludeInForm"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedModelSortFromStore(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelSortFromStore = lines.JoinArr(vbNewLine)
        GetAllRelatedModelSortFromStore = GetGeneratedByFunctionSnippet(GetAllRelatedModelSortFromStore, "GetAllRelatedModelSortFromStore")
        CopyToClipboard GetAllRelatedModelSortFromStore
    End If

End Function

Public Function GetAllModelFormRequiredListImportFromModel(frm As Object, Optional SeqModelID = "", Optional CallingFromModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND Not ExcludeInRequiredList"
    
    If Not isFalse(CallingFromModelID) Then
        sqlStr = sqlStr & " AND RightModelID <> " & CallingFromModelID
    End If
    
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetModelFormRequiredListImport(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFormRequiredListImportFromModel = lines.JoinArr(vbNewLine)
        CopyToClipboard GetAllModelFormRequiredListImportFromModel
    End If

End Function

Public Function GetAllModelFormRequiredListImport(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID, LeftModelID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        lines.Add GetAllModelFormRequiredListImportFromModel(frm, LeftModelID, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFormRequiredListImport = lines.JoinArr(vbNewLine)
        GetAllModelFormRequiredListImport = GetGeneratedByFunctionSnippet(GetAllModelFormRequiredListImport, "GetAllModelFormRequiredListImport")
        CopyToClipboard GetAllModelFormRequiredListImport
    End If

End Function

Public Function GetActionCell(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim IsRowAction: IsRowAction = rs.fields("IsRowAction")
    Dim ModelName: ModelName = rs.fields("ModelName")
    
    If IsRowAction Then
        GetActionCell = "<ModelRowActions cell={cell} />"
        ''WriteToModelrowactions_tsx frm, SeqModelID
    Else
        GetActionCell = "<DeleteRowColumn {...cell} />"
    End If
    
    CopyToClipboard GetActionCell
    
End Function

Public Function WriteToModelrowactions_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim PrimaryKey: PrimaryKey = ELookup("tblSeqModelFields", "PrimaryKey AND SeqModelID = " & SeqModelID, "FieldName")
    Dim SlugOrId: SlugOrId = IIf(isFalse(SlugField), PrimaryKey, "slug")
    
    WriteToModelrowactions_tsx = GetReplacedTemplate(rs, "ModelRowActions.tsx")
    WriteToModelrowactions_tsx = replace(WriteToModelrowactions_tsx, "[SlugOrId]", SlugOrId)
    WriteToModelrowactions_tsx = GetGeneratedByFunctionSnippet(WriteToModelrowactions_tsx, "WriteToModelrowactions_tsx", "ModelRowActions.tsx")
    CopyToClipboard WriteToModelrowactions_tsx
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\components\decks\DeckRowActions.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "RowActions.tsx"
    WriteToFile filePath, WriteToModelrowactions_tsx, SeqModelID
    
End Function

Public Function GetUpdateFunctionWithRelationshipNext13(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetUpdateFunctionWithRelationshipNext13 = GetReplacedTemplate(rs, "Update With Relationship Next 13")
    GetUpdateFunctionWithRelationshipNext13 = replace(GetUpdateFunctionWithRelationshipNext13, "[GetAllRelatedPluralizedModelName]", GetAllRelatedPluralizedModelName(frm, SeqModelID))
    GetUpdateFunctionWithRelationshipNext13 = replace(GetUpdateFunctionWithRelationshipNext13, "[GetAllRelatedModelUpdateOrInsert]", GetAllRelatedModelUpdateOrInsert(frm, SeqModelID))
    GetUpdateFunctionWithRelationshipNext13 = replace(GetUpdateFunctionWithRelationshipNext13, "[GetAllRelatedModelKeyValue]", GetAllRelatedModelKeyValue(frm, SeqModelID))
    GetUpdateFunctionWithRelationshipNext13 = replace(GetUpdateFunctionWithRelationshipNext13, "[GetAllRelatedSimpleModelFromRes]", GetAllRelatedSimpleModelFromRes(frm, SeqModelID))
    GetUpdateFunctionWithRelationshipNext13 = replace(GetUpdateFunctionWithRelationshipNext13, "[GetAllThroughModelUpdateOrInsert]", GetAllThroughModelUpdateOrInsert(frm, SeqModelID))
    GetUpdateFunctionWithRelationshipNext13 = replace(GetUpdateFunctionWithRelationshipNext13, "[GetAllRelatedDropzoneModel]", GetAllRelatedDropzoneModel(frm, SeqModelID))
    GetUpdateFunctionWithRelationshipNext13 = replace(GetUpdateFunctionWithRelationshipNext13, "[GetAllRelatedDropzoneModelUpdateOrInsert]", GetAllRelatedDropzoneModelUpdateOrInsert(frm, SeqModelID))
    GetUpdateFunctionWithRelationshipNext13 = GetGeneratedByFunctionSnippet(GetUpdateFunctionWithRelationshipNext13, "GetUpdateFunctionWithRelationshipNext13", "Update With Relationship Next 13")
    CopyToClipboard GetUpdateFunctionWithRelationshipNext13
    
End Function

Public Function GetAllRelatedPluralizedModelName(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID, LeftPluralizedModelName FROM qrySeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND NOT IsSimpleRelationship AND Not ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        Dim LeftPluralizedModelName: LeftPluralizedModelName = rs.fields("LeftPluralizedModelName"): If ExitIfTrue(isFalse(LeftPluralizedModelName), "LeftPluralizedModelName is empty..") Then Exit Function
        lines.Add LeftPluralizedModelName
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedPluralizedModelName = "const { " & lines.JoinArr(",") & " } = res;"
        GetAllRelatedPluralizedModelName = GetGeneratedByFunctionSnippet(GetAllRelatedPluralizedModelName, "GetAllRelatedPluralizedModelName", "")
        CopyToClipboard GetAllRelatedPluralizedModelName
    End If

End Function

Public Function GetAllRelatedModelUpdateOrInsert(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND NOT IsSimpleRelationship AND NOT ExcludeInForm"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedModelUpdateOrInsert(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelUpdateOrInsert = lines.JoinArr(vbNewLine)
        GetAllRelatedModelUpdateOrInsert = GetGeneratedByFunctionSnippet(GetAllRelatedModelUpdateOrInsert, "GetAllRelatedModelUpdateOrInsert")
        CopyToClipboard GetAllRelatedModelUpdateOrInsert
    End If

End Function

Public Function GetAllRelatedModelKeyValue(frm As Object, Optional SeqModelID = "", Optional singleRoute As Boolean = True) As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND Not ExcludeInForm"
        
    If Not singleRoute Then
        sqlStr = sqlStr & " AND Not isSimpleRelationship"
    End If

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedModelKeyValue(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelKeyValue = lines.JoinArr(vbNewLine)
        GetAllRelatedModelKeyValue = GetGeneratedByFunctionSnippet(GetAllRelatedModelKeyValue, "GetAllRelatedModelKeyValue")
        CopyToClipboard GetAllRelatedModelKeyValue
    End If

End Function

Public Function GetModelRowActionsImport(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim IsRowAction: IsRowAction = rs.fields("IsRowAction")
    
    If IsRowAction Then
        GetModelRowActionsImport = GetReplacedTemplate(rs, "GetModelRowActionsImport")
        GetModelRowActionsImport = GetGeneratedByFunctionSnippet(GetModelRowActionsImport, "GetModelRowActionsImport", "GetModelRowActionsImport")
        CopyToClipboard GetModelRowActionsImport
    End If
    
End Function

Public Function WriteToModeldatatable_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToModeldatatable_tsx = GetReplacedTemplate(rs, "ModelDataTable.tsx")
    WriteToModeldatatable_tsx = replace(WriteToModeldatatable_tsx, "[GetAllHiddenColumns]", GetAllHiddenColumns(frm, SeqModelID))
    WriteToModeldatatable_tsx = GetGeneratedByFunctionSnippet(WriteToModeldatatable_tsx, "WriteToModeldatatable_tsx", "ModelDataTable.tsx")
    CopyToClipboard WriteToModeldatatable_tsx
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\components\cards\CardDataTable.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "DataTable.tsx"
    WriteToFile filePath, WriteToModeldatatable_tsx, SeqModelID
    
End Function

Public Function GetAllRelatedRightModelImport(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID,RightModelID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID
    Set rs = ReturnRecordset(sqlStr)
    Dim RightModels As New clsArray
    Do Until rs.EOF
        Dim RightModelID: RightModelID = rs.fields("RightModelID")
        RightModelID = Trim(RightModelID)
        Debug.Print RightModelID
        If Not RightModels.InArray(RightModelID) Then
            Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
            lines.Add GetRelatedRightModelImport(frm, SeqModelRelationshipID)
            RightModels.Add RightModelID
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedRightModelImport = lines.JoinArr(vbNewLine)
        GetAllRelatedRightModelImport = GetGeneratedByFunctionSnippet(GetAllRelatedRightModelImport, "GetAllRelatedRightModelImport")
        CopyToClipboard GetAllRelatedRightModelImport
    End If

End Function

Public Function GetAllRelatedRightModelInterface(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedRightModelInterface(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedRightModelInterface = lines.JoinArr(vbNewLine)
        GetAllRelatedRightModelInterface = GetGeneratedByFunctionSnippet(GetAllRelatedRightModelInterface, "GetAllRelatedRightModelInterface")
        CopyToClipboard GetAllRelatedRightModelInterface
    End If

End Function

Public Function GetAllSQLRightJoinSnippets(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND NOT ExcludeInTable"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSQLRightJoinSnippetFromRelationship(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSQLRightJoinSnippets = lines.JoinArr(vbNewLine)
        GetAllSQLRightJoinSnippets = GetGeneratedByFunctionSnippet(GetAllSQLRightJoinSnippets, "GetAllSQLRightJoinSnippets")
        CopyToClipboard GetAllSQLRightJoinSnippets
    End If

End Function

Public Function GetAllRightModelJoinCancellationSnippet(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND NOT ExcludeInTable"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRightModelJoinCancellationSnippet(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRightModelJoinCancellationSnippet = lines.JoinArr(vbNewLine)
        GetAllRightModelJoinCancellationSnippet = GetGeneratedByFunctionSnippet(GetAllRightModelJoinCancellationSnippet, "GetAllRightModelJoinCancellationSnippet")
        CopyToClipboard GetAllRightModelJoinCancellationSnippet
    End If

End Function

Public Function GetAllRightJoinName(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND NOT ExcludeInTable"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRightJoinName(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRightJoinName = lines.JoinArr(vbNewLine)
        GetAllRightJoinName = GetGeneratedByFunctionSnippet(GetAllRightJoinName, "GetAllRightJoinName")
        CopyToClipboard GetAllRightJoinName
    End If

End Function

Public Function GetAllGetmodelsqlChildNext13(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND NOT ExcludeInTable"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRightModelgetModelSQLSnippet(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllGetmodelsqlChildNext13 = lines.JoinArr(vbNewLine)
        GetAllGetmodelsqlChildNext13 = GetGeneratedByFunctionSnippet(GetAllGetmodelsqlChildNext13, "GetAllGetmodelsqlChildNext13")
        CopyToClipboard GetAllGetmodelsqlChildNext13
    End If

End Function

Public Function GetAllUniqueKeysOption(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & " AND Not UniqueWith IS NULL ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetUniqueKeysOption(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllUniqueKeysOption = "uniqueKeys: {" & vbNewLine & lines.JoinArr(vbNewLine) & vbNewLine & "},"
        GetAllUniqueKeysOption = GetGeneratedByFunctionSnippet(GetAllUniqueKeysOption, "GetAllUniqueKeysOption")
        CopyToClipboard GetAllUniqueKeysOption
    End If

End Function

Public Function GetAllSQLLeftJoinSnippets(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND NOT ExcludeInTable " & _
        " ORDER BY SeqModelRelationshipID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSQLLeftJoinSnippetFromRelationship(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSQLLeftJoinSnippets = lines.JoinArr(vbNewLine)
        GetAllSQLLeftJoinSnippets = GetGeneratedByFunctionSnippet(GetAllSQLLeftJoinSnippets, "GetAllSQLLeftJoinSnippets")
        CopyToClipboard GetAllSQLLeftJoinSnippets
    End If

End Function

Public Function GetAllLeftModelJoinCancellationSnippet(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND NOT ExcludeInTable"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetLeftModelJoinCancellationSnippet(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllLeftModelJoinCancellationSnippet = lines.JoinArr(vbNewLine)
        GetAllLeftModelJoinCancellationSnippet = GetGeneratedByFunctionSnippet(GetAllLeftModelJoinCancellationSnippet, "GetAllLeftModelJoinCancellationSnippet")
        CopyToClipboard GetAllLeftModelJoinCancellationSnippet
    End If

End Function

Public Function GetAllLeftJoinName(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND NOT ExcludeInTable"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetLeftJoinName(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllLeftJoinName = lines.JoinArr(vbNewLine)
        GetAllLeftJoinName = GetGeneratedByFunctionSnippet(GetAllLeftJoinName, "GetAllLeftJoinName")
        CopyToClipboard GetAllLeftJoinName
    End If

End Function

Public Function GetAllLeftModelReduceResultAndRemoveDuplicates(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND NOT ExcludeInTable"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetLeftModelReduceResultAndRemoveDuplicates(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllLeftModelReduceResultAndRemoveDuplicates = lines.JoinArr(vbNewLine)
        GetAllLeftModelReduceResultAndRemoveDuplicates = GetGeneratedByFunctionSnippet(GetAllLeftModelReduceResultAndRemoveDuplicates, "GetAllLeftModelReduceResultAndRemoveDuplicates")
        CopyToClipboard GetAllLeftModelReduceResultAndRemoveDuplicates
    End If

End Function

Public Function GetAllGetmodelsqlLeftModelChildNext13(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetGetmodelsqlLeftModelChildNext13(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllGetmodelsqlLeftModelChildNext13 = lines.JoinArr(vbNewLine)
        GetAllGetmodelsqlLeftModelChildNext13 = GetGeneratedByFunctionSnippet(GetAllGetmodelsqlLeftModelChildNext13, "GetAllGetmodelsqlLeftModelChildNext13")
        CopyToClipboard GetAllGetmodelsqlLeftModelChildNext13
    End If

End Function

Public Function GetAllRelatedRightModelListFromRelatedModel(frm As Object, Optional SeqModelID = "", Optional ParentModelID) As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID, RightModelID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND Not ExcludeInRequiredList"
    If Not isFalse(ParentModelID) Then
        sqlStr = sqlStr & " AND RightModelID <> " & ParentModelID
    End If
    
    Dim RightModelIDs As New clsArray
    
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim RightModelID: RightModelID = Trim(rs.fields("RightModelID"))
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        If Not RightModelIDs.InArray(RightModelID) Then
            lines.Add GetRelatedRightModelListFromRelatedModel(frm, SeqModelRelationshipID)
            RightModelIDs.Add RightModelID
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedRightModelListFromRelatedModel = lines.JoinArr(vbNewLine)
        GetAllRelatedRightModelListFromRelatedModel = GetGeneratedByFunctionSnippet(GetAllRelatedRightModelListFromRelatedModel, "GetAllRelatedRightModelListFromRelatedModel")
        CopyToClipboard GetAllRelatedRightModelListFromRelatedModel
    End If

End Function

Public Function GetAllModelFormRequiredRightModelListImport(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim RightModelIDs As New clsArray
    
    ''Unique RightModelID only
    sqlStr = "SELECT SeqModelRelationshipID, RightModelID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND Not ExcludeInRequiredList"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim RightModelID: RightModelID = Trim(rs.fields("RightModelID")): If ExitIfTrue(isFalse(RightModelID), "RightModelID is empty..") Then Exit Function
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        If Not RightModelIDs.InArray(RightModelID) Then
            lines.Add GetModelFormRequiredRightModelListImport(frm, SeqModelRelationshipID)
            RightModelIDs.Add RightModelID
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFormRequiredRightModelListImport = lines.JoinArr(vbNewLine)
        GetAllModelFormRequiredRightModelListImport = GetGeneratedByFunctionSnippet(GetAllModelFormRequiredRightModelListImport, "GetAllModelFormRequiredRightModelListImport")
        CopyToClipboard GetAllModelFormRequiredRightModelListImport
    End If

End Function

Public Function WriteToModelsRouteApiNoForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToModelsRouteApiNoForm = GetReplacedTemplate(rs, "models route next 13 no form")
    WriteToModelsRouteApiNoForm = replace(WriteToModelsRouteApiNoForm, "[GetAllFieldsToUpdateBySeqModel]", GetAllFieldsToUpdateBySeqModel(frm, SeqModelID))
    WriteToModelsRouteApiNoForm = replace(WriteToModelsRouteApiNoForm, "[GetAllModelAttributesBySeqModel]", GetAllModelAttributesBySeqModel(frm, SeqModelID))
    WriteToModelsRouteApiNoForm = replace(WriteToModelsRouteApiNoForm, "[GenerateIncludeOption]", GenerateIncludeOption(frm, SeqModelID))
    WriteToModelsRouteApiNoForm = replace(WriteToModelsRouteApiNoForm, "[GenerateAttributesOption]", GenerateAttributesOption(frm, SeqModelID))
    WriteToModelsRouteApiNoForm = GetGeneratedByFunctionSnippet(WriteToModelsRouteApiNoForm, "WriteToModelsRouteApiNoForm", "models route next 13 no form")
    CopyToClipboard WriteToModelsRouteApiNoForm
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\app\api\card-card-keywords\route.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\app\api\" & ModelPath & "\route.ts"
    WriteToFile filePath, WriteToModelsRouteApiNoForm, SeqModelID
    
    
End Function

Public Function GetAllAPIRelatedLeftModelImportBySeqModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Not ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetAPIRelatedLeftModelImport(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllAPIRelatedLeftModelImportBySeqModel = lines.JoinArr(vbNewLine)
        GetAllAPIRelatedLeftModelImportBySeqModel = GetGeneratedByFunctionSnippet(GetAllAPIRelatedLeftModelImportBySeqModel, "GetAllAPIRelatedLeftModelImportBySeqModel")
        CopyToClipboard GetAllAPIRelatedLeftModelImportBySeqModel
    End If

End Function

Public Function GetAllAPIRelatedRightModelImportBySeqModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetAPIRelatedRightModelImport(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllAPIRelatedRightModelImportBySeqModel = lines.JoinArr(vbNewLine)
        GetAllAPIRelatedRightModelImportBySeqModel = GetGeneratedByFunctionSnippet(GetAllAPIRelatedRightModelImportBySeqModel, "GetAllAPIRelatedRightModelImportBySeqModel")
        CopyToClipboard GetAllAPIRelatedRightModelImportBySeqModel
    End If

End Function

Public Function GetAllRelatedSimpleModelFromRes(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND IsSimpleRelationship AND Not Through IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedSimpleModelFromRes(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedSimpleModelFromRes = lines.JoinArr(vbNewLine)
        GetAllRelatedSimpleModelFromRes = GetGeneratedByFunctionSnippet(GetAllRelatedSimpleModelFromRes, "GetAllRelatedSimpleModelFromRes")
        CopyToClipboard GetAllRelatedSimpleModelFromRes
    End If

End Function

Public Function GetAllThroughModelUpdateOrInsert(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND IsSimpleRelationship AND Not Through IS NULL"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetThroughModelUpdateOrInsert(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllThroughModelUpdateOrInsert = lines.JoinArr(vbNewLine)
        GetAllThroughModelUpdateOrInsert = GetGeneratedByFunctionSnippet(GetAllThroughModelUpdateOrInsert, "GetAllThroughModelUpdateOrInsert")
        CopyToClipboard GetAllThroughModelUpdateOrInsert
    End If

End Function

Public Function GetAllRelatedListFromRelatedModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetAllRelatedListFromModel(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedListFromRelatedModel = lines.JoinArr(vbNewLine)
        GetAllRelatedListFromRelatedModel = GetGeneratedByFunctionSnippet(GetAllRelatedListFromRelatedModel, "GetAllRelatedListFromRelatedModel")
        CopyToClipboard GetAllRelatedListFromRelatedModel
    End If

End Function

Public Function GetAllOriginalSimpleRelatedModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetOriginalSimpleRelatedModel(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllOriginalSimpleRelatedModel = lines.JoinArr(vbNewLine)
        GetAllOriginalSimpleRelatedModel = GetGeneratedByFunctionSnippet(GetAllOriginalSimpleRelatedModel, "GetAllOriginalSimpleRelatedModel")
        CopyToClipboard GetAllOriginalSimpleRelatedModel
    End If

End Function

Public Function GetAllRelatedSimpleModelMapToInitialValue(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedSimpleModelMapToInitialValue(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedSimpleModelMapToInitialValue = lines.JoinArr(vbNewLine)
        GetAllRelatedSimpleModelMapToInitialValue = GetGeneratedByFunctionSnippet(GetAllRelatedSimpleModelMapToInitialValue, "GetAllRelatedSimpleModelMapToInitialValue")
        CopyToClipboard GetAllRelatedSimpleModelMapToInitialValue
    End If

End Function
Public Function GetAllAddedAndDeletedSimpleRelationship(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetAddedAndDeletedSimpleRelationship(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllAddedAndDeletedSimpleRelationship = lines.JoinArr(vbNewLine)
        GetAllAddedAndDeletedSimpleRelationship = GetGeneratedByFunctionSnippet(GetAllAddedAndDeletedSimpleRelationship, "GetAllAddedAndDeletedSimpleRelationship")
        CopyToClipboard GetAllAddedAndDeletedSimpleRelationship
    End If

End Function
Public Function GetAllRelatedSimplePayloadAssignment(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedSimplePayloadAssignment(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedSimplePayloadAssignment = lines.JoinArr(vbNewLine)
        GetAllRelatedSimplePayloadAssignment = GetGeneratedByFunctionSnippet(GetAllRelatedSimplePayloadAssignment, "GetAllRelatedSimplePayloadAssignment")
        CopyToClipboard GetAllRelatedSimplePayloadAssignment
    End If

End Function
Public Function GetAllRelatedSimpleFacetedControl(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedSimpleFacetedControl(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedSimpleFacetedControl = lines.JoinArr(vbNewLine)
        GetAllRelatedSimpleFacetedControl = GetGeneratedByFunctionSnippet(GetAllRelatedSimpleFacetedControl, "GetAllRelatedSimpleFacetedControl", "", True)
        CopyToClipboard GetAllRelatedSimpleFacetedControl
    End If

End Function
Public Function GetAllUpdateOriginalRelatedSimpleModels(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetUpdateOriginalRelatedSimpleModels(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllUpdateOriginalRelatedSimpleModels = lines.JoinArr(vbNewLine)
        GetAllUpdateOriginalRelatedSimpleModels = GetGeneratedByFunctionSnippet(GetAllUpdateOriginalRelatedSimpleModels, "GetAllUpdateOriginalRelatedSimpleModels")
        CopyToClipboard GetAllUpdateOriginalRelatedSimpleModels
    End If

End Function
Public Function GetAllRelatedIDSimple(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedIDSimple(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedIDSimple = lines.JoinArr(vbNewLine)
        GetAllRelatedIDSimple = GetGeneratedByFunctionSnippet(GetAllRelatedIDSimple, "GetAllRelatedIDSimple")
        CopyToClipboard GetAllRelatedIDSimple
    End If

End Function
Public Function GetAllSimpleRelatedKey(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSimpleRelatedKey(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSimpleRelatedKey = lines.JoinArr(vbNewLine)
        GetAllSimpleRelatedKey = GetGeneratedByFunctionSnippet(GetAllSimpleRelatedKey, "GetAllSimpleRelatedKey")
        CopyToClipboard GetAllSimpleRelatedKey
    End If

End Function
Public Function GetAllSimplePluralizedFieldName(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSimplePluralizedFieldName(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSimplePluralizedFieldName = lines.JoinArr(" ")
        GetAllSimplePluralizedFieldName = GetGeneratedByFunctionSnippet(GetAllSimplePluralizedFieldName, "GetAllSimplePluralizedFieldName")
        CopyToClipboard GetAllSimplePluralizedFieldName
    End If

End Function

Public Function GetAllSimpleRelatedKeyPayload(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSimpleRelatedKeyPayload(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSimpleRelatedKeyPayload = lines.JoinArr(vbNewLine)
        GetAllSimpleRelatedKeyPayload = GetGeneratedByFunctionSnippet(GetAllSimpleRelatedKeyPayload, "GetAllSimpleRelatedKeyPayload")
        CopyToClipboard GetAllSimpleRelatedKeyPayload
    End If

End Function

Public Function GetMultiCreateModelPOSTRoute(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetMultiCreateModelPOSTRoute = GetReplacedTemplate(rs, "GetMultiCreateModelPOSTRoute")
    GetMultiCreateModelPOSTRoute = GetGeneratedByFunctionSnippet(GetMultiCreateModelPOSTRoute, "GetMultiCreateModelPOSTRoute", "GetMultiCreateModelPOSTRoute")
    CopyToClipboard GetMultiCreateModelPOSTRoute
    
End Function

Public Function GetSingleCreateModelPOSTRoute(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim GetSlugReturn: GetSlugReturn = IIf(isFalse(SlugField), "", "slug: new" & ModelName & ".slug,")
    
    GetSingleCreateModelPOSTRoute = GetReplacedTemplate(rs, "GetSingleCreateModelPOSTRoute")
    GetSingleCreateModelPOSTRoute = replace(GetSingleCreateModelPOSTRoute, "[GetAllSimpleModelInserts]", GetAllSimpleModelInserts(frm, SeqModelID))
    GetSingleCreateModelPOSTRoute = replace(GetSingleCreateModelPOSTRoute, "[GetAllSimplePluralizedModelName]", GetAllSimplePluralizedModelName(frm, SeqModelID))
    GetSingleCreateModelPOSTRoute = replace(GetSingleCreateModelPOSTRoute, "[GetSlugReturn]", GetSlugReturn)
    GetSingleCreateModelPOSTRoute = replace(GetSingleCreateModelPOSTRoute, "[GetAllRelatedModelUpdateOrInsert]", GetAllRelatedModelUpdateOrInsert(frm, SeqModelID))
    GetSingleCreateModelPOSTRoute = replace(GetSingleCreateModelPOSTRoute, "[GetAllRelatedPluralizedModelName]", GetAllRelatedPluralizedModelName(frm, SeqModelID))
    GetSingleCreateModelPOSTRoute = replace(GetSingleCreateModelPOSTRoute, "[GetAllRelatedModelKeyValue]", GetAllRelatedModelKeyValue(frm, SeqModelID, False))
    GetSingleCreateModelPOSTRoute = GetGeneratedByFunctionSnippet(GetSingleCreateModelPOSTRoute, "GetSingleCreateModelPOSTRoute", "GetSingleCreateModelPOSTRoute")
    CopyToClipboard GetSingleCreateModelPOSTRoute
    
End Function

Public Function GetAllSimpleModelInserts(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSimpleModelInserts(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSimpleModelInserts = lines.JoinArr(vbNewLine)
        GetAllSimpleModelInserts = GetGeneratedByFunctionSnippet(GetAllSimpleModelInserts, "GetAllSimpleModelInserts")
        CopyToClipboard GetAllSimpleModelInserts
    End If

End Function

Public Function GetAllSimplePluralizedModelName(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSimplePluralizedModelName(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSimplePluralizedModelName = lines.JoinArr(vbNewLine)
        GetAllSimplePluralizedModelName = GetGeneratedByFunctionSnippet(GetAllSimplePluralizedModelName, "GetAllSimplePluralizedModelName")
        CopyToClipboard GetAllSimplePluralizedModelName
    End If

End Function
Public Function GetAllCreateSimpleModelFromRoute(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetCreateSimpleModelFromRoute(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllCreateSimpleModelFromRoute = lines.JoinArr(vbNewLine)
        GetAllCreateSimpleModelFromRoute = GetGeneratedByFunctionSnippet(GetAllCreateSimpleModelFromRoute, "GetAllCreateSimpleModelFromRoute")
        CopyToClipboard GetAllCreateSimpleModelFromRoute
    End If

End Function

Public Function GetOnSuccessFormikDeleteDialogTable(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetOnSuccessFormikDeleteDialogTable = GetReplacedTemplate(rs, "GetOnSuccessFormikDeleteDialogTable")
    GetOnSuccessFormikDeleteDialogTable = GetGeneratedByFunctionSnippet(GetOnSuccessFormikDeleteDialogTable, "GetOnSuccessFormikDeleteDialogTable", "GetOnSuccessFormikDeleteDialogTable")
    CopyToClipboard GetOnSuccessFormikDeleteDialogTable
    
End Function

Public Function GetAllSimpleOriginalModelState(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Relationship <> ""1:1""" & _
        " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSimpleOriginalModelState(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSimpleOriginalModelState = lines.JoinArr(vbNewLine)
        GetAllSimpleOriginalModelState = GetGeneratedByFunctionSnippet(GetAllSimpleOriginalModelState, "GetAllSimpleOriginalModelState")
        CopyToClipboard GetAllSimpleOriginalModelState
    End If

End Function
Public Function GetAllRightModelDefaultList(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND Relationship <> ""1:1"" AND Not ExcludeInRequiredList"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRightModelDefaultList(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRightModelDefaultList = lines.JoinArr(vbNewLine)
        GetAllRightModelDefaultList = GetGeneratedByFunctionSnippet(GetAllRightModelDefaultList, "GetAllRightModelDefaultList")
        CopyToClipboard GetAllRightModelDefaultList
    End If

End Function

Public Function WriteToFullTextIndexMigrationFile(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim AllDatabaseFieldName: AllDatabaseFieldName = GetAllDatabaseFieldName(frm, SeqModelID)
    
    Dim fieldArr As New clsArray: fieldArr.arr = AllDatabaseFieldName
    Dim GetAllQUnderscoreFieldName: GetAllQUnderscoreFieldName = fieldArr.JoinArr("_")
    
    fieldArr.EscapeItems
    Dim GetAllQFieldName: GetAllQFieldName = fieldArr.JoinArr

    WriteToFullTextIndexMigrationFile = GetReplacedTemplate(rs, "FULLTEXT index for q fields")
    WriteToFullTextIndexMigrationFile = replace(WriteToFullTextIndexMigrationFile, "[GetAllQFieldName]", GetAllQFieldName)
    WriteToFullTextIndexMigrationFile = replace(WriteToFullTextIndexMigrationFile, "[GetAllQUnderscoreFieldName]", GetAllQUnderscoreFieldName)
    WriteToFullTextIndexMigrationFile = GetGeneratedByFunctionSnippet(WriteToFullTextIndexMigrationFile, "WriteToFullTextIndexMigrationFile", "FULLTEXT index for q fields")
    CopyToClipboard WriteToFullTextIndexMigrationFile
    
    Dim fileName: fileName = ConvertToCustomTimestamp & "-add_full_text_" & GetAllQUnderscoreFieldName & ".js"
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\migrations\20230808130427-create_marvelduel_card.js
    Dim filePath: filePath = ClientPath & "src\mugrations\" & fileName
    WriteToFile filePath, WriteToFullTextIndexMigrationFile, SeqModelID
    
End Function

Public Function GetAllDatabaseFieldName(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    sqlStr = "SELECT SeqModelFilterID FROM tblSeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterQueryName = ""q"""
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetDatabaseFieldName(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllDatabaseFieldName = lines.JoinArr(",")
        CopyToClipboard GetAllDatabaseFieldName
    End If

End Function

Public Function GetAddFulltextIndexSql(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim AllDatabaseFieldName: AllDatabaseFieldName = GetAllDatabaseFieldName(frm, SeqModelID)
    
    Dim fields As New clsArray: fields.arr = AllDatabaseFieldName
    
    Dim GetAllQUnderscoreFieldName: GetAllQUnderscoreFieldName = fields.JoinArr("_")
    
    GetAddFulltextIndexSql = GetReplacedTemplate(rs, "ADD FULLTEXT INDEX SQL")
    GetAddFulltextIndexSql = replace(GetAddFulltextIndexSql, "[GetAllQUnderscoreFieldName]", GetAllQUnderscoreFieldName)
    GetAddFulltextIndexSql = replace(GetAddFulltextIndexSql, "[GetAllDatabaseFieldName]", AllDatabaseFieldName)
    ''GetAddFulltextIndexSql = GetGeneratedByFunctionSnippet(GetAddFulltextIndexSql, "GetAddFulltextIndexSql", "ADD FULLTEXT INDEX SQL")
    CopyToClipboard GetAddFulltextIndexSql
    
End Function

Public Function GetDropFulltextIndexSql(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim AllDatabaseFieldName: AllDatabaseFieldName = GetAllDatabaseFieldName(frm, SeqModelID)
    
    Dim fields As New clsArray: fields.arr = AllDatabaseFieldName
    
    Dim GetAllQUnderscoreFieldName: GetAllQUnderscoreFieldName = fields.JoinArr("_")

    GetDropFulltextIndexSql = GetReplacedTemplate(rs, "DROP FULLTEXT INDEX SQL")
    GetDropFulltextIndexSql = replace(GetDropFulltextIndexSql, "[GetAllQUnderscoreFieldName]", GetAllQUnderscoreFieldName)
    GetDropFulltextIndexSql = GetGeneratedByFunctionSnippet(GetDropFulltextIndexSql, "GetDropFulltextIndexSql", "DROP FULLTEXT INDEX SQL")
    CopyToClipboard GetDropFulltextIndexSql
    
End Function

Public Function GetAllSetOriginalSimpleModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetSetOriginalSimpleModel(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSetOriginalSimpleModel = lines.JoinArr(vbNewLine)
        GetAllSetOriginalSimpleModel = GetGeneratedByFunctionSnippet(GetAllSetOriginalSimpleModel, "GetAllSetOriginalSimpleModel")
        CopyToClipboard GetAllSetOriginalSimpleModel
    End If

End Function

Public Function GetAllRightModelListImportForColumn(frm As Object, Optional SeqModelID = "", Optional ParentModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND NOT ExcludeInRequiredList"
    If Not isFalse(ParentModelID) Then
        sqlStr = sqlStr & " AND RightModelID <> " & ParentModelID
    End If
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRightModelListImportForColumn(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRightModelListImportForColumn = lines.JoinArr(vbNewLine)
        GetAllRightModelListImportForColumn = GetGeneratedByFunctionSnippet(GetAllRightModelListImportForColumn, "GetAllRightModelListImportForColumn")
        CopyToClipboard GetAllRightModelListImportForColumn
    End If

End Function

Public Function GetAllUseRightModelListForColumn(frm As Object, Optional SeqModelID = "", Optional ParentModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND NOT ExcludeInRequiredList"
    If Not isFalse(ParentModelID) Then
        sqlStr = sqlStr & " AND RightModelID <> " & ParentModelID
    End If
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetUseRightModelListForColumn(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllUseRightModelListForColumn = lines.JoinArr(vbNewLine)
        GetAllUseRightModelListForColumn = GetGeneratedByFunctionSnippet(GetAllUseRightModelListForColumn, "GetAllUseRightModelListForColumn")
        CopyToClipboard GetAllUseRightModelListForColumn
    End If

End Function
Public Function GetAllRelatedLeftModelImportRoute(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Not IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedLeftModelImportRoute(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedLeftModelImportRoute = lines.JoinArr(vbNewLine)
        GetAllRelatedLeftModelImportRoute = GetGeneratedByFunctionSnippet(GetAllRelatedLeftModelImportRoute, "GetAllRelatedLeftModelImportRoute")
        CopyToClipboard GetAllRelatedLeftModelImportRoute
    End If

End Function


Public Function GetAllRelatedListFromRightRelatedModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID, LeftModelID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND Not IsSimpleRelationship" & _
        " AND Not ExcludeInForm"
    Set rs = ReturnRecordset(sqlStr)
    ''makeQuery sqlStr
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID"): If ExitIfTrue(isFalse(LeftModelID), "LeftModelID is empty..") Then Exit Function
        lines.Add GetAllRelatedRightModelListFromRelatedModelWithParent(frm, LeftModelID, SeqModelID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedListFromRightRelatedModel = lines.JoinArr(vbNewLine)
        GetAllRelatedListFromRightRelatedModel = GetGeneratedByFunctionSnippet(GetAllRelatedListFromRightRelatedModel, "GetAllRelatedListFromRightRelatedModel")
        CopyToClipboard GetAllRelatedListFromRightRelatedModel
    End If

End Function

Public Function GetAllRelatedRightModelListFromRelatedModelWithParent(frm As Object, Optional SeqModelID = "", Optional ParentModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND Not ExcludeInForm AND Not ExcludeInRequiredList"
    
    If Not isFalse(ParentModelID) Then
        sqlStr = sqlStr & " AND RightModelID <> " & ParentModelID
    End If
    
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedRightModelListFromRelatedModelWithParent(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedRightModelListFromRelatedModelWithParent = lines.JoinArr(vbNewLine)
        GetAllRelatedRightModelListFromRelatedModelWithParent = GetGeneratedByFunctionSnippet(GetAllRelatedRightModelListFromRelatedModelWithParent, "GetAllRelatedRightModelListFromRelatedModelWithParent")
        CopyToClipboard GetAllRelatedRightModelListFromRelatedModelWithParent
    End If

End Function

Public Function GetAllRelatedLeftArrayValidation(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND NOT ExcludeInForm AND Not IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedLeftArrayValidation(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedLeftArrayValidation = lines.JoinArr(vbNewLine)
        GetAllRelatedLeftArrayValidation = GetGeneratedByFunctionSnippet(GetAllRelatedLeftArrayValidation, "GetAllRelatedLeftArrayValidation")
        CopyToClipboard GetAllRelatedLeftArrayValidation
    End If

End Function

Public Function GetAllRelatedModelEmptyArraySimpleOnly(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND IsSimpleRelationship"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedModelEmptyArraySimpleOnly(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelEmptyArraySimpleOnly = lines.JoinArr(vbNewLine)
        GetAllRelatedModelEmptyArraySimpleOnly = GetGeneratedByFunctionSnippet(GetAllRelatedModelEmptyArraySimpleOnly, "GetAllRelatedModelEmptyArraySimpleOnly")
        CopyToClipboard GetAllRelatedModelEmptyArraySimpleOnly
    End If

End Function

Public Function GetAllRequiredListForTableForm(frm As Object, Optional SeqModelID = "", Optional ParentModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND Not ExcludeInRequiredList"
    If Not isFalse(ParentModelID) Then
        sqlStr = sqlStr & " AND RightModelID <> " & ParentModelID
    End If

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRequiredListForTableForm(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRequiredListForTableForm = lines.JoinArr(vbNewLine)
        GetAllRequiredListForTableForm = GetGeneratedByFunctionSnippet(GetAllRequiredListForTableForm, "GetAllRequiredListForTableForm")
        CopyToClipboard GetAllRequiredListForTableForm
    End If

End Function

Public Function WriteToMultiRoute_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim IsSupabase: IsSupabase = rs.fields("IsSupabase")
    Dim templateName: templateName = IIf(IsSupabase, "multi routes.ts supabase", "multi route.ts")
    
    WriteToMultiRoute_ts = GetReplacedTemplate(rs, templateName)
    ''WriteToMultiRoute_ts = replace(WriteToMultiRoute_ts, "[GetMultiCreateModelPOSTRoute]", GetMultiCreateModelPOSTRoute(frm, SeqModelID))
    WriteToMultiRoute_ts = GetGeneratedByFunctionSnippet(WriteToMultiRoute_ts, "WriteToMultiRoute_ts", templateName)
    CopyToClipboard WriteToMultiRoute_ts
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\app\api\decks\multi\route.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\app\api\" & ModelPath & "\multi\route.ts"
    WriteToFile filePath, WriteToMultiRoute_ts, SeqModelID, "WriteToMultiRoute_ts"
    
End Function

Public Function WriteToModelmulticreatedeletedialog_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    WriteToModelmulticreatedeletedialog_tsx = GetReplacedTemplate(rs, "ModelMultiCreateDeleteDialog.tsx")
    WriteToModelmulticreatedeletedialog_tsx = GetGeneratedByFunctionSnippet(WriteToModelmulticreatedeletedialog_tsx, "WriteToModelmulticreatedeletedialog_tsx", "ModelMultiCreateDeleteDialog.tsx")
    CopyToClipboard WriteToModelmulticreatedeletedialog_tsx
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\components\decks\DeckDeleteDialog.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "MultiCreateDeleteDialog.tsx"
    WriteToFile filePath, WriteToModelmulticreatedeletedialog_tsx, SeqModelID
    
End Function

Public Function WriteToModellibs_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToModellibs_ts = GetReplacedTemplate(rs, "ModelLibs.ts")
    WriteToModellibs_ts = replace(WriteToModellibs_ts, "[GetAllFieldsToUpdateBySeqModel]", GetAllFieldsToUpdateBySeqModel(frm, SeqModelID))
    WriteToModellibs_ts = replace(WriteToModellibs_ts, "[GetPrimaryKey]", Esc(GetPrimaryKeyField(frm, SeqModelID)))
    WriteToModellibs_ts = replace(WriteToModellibs_ts, "[GetAllOptionalFields]", GetAllOptionalFields(frm, SeqModelID))
    WriteToModellibs_ts = replace(WriteToModellibs_ts, "[GetAllOptionFieldTypes]", GetAllOptionFieldTypes(frm, SeqModelID))
    WriteToModellibs_ts = GetGeneratedByFunctionSnippet(WriteToModellibs_ts, "WriteToModellibs_ts", "ModelLibs.ts")
    CopyToClipboard WriteToModellibs_ts
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13\src\utils\api\HeroLibs.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\utils\api\" & ModelName & "Libs.ts"
    WriteToFile filePath, WriteToModellibs_ts, SeqModelID
    
    
End Function

Public Function GetAllHiddenColumns(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " AND (ControlType = ""Hidden"" OR  HideInTable) AND NOT PrimaryKey"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetHiddenColumns(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllHiddenColumns = lines.JoinArr(vbNewLine)
        GetAllHiddenColumns = GetGeneratedByFunctionSnippet(GetAllHiddenColumns, "GetAllHiddenColumns")
        CopyToClipboard GetAllHiddenColumns
    End If

End Function

Public Function GetAllRelatedDropzoneModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND IncludeAsDropzone ORDER BY SeqModelRelationshipID"

    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedDropzoneModel(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedDropzoneModel = "const {" & vbNewLine & lines.JoinArr(vbNewLine) & vbNewLine & "} = res;"
        GetAllRelatedDropzoneModel = GetGeneratedByFunctionSnippet(GetAllRelatedDropzoneModel, "GetAllRelatedDropzoneModel")
        CopyToClipboard GetAllRelatedDropzoneModel
    End If

End Function

Public Function GetAllRelatedDropzoneModelUpdateOrInsert(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & " AND IncludeAsDropzone ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRelatedDropzoneModelUpdateOrInsert(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedDropzoneModelUpdateOrInsert = lines.JoinArr(vbNewLine)
        GetAllRelatedDropzoneModelUpdateOrInsert = GetGeneratedByFunctionSnippet(GetAllRelatedDropzoneModelUpdateOrInsert, "GetAllRelatedDropzoneModelUpdateOrInsert")
        CopyToClipboard GetAllRelatedDropzoneModelUpdateOrInsert
    End If

End Function

Public Function WriteToModelfiledeletedialog_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToModelfiledeletedialog_tsx = GetReplacedTemplate(rs, "ModelFileDeleteDialog.tsx")
    WriteToModelfiledeletedialog_tsx = GetGeneratedByFunctionSnippet(WriteToModelfiledeletedialog_tsx, "WriteToModelfiledeletedialog_tsx", "ModelFileDeleteDialog.tsx")
    CopyToClipboard WriteToModelfiledeletedialog_tsx
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\components\task-notes\TaskNoteFileDeleteDialog.tsx
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "FileDeleteDialog.tsx"
    WriteToFile filePath, WriteToModelfiledeletedialog_tsx, SeqModelID
    
End Function

Public Function WriteToUsemodelfiledeletedialog_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToUsemodelfiledeletedialog_tsx = GetReplacedTemplate(rs, "useModelFileDeleteDialog.tsx")
    WriteToUsemodelfiledeletedialog_tsx = GetGeneratedByFunctionSnippet(WriteToUsemodelfiledeletedialog_tsx, "WriteToUsemodelfiledeletedialog_tsx", "useModelFileDeleteDialog.tsx")
    CopyToClipboard WriteToUsemodelfiledeletedialog_tsx
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\hooks\task-notes\useTaskNoteFileDeleteDialog.tsx
    Dim filePath: filePath = ClientPath & "src\hooks\" & ModelPath & "\use" & ModelName & "FileDeleteDialog.tsx"
    WriteToFile filePath, WriteToUsemodelfiledeletedialog_tsx, SeqModelID
End Function

Public Function GetAllLeftModelDropzoneImport(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & _
        " AND IncludeAsDropzone ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetLeftModelDropzoneImport(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllLeftModelDropzoneImport = lines.JoinArr(vbNewLine)
        GetAllLeftModelDropzoneImport = GetGeneratedByFunctionSnippet(GetAllLeftModelDropzoneImport, "GetAllLeftModelDropzoneImport")
        CopyToClipboard GetAllLeftModelDropzoneImport
    End If

End Function

Public Function GetAllModelFilesInitial(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & _
        " AND IncludeAsDropzone ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetModelFilesInitial(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFilesInitial = lines.JoinArr(vbNewLine)
        GetAllModelFilesInitial = GetGeneratedByFunctionSnippet(GetAllModelFilesInitial, "GetAllModelFilesInitial")
        CopyToClipboard GetAllModelFilesInitial
    End If

End Function

Public Function GetAllModelFilesInitialModification(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & _
        " AND IncludeAsDropzone ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetModelFilesInitialModification(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelFilesInitialModification = lines.JoinArr(vbNewLine)
        GetAllModelFilesInitialModification = GetGeneratedByFunctionSnippet(GetAllModelFilesInitialModification, "GetAllModelFilesInitialModification")
        CopyToClipboard GetAllModelFilesInitialModification
    End If

End Function

Public Function GetAllModelDropzoneComponent(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & _
        " AND IncludeAsDropzone ORDER BY SeqModelRelationshipID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetModelDropzoneComponent(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllModelDropzoneComponent = lines.JoinArr(vbNewLine)
        GetAllModelDropzoneComponent = GetGeneratedByFunctionSnippet(GetAllModelDropzoneComponent, "GetAllModelDropzoneComponent", , True)
        CopyToClipboard GetAllModelDropzoneComponent
    End If

End Function

Public Function GetAllLeftModelsToReduce(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE RightModelID = " & SeqModelID & _
        " AND NOT ExcludeInTable ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetLeftModelToReduce(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllLeftModelsToReduce = lines.JoinArr(vbNewLine)
        GetAllLeftModelsToReduce = GetGeneratedByFunctionSnippet(GetAllLeftModelsToReduce, "GetAllLeftModelsToReduce")
        CopyToClipboard GetAllLeftModelsToReduce
    End If

End Function

Public Function GetAllRightModelPushPlaceholder(frm As Object, Optional SeqModelID = "", Optional RightModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & _
        " AND RightModelID = " & RightModelID & " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetRightModelPushPlaceholder(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRightModelPushPlaceholder = lines.JoinArr(vbNewLine)
        GetAllRightModelPushPlaceholder = GetGeneratedByFunctionSnippet(GetAllRightModelPushPlaceholder, "GetAllRightModelPushPlaceholder")
        CopyToClipboard GetAllRightModelPushPlaceholder
    End If

End Function

Public Function GetAllPlaceholderModelFromField(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE RelatedModelID = " & SeqModelID & _
        " ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetPlaceholderModelFromField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllPlaceholderModelFromField = lines.JoinArr(vbNewLine)
        GetAllPlaceholderModelFromField = GetGeneratedByFunctionSnippet(GetAllPlaceholderModelFromField, "GetAllPlaceholderModelFromField")
        CopyToClipboard GetAllPlaceholderModelFromField
    End If

End Function

Public Function GetAllRelatedModelListFromField(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT RelatedModelID,SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND NOT RelatedModelID IS NULL ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim RelatedModelIDs As New clsArray
    
    Do Until rs.EOF
        Dim RelatedModelID: RelatedModelID = Trim(rs.fields("RelatedModelID"))
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        If Not RelatedModelIDs.InArray(RelatedModelID) Then
            lines.Add GetRelatedModelListFromField(frm, SeqModelFieldID)
            RelatedModelIDs.Add RelatedModelID
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRelatedModelListFromField = lines.JoinArr(vbNewLine)
        GetAllRelatedModelListFromField = GetGeneratedByFunctionSnippet(GetAllRelatedModelListFromField, "GetAllRelatedModelListFromField")
        CopyToClipboard GetAllRelatedModelListFromField
    End If

End Function

Public Function GetAllRequiredListImportForRelatedModelFromField(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT RelatedModelID,SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND Not RelatedModelID IS NULL ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim RelatedModelIDs As New clsArray
    
    Do Until rs.EOF
        Dim RelatedModelID: RelatedModelID = Trim(rs.fields("RelatedModelID")): If ExitIfTrue(isFalse(RelatedModelID), "RelatedModelID is empty..") Then Exit Function
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        If Not RelatedModelIDs.InArray(RelatedModelID) Then
            lines.Add GetRequiredListImportForRelatedModelFromField(frm, SeqModelFieldID)
            GetuseModelListts frm, RelatedModelID
            RelatedModelIDs.Add RelatedModelID
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllRequiredListImportForRelatedModelFromField = lines.JoinArr(vbNewLine)
        GetAllRequiredListImportForRelatedModelFromField = GetGeneratedByFunctionSnippet(GetAllRequiredListImportForRelatedModelFromField, "GetAllRequiredListImportForRelatedModelFromField")
        CopyToClipboard GetAllRequiredListImportForRelatedModelFromField
    End If

End Function

Public Function GetAllOptionalFields(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND (AllowNull OR DataType = ""Boolean"") ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetOptionalField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllOptionalFields = lines.JoinArr(vbNewLine)
        GetAllOptionalFields = GetGeneratedByFunctionSnippet(GetAllOptionalFields, "GetAllOptionalFields")
        CopyToClipboard GetAllOptionalFields
    End If

End Function

Public Function GetAllOptionFieldTypes(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND (AllowNull OR DataType = ""Boolean"") ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetModelFieldType(frm, SeqModelFieldID, , True)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllOptionFieldTypes = lines.JoinArr(vbNewLine)
        GetAllOptionFieldTypes = GetGeneratedByFunctionSnippet(GetAllOptionFieldTypes, "GetAllOptionFieldTypes")
        CopyToClipboard GetAllOptionFieldTypes
    End If

End Function

Public Function GetCodeOriginallyFromModelTable(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetCodeOriginallyFromModelTable = GetReplacedTemplate(rs, "GetCodeOriginallyFromModelTable")
    GetCodeOriginallyFromModelTable = replace(GetCodeOriginallyFromModelTable, "[GetAllQueryKeyValueOfGetPluralizedModelName]", GetAllQueryKeyValueOfGetPluralizedModelName(frm, SeqModelID))
    GetCodeOriginallyFromModelTable = replace(GetCodeOriginallyFromModelTable, "[GetAllFilterQueryNameBySeqModel]", GetAllFilterQueryNameBySeqModel(frm, SeqModelID))
    GetCodeOriginallyFromModelTable = GetGeneratedByFunctionSnippet(GetCodeOriginallyFromModelTable, "GetCodeOriginallyFromModelTable", "GetCodeOriginallyFromModelTable")
    CopyToClipboard GetCodeOriginallyFromModelTable
    
End Function

Public Function GetAllQueryKeyValueOfGetPluralizedModelName(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim uqFilterQueryNames As New clsArray
   
    sqlStr = "SELECT SeqModelFilterID, FilterQueryName FROM tblSeqModelFilters WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelFilterID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName"): If ExitIfTrue(isFalse(FilterQueryName), "FilterQueryName is empty..") Then Exit Function
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        If Not uqFilterQueryNames.InArray(FilterQueryName) Then
            lines.Add GetQueryKeyValueOfGetPluralizedModelName(frm, SeqModelFilterID)
            uqFilterQueryNames.Add "FilterQueryName", True
        End If
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllQueryKeyValueOfGetPluralizedModelName = lines.JoinArr(vbNewLine)
        GetAllQueryKeyValueOfGetPluralizedModelName = GetGeneratedByFunctionSnippet(GetAllQueryKeyValueOfGetPluralizedModelName, "GetAllQueryKeyValueOfGetPluralizedModelName")
        CopyToClipboard GetAllQueryKeyValueOfGetPluralizedModelName
    End If

End Function

Public Function WriteToUsemodelnametable_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToUsemodelnametable_tsx = GetReplacedTemplate(rs, "useModelNameTable.tsx")
    WriteToUsemodelnametable_tsx = replace(WriteToUsemodelnametable_tsx, "[GetAllSearchParamsBySeqModel]", GetAllSearchParamsBySeqModel(frm, SeqModelID))
    WriteToUsemodelnametable_tsx = replace(WriteToUsemodelnametable_tsx, "[GetAllFilterQueryNameBySeqModel]", GetAllFilterQueryNameBySeqModel(frm, SeqModelID))
    WriteToUsemodelnametable_tsx = GetGeneratedByFunctionSnippet(WriteToUsemodelnametable_tsx, "WriteToUsemodelnametable_tsx", "useModelNameTable.tsx")
    CopyToClipboard WriteToUsemodelnametable_tsx
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\hooks\tasks\useTaskTable.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\hooks\" & ModelPath & "\use" & ModelName & "Table.ts"
    WriteToFile filePath, WriteToUsemodelnametable_tsx, SeqModelID
    
End Function

Public Function WriteToUsemodelpageparams_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToUsemodelpageparams_tsx = GetReplacedTemplate(rs, "useModelPageParams.tsx")
    WriteToUsemodelpageparams_tsx = replace(WriteToUsemodelpageparams_tsx, "[GetAllSearchParamsBySeqModel]", GetAllSearchParamsBySeqModel(frm, SeqModelID))
    WriteToUsemodelpageparams_tsx = replace(WriteToUsemodelpageparams_tsx, "[GetAllFilterQueryNameBySeqModel]", GetAllFilterQueryNameBySeqModel(frm, SeqModelID))
    WriteToUsemodelpageparams_tsx = GetGeneratedByFunctionSnippet(WriteToUsemodelpageparams_tsx, "WriteToUsemodelpageparams_tsx", "useModelPageParams.tsx")
    CopyToClipboard WriteToUsemodelpageparams_tsx
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\hooks\tasks\useTaskPageParams.tsx
    Dim filePath: filePath = ClientPath & "src\hooks\" & ModelPath & "\use" & ModelName & "PageParams.tsx"
    WriteToFile filePath, WriteToUsemodelpageparams_tsx, SeqModelID
    
    
End Function

Public Function GetItemIcons(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim NavItemIcon: NavItemIcon = rs.fields("NavItemIcon"): If ExitIfTrue(isFalse(NavItemIcon), "NavItemIcon is empty..") Then Exit Function
    
    GetItemIcons = NavItemIcon
    CopyToClipboard GetItemIcons
    
End Function

Public Function WriteToModelconfig_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''Update the ColumnsOccupied of the tblSeqModelFields if null
    RunSQL "UPDATE tblSeqModelFields SET ColumnsOccupied = 12 WHERE ColumnsOccupied IS NULL AND SeqModelID = " & SeqModelID
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), "BackendProjectID is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "Model Name is empty..") Then Exit Function
    Dim NavItemIcon: NavItemIcon = rs.fields("NavItemIcon")
    
    Dim vNavItemIcon: vNavItemIcon = IIf(Not isFalse(NavItemIcon), "import { " & NavItemIcon & " } from ""lucide-react""", "")
    WriteToModelconfig_ts = GetReplacedTemplate(rs, "ModelConfig.ts")
    WriteToModelconfig_ts = replace(WriteToModelconfig_ts, "[vNavItemIcon]", vNavItemIcon)
    WriteToModelconfig_ts = replace(WriteToModelconfig_ts, "[GetSeqModelKeys]", GetSeqModelKeys(frm, SeqModelID))
    WriteToModelconfig_ts = replace(WriteToModelconfig_ts, "[GetAllSeqModelFieldKeys]", GetAllSeqModelFieldKeys(frm, SeqModelID))
    WriteToModelconfig_ts = replace(WriteToModelconfig_ts, "[GetAllSeqModelFilterKeys]", GetAllSeqModelFilterKeys(frm, SeqModelID))
    WriteToModelconfig_ts = replace(WriteToModelconfig_ts, "[GetAllSeqModelSortKeys]", GetAllSeqModelSortKeys(frm, SeqModelID))
    WriteToModelconfig_ts = replace(WriteToModelconfig_ts, "[GetAllSeqModelHookKeys]", GetAllSeqModelHookKeys(frm, SeqModelID))
    WriteToModelconfig_ts = replace(WriteToModelconfig_ts, "[GetAllSeqModelFieldGroups]", GetAllSeqModelFieldGroups(frm, SeqModelID))
    WriteToModelconfig_ts = replace(WriteToModelconfig_ts, "[GetAllSeqModelEmbeddings]", GetAllSeqModelEmbeddings(frm, SeqModelID))
    
    WriteToModelconfig_ts = replace(WriteToModelconfig_ts, "fieldValue: Null,", "fieldValue: ""null"",")
    
    WriteToModelconfig_ts = GetGeneratedByFunctionSnippet(WriteToModelconfig_ts, "WriteToModelconfig_ts", "ModelConfig.ts")
    CopyToClipboard WriteToModelconfig_ts
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\utils\config\TaskConfig.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\utils\config\" & ModelName & "Config.ts"
    WriteToFile filePath, WriteToModelconfig_ts, SeqModelID, "WriteToModelconfig_ts"
    
    ''WriteToModelconfig_tsInterface frm, BackendProjectID
    ''WriteToAppconfig_ts frm, BackendProjectID
    
End Function

Public Function GetSeqModelKeys(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GetSeqModelKeys = GetKVPairs("qrySeqModels", rs)
    GetSeqModelKeys = GetGeneratedByFunctionSnippet(GetSeqModelKeys, "GetSeqmodelkeys", "")
    
End Function


Public Function GetAllSeqModelFieldKeys(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetSeqModelFieldKeys(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSeqModelFieldKeys = lines.JoinArr(vbNewLine)
        GetAllSeqModelFieldKeys = GetGeneratedByFunctionSnippet(GetAllSeqModelFieldKeys, "GetAllSeqModelFieldKeys")
        CopyToClipboard GetAllSeqModelFieldKeys
    End If

End Function

Public Function GetAllSeqModelFilterKeys(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFilterID FROM tblSeqModelFilters WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY FilterOrder,SeqModelFilterID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetSeqModelFilterKeys(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSeqModelFilterKeys = lines.JoinArr(vbNewLine)
        GetAllSeqModelFilterKeys = GetGeneratedByFunctionSnippet(GetAllSeqModelFilterKeys, "GetAllSeqModelFilterKeys")
        CopyToClipboard GetAllSeqModelFilterKeys
    End If

End Function

Public Function GetAllSeqModelSortKeys(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelSortID FROM tblSeqModelSorts WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelSortID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelSortID: SeqModelSortID = rs.fields("SeqModelSortID")
        lines.Add GetSeqModelSortKeys(frm, SeqModelSortID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSeqModelSortKeys = lines.JoinArr(vbNewLine)
        GetAllSeqModelSortKeys = GetGeneratedByFunctionSnippet(GetAllSeqModelSortKeys, "GetAllSeqModelSortKeys")
        CopyToClipboard GetAllSeqModelSortKeys
    End If

End Function

Public Function GetAllSeqModelHookKeys(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelHookID FROM tblSeqModelHooks WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelHookID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelHookID: SeqModelHookID = rs.fields("SeqModelHookID")
        lines.Add GetSeqModelHookKeys(frm, SeqModelHookID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSeqModelHookKeys = lines.JoinArr(vbNewLine)
        GetAllSeqModelHookKeys = GetGeneratedByFunctionSnippet(GetAllSeqModelHookKeys, "GetAllSeqModelHookKeys")
        CopyToClipboard GetAllSeqModelHookKeys
    End If

End Function

Public Function GetAllSeqModelFieldGroups(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldGroupID FROM tblSeqModelFieldGroups WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelFieldGroupID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldGroupID: SeqModelFieldGroupID = rs.fields("SeqModelFieldGroupID")
        lines.Add GetSeqModelFieldGroups(frm, SeqModelFieldGroupID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSeqModelFieldGroups = lines.JoinArr(vbNewLine)
        GetAllSeqModelFieldGroups = GetGeneratedByFunctionSnippet(GetAllSeqModelFieldGroups, "GetAllSeqModelFieldGroups")
        CopyToClipboard GetAllSeqModelFieldGroups
    End If

End Function

Public Function GetAllSeqModelEmbeddings(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelEmbeddingID FROM tblSeqModelEmbeddings WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelEmbeddingID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelEmbeddingID: SeqModelEmbeddingID = rs.fields("SeqModelEmbeddingID")
        lines.Add GetSeqModelEmbeddings(frm, SeqModelEmbeddingID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllSeqModelEmbeddings = lines.JoinArr(vbNewLine)
        GetAllSeqModelEmbeddings = GetGeneratedByFunctionSnippet(GetAllSeqModelEmbeddings, "GetAllSeqModelEmbeddings")
        CopyToClipboard GetAllSeqModelEmbeddings
    End If

End Function


Public Function GetModelConfigImports(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelConfigImports = GetReplacedTemplate(rs, "GetModelConfigImports")
    GetModelConfigImports = GetGeneratedByFunctionSnippet(GetModelConfigImports, "GetModelConfigImports", "GetModelConfigImports", , True)
    CopyToClipboard GetModelConfigImports
    
End Function

Public Function WriteToModelsinglecolumn_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    Dim IsTable: IsTable = rs.fields("IsTable")
    
    WriteToModelsinglecolumn_tsx = GetReplacedTemplate(rs, "ModelSingleColumn.tsx")
    
    Dim GetFormikShapeOrModel: GetFormikShapeOrModel = IIf(IsTable, "Model", "FormikShape")
    
    WriteToModelsinglecolumn_tsx = replace(WriteToModelsinglecolumn_tsx, "[GetFormikShapeOrModel]", GetFormikShapeOrModel)
    WriteToModelsinglecolumn_tsx = GetGeneratedByFunctionSnippet(WriteToModelsinglecolumn_tsx, "WriteToModelsinglecolumn_tsx", "ModelSingleColumn.tsx")
    CopyToClipboard WriteToModelsinglecolumn_tsx
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\components\tasks\TaskSingleColumn.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "SingleColumn.tsx"
    WriteToFile filePath, WriteToModelsinglecolumn_tsx, SeqModelID, "WriteToModelsinglecolumn_tsx"
    
End Function

Public Function GetBackendModelImports(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetBackendModelImports = GetReplacedTemplate(rs, "GetBackendModelImports")
    GetBackendModelImports = GetGeneratedByFunctionSnippet(GetBackendModelImports, "GetBackendModelImports", "GetBackendModelImports", , True)
    CopyToClipboard GetBackendModelImports
    
End Function
Public Function GetBackedModelName(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetBackedModelName = GetReplacedTemplate(rs, "GetBackedModelName")
    GetBackedModelName = GetGeneratedByFunctionSnippet(GetBackedModelName, "GetBackedModelName", "GetBackedModelName", , True)
    CopyToClipboard GetBackedModelName
End Function

Public Function WriteToModelsRoutes_tsUsingModelconfig(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim UsePostgREST: UsePostgREST = rs.fields("UsePostgREST")
    Dim IsSupabase: IsSupabase = rs.fields("IsSupabase")
    
    Dim templateName: templateName = IIf(IsSupabase, "models routes.ts supabase", "models routes.ts using modelConfig")
    
    If UsePostgREST Then
        templateName = "models routes.ts postgREST"
    End If
    
    If isPresent("qrySeqModelEmbeddings", "SeqModelID = " & SeqModelID) Then
        templateName = "models routes.ts using postgreSQL"
    End If
    
    WriteToModelsRoutes_tsUsingModelconfig = GetReplacedTemplate(rs, templateName)
    WriteToModelsRoutes_tsUsingModelconfig = GetGeneratedByFunctionSnippet(WriteToModelsRoutes_tsUsingModelconfig, "WriteToModelsRoutes_tsUsingModelconfig", templateName)
    CopyToClipboard WriteToModelsRoutes_tsUsingModelconfig
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\app\api\tasks\route.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "Project Path is empty..") Then Exit Function
    Dim filePath: filePath = ProjectPath & "src\app\api\" & ModelPath & "\route.ts"
    WriteToFile filePath, WriteToModelsRoutes_tsUsingModelconfig, SeqModelID, "WriteToModelsRoutes_tsUsingModelconfig"

End Function
Public Function WriteToModelRoute_tsUsingModelconfig(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim IsSupabase: IsSupabase = rs.fields("IsSupabase")
    
    Dim templateName: templateName = IIf(IsSupabase, "model route supabase", "model route.ts using modelConfig")
    WriteToModelRoute_tsUsingModelconfig = GetReplacedTemplate(rs, templateName)
    WriteToModelRoute_tsUsingModelconfig = GetGeneratedByFunctionSnippet(WriteToModelRoute_tsUsingModelconfig, "WriteToModelRoute_tsUsingModelconfig", templateName)
    CopyToClipboard WriteToModelRoute_tsUsingModelconfig
    
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\app\api\tasks\[id]\route.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "Project Path is empty..") Then Exit Function
    Dim filePath: filePath = ProjectPath & "src\app\api\" & ModelPath & "\[id]\route.ts"
    WriteToFile filePath, WriteToModelRoute_tsUsingModelconfig, SeqModelID, "WriteToModelRoute_tsUsingModelconfig"
End Function

Public Function WriteToModelform_tsxUsingModelconfig(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToModelform_tsxUsingModelconfig = GetReplacedTemplate(rs, "ModelForm.tsx using modelConfig", "SlugField")
    WriteToModelform_tsxUsingModelconfig = GetGeneratedByFunctionSnippet(WriteToModelform_tsxUsingModelconfig, "WriteToModelform_tsxUsingModelconfig", "ModelForm.tsx using modelConfig")
    CopyToClipboard WriteToModelform_tsxUsingModelconfig
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\components\tasks\TaskForm.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "Form.tsx"
    WriteToFile filePath, WriteToModelform_tsxUsingModelconfig, SeqModelID, "WriteToModelform_tsxUsingModelconfig"
    
End Function

Public Function WriteToGetmodelrowaction_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetmodelrowaction_tsx = GetReplacedTemplate(rs, "getModelRowAction.tsx")
    WriteToGetmodelrowaction_tsx = GetGeneratedByFunctionSnippet(WriteToGetmodelrowaction_tsx, "WriteToGetmodelrowaction_tsx", "getModelRowAction.tsx")
    CopyToClipboard WriteToGetmodelrowaction_tsx
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\lib\tasks\getTaskRowActions.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\get" & ModelName & "RowActions.tsx"
    WriteToFile filePath, WriteToGetmodelrowaction_tsx, SeqModelID, "WriteToGetmodelrowaction_tsx"
    
End Function

Public Function WriteToGetmodelcolumnstobeoverriden_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetmodelcolumnstobeoverriden_tsx = GetReplacedTemplate(rs, "getModelColumnsToBeOverriden.tsx")
    WriteToGetmodelcolumnstobeoverriden_tsx = GetGeneratedByFunctionSnippet(WriteToGetmodelcolumnstobeoverriden_tsx, "WriteToGetmodelcolumnstobeoverriden_tsx", "getModelColumnsToBeOverriden.tsx")
    CopyToClipboard WriteToGetmodelcolumnstobeoverriden_tsx
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    ''C:\Users\User\Desktop\Web Development\task-manager-next-13\src\lib\tasks\getTaskColumnsToBeOverriden.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\get" & ModelName & "ColumnsToBeOverriden.tsx"
    WriteToFile filePath, WriteToGetmodelcolumnstobeoverriden_tsx, SeqModelID, "WriteToGetmodelcolumnstobeoverriden_tsx"
    
End Function

Public Function WriteToModelstateholder_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToModelstateholder_tsx = GetReplacedTemplate(rs, "ModelStateHolder.tsx")
    WriteToModelstateholder_tsx = GetGeneratedByFunctionSnippet(WriteToModelstateholder_tsx, "WriteToModelstateholder_tsx", "ModelStateHolder.tsx")
    CopyToClipboard WriteToModelstateholder_tsx
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    ''C:\Users\User\Desktop\Web Development\personal-finance-next-13\src\components\account-titles\AccountTitleStateHolder.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "StateHolder.tsx"
    WriteToFile filePath, WriteToModelstateholder_tsx, SeqModelID, "WriteToModelstateholder_tsx"
    
End Function

Public Function GetLucideIcon(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetLucideIcon = GetReplacedTemplate(rs, "GetLucideIcon")
    GetLucideIcon = GetGeneratedByFunctionSnippet(GetLucideIcon, "GetLucideIcon", "GetLucideIcon")
    CopyToClipboard GetLucideIcon
    
End Function
Public Function GetLucideIconItem(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetLucideIconItem = GetReplacedTemplate(rs, "GetLucideIcon")
    GetLucideIconItem = GetGeneratedByFunctionSnippet(GetLucideIconItem, "GetLucideIconItem", "GetLucideIcon")
    CopyToClipboard GetLucideIconItem
End Function


Public Function GetComponentImport(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName")
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected, rs2 As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblFunctionChainItems"
          .AddFilter "tblFunctionChainItems.Note Like " & Esc("*Special File*")
          .AddFilter "NOT FilePathTemplate IS NULL"
          .AddFilter "NOT ModelOptionImportStatement IS NULL"
          .fields = "FilePathTemplate,ModelOptionImportStatement"
          .joins.Add GenerateJoinObj("tblModelButtons", "ModelButtonID")
          .OrderBy = "FunctionOrder,FunctionChainItemID"
          Set rs2 = .Recordset
    End With
    
    
    
    Do Until rs2.EOF
        Dim FilePathTemplate: FilePathTemplate = rs2.fields("FilePathTemplate")
        Dim filePath: filePath = GetReplacedTemplate(rs, "", , FilePathTemplate)
        Dim ModelOptionImportStatement: ModelOptionImportStatement = GetReplacedTemplate(rs, "", , rs2.fields("ModelOptionImportStatement"))
        
        Dim ShouldImport: ShouldImport = isPresent("tblSeqModelFiles", "filePath = " & Esc(filePath) & " AND SeqModelID = " & _
            SeqModelID & " AND IsProtected")
        If ShouldImport Then
            lines.Add ModelOptionImportStatement
        End If
        rs2.MoveNext
    Loop
    
    GetComponentImport = ""
    If lines.count > 0 Then
        GetComponentImport = lines.JoinArr(vbNewLine)
        GetComponentImport = GetGeneratedByFunctionSnippet(GetComponentImport, "GetComponentImport")
        CopyToClipboard GetComponentImport
    End If
    
    CopyToClipboard GetComponentImport
    
End Function
Public Function GetModelSingleColumn(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelSingleColumn = GetReplacedTemplate(rs, "GetModelSingleColumn")
    GetModelSingleColumn = GetGeneratedByFunctionSnippet(GetModelSingleColumn, "GetModelSingleColumn", "GetModelSingleColumn")
    CopyToClipboard GetModelSingleColumn
End Function

Public Function GetBasicSeqModelReplacement(frm As Object, Optional SeqModelID = "", Optional templateName = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetBasicSeqModelReplacement = GetReplacedTemplate(rs, templateName)
    GetBasicSeqModelReplacement = GetGeneratedByFunctionSnippet(GetBasicSeqModelReplacement, "GetBasicSeqModelReplacement", templateName)
    CopyToClipboard GetBasicSeqModelReplacement
    
End Function

Public Function GetModelForm(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelForm = GetReplacedTemplate(rs, "GetModelForm")
    GetModelForm = GetGeneratedByFunctionSnippet(GetModelForm, "GetModelForm", "GetModelForm")
    CopyToClipboard GetModelForm
End Function

Public Function GetPostgresqlCreateTableStatement(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetPostgresqlCreateTableStatement = GetReplacedTemplate(rs, "PostgreSQL CREATE TABLE Statement")
    GetPostgresqlCreateTableStatement = replace(GetPostgresqlCreateTableStatement, "[GetAllPostgreSQLCreateTableFields]", GetAllPostgreSQLCreateTableFields(frm, SeqModelID))
    GetPostgresqlCreateTableStatement = replace(GetPostgresqlCreateTableStatement, "[GetAllPostgreSQLCreateTableRelationships]", GetAllPostgreSQLCreateTableRelationships(frm, SeqModelID))
    GetPostgresqlCreateTableStatement = replace(GetPostgresqlCreateTableStatement, "[GetAllPostgreSQLCreateUniqueIndexes]", GetAllPostgreSQLCreateUniqueIndexes(frm, SeqModelID))
    GetPostgresqlCreateTableStatement = replace(GetPostgresqlCreateTableStatement, "[GetCreatedAndUpdatedIndex]", GetCreatedAndUpdatedIndex(frm, SeqModelID))
    GetPostgresqlCreateTableStatement = replace(GetPostgresqlCreateTableStatement, "[GetDeleteTrigger]", GetDeleteTrigger(frm, SeqModelID))
    ''GetPostgresqlCreateTableStatement = GetGeneratedByFunctionSnippet(GetPostgresqlCreateTableStatement, "GetPostgresqlCreateTableStatement", "PostgreSQL CREATE TABLE Statement")
    CopyToClipboard GetPostgresqlCreateTableStatement
    
    DoCmd.OpenForm "frmClipboardForms"
    Forms("frmClipboardForms")("Snippet") = GetPostgresqlCreateTableStatement
    CopyFieldContent Forms("frmClipboardForms"), "Snippet"
    
    OpenSupabaseSqlEditorThenClose_frmClipboardForms frm, SeqModelID
    
End Function

Public Function GetAllPostgreSQLCreateTableFields(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim Timestamps: Timestamps = rs.fields("Timestamps")
    Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Function
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND Expression IS NULL AND NOT IsGeneratedField ORDER BY FieldOrder"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetPostgreSQLCreateTableField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If Not isFalse(SlugField) Then
        lines.Add """slug"" citext NOT NULL UNIQUE GENERATED ALWAYS AS (slugify(" & SlugField & ")) STORED,"
    End If
    
    If Timestamps Then
        lines.Add """created_at"" TIMESTAMP DEFAULT current_timestamp,"
        lines.Add """updated_at"" TIMESTAMP DEFAULT NULL,"
    End If
    
    If isPresent("qrySeqModelEmbeddings", "SeqModelID = " & SeqModelID) Then
        lines.Add """embedding"" vector(1536),"
    End If
    
    If lines.count > 0 Then
        GetAllPostgreSQLCreateTableFields = lines.JoinArr(vbNewLine)
        If Not isPresent("tblSeqModelRelationships", "LeftModelID = " & SeqModelID) Then
            GetAllPostgreSQLCreateTableFields = Left(GetAllPostgreSQLCreateTableFields, Len(GetAllPostgreSQLCreateTableFields) - 1)
        End If
        ''GetAllPostgreSQLCreateTableFields = GetGeneratedByFunctionSnippet(GetAllPostgreSQLCreateTableFields, "GetAllPostgreSQLCreateTableFields")
        CopyToClipboard GetAllPostgreSQLCreateTableFields
    End If

End Function

Public Function GetAllPostgreSQLCreateTableRelationships(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & _
        " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetPostgreSQLCreateTableRelationship(frm, SeqModelRelationshipID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllPostgreSQLCreateTableRelationships = lines.JoinArr(vbNewLine)
        GetAllPostgreSQLCreateTableRelationships = Left(GetAllPostgreSQLCreateTableRelationships, Len(GetAllPostgreSQLCreateTableRelationships) - 1)
        ''GetAllPostgreSQLCreateTableRelationships = GetGeneratedByFunctionSnippet(GetAllPostgreSQLCreateTableRelationships, "GetAllPostgreSQLCreateTableRelationships")
        CopyToClipboard GetAllPostgreSQLCreateTableRelationships
    End If

End Function

Public Function GetDROPStatementForPostgreSQL(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetDROPStatementForPostgreSQL = GetReplacedTemplate(rs, "DROP postgreSQL table")
    ''GetDROPStatementForPostgreSQL = GetGeneratedByFunctionSnippet(GetDROPStatementForPostgreSQL, "GetDROPStatementForPostgreSQL", "DROP postgreSQL table")
    CopyToClipboard GetDROPStatementForPostgreSQL
    
End Function

Public Function GetPostgrePlainCreateStatement(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetPostgrePlainCreateStatement = GetReplacedTemplate(rs, "GetPostgrePlainCreateStatement")
    GetPostgrePlainCreateStatement = replace(GetPostgrePlainCreateStatement, "[GetAllPostgreSQLCreateTableFields]", GetAllPostgreSQLCreateTableFields(frm, SeqModelID))
    GetPostgrePlainCreateStatement = replace(GetPostgrePlainCreateStatement, "[GetAllPostgreSQLCreateTableRelationships]", GetAllPostgreSQLCreateTableRelationships(frm, SeqModelID))
    ''GetAllPostgreSQLCreateTableFields --> ''GetAllPostgreSQLCreateTableRelationships
    
    ''GetPostgrePlainCreateStatement = GetGeneratedByFunctionSnippet(GetPostgrePlainCreateStatement, "GetPostgrePlainCreateStatement", "GetPostgrePlainCreateStatement")
    CopyToClipboard GetPostgrePlainCreateStatement
    
End Function
Public Function GetDeletePostgresql(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetDeletePostgresql = GetReplacedTemplate(rs, "GetDeletePostgresql")
    ''GetDeletePostgresql = GetGeneratedByFunctionSnippet(GetDeletePostgresql, "GetDeletePostgresql", "GetDeletePostgresql")
    CopyToClipboard GetDeletePostgresql
End Function
Public Function GetAlterPolicyStatements(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetAlterPolicyStatements = GetReplacedTemplate(rs, "GetAlterPolicyStatements")
    ''GetAlterPolicyStatements = GetGeneratedByFunctionSnippet(GetAlterPolicyStatements, "GetAlterPolicyStatements", "GetAlterPolicyStatements")
    CopyToClipboard GetAlterPolicyStatements
End Function

Public Function ExportToCsvSqlStatement(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''qryTableToExportToCSV
    Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Function
    Dim fields: fields = GetFieldsToBeExported(frm, SeqModelID)
    
    Dim xlPath: xlPath = CurrentProject.path & "\" & TableName & ".xlsx"
    On Error Resume Next
    DoCmd.DeleteObject acTable, "tblToBeExported"
    sqlStr = "SELECT " & fields & " INTO tblToBeExported FROM " & TableName
    RunSQL sqlStr
    DoCmd.TransferSpreadsheet acExport, , "tblToBeExported", xlPath, True
    
    CopyExcelDataFromAccess xlPath
    
End Function

Private Sub CopyExcelDataFromAccess(xlPath)
    Dim xlApp As Object
    Dim xlWorkBook As Object
    Dim xlWorkSheet As Object
    Dim UsedRange As Object
    
    ' Initialize Excel Application
    Set xlApp = CreateObject("Excel.Application")
    
    ' Make Excel visible
    xlApp.Visible = True
    
    ' Open the Excel file
    Set xlWorkBook = xlApp.Workbooks.Open(xlPath)
    
    ' Set the worksheet to the first sheet
    Set xlWorkSheet = xlWorkBook.Sheets(1)
    
    ' Select the used range in the worksheet
    Set UsedRange = xlWorkSheet.UsedRange
    
    ' Copy the selected range
    UsedRange.Copy
    
    ' Close the Excel file (you can change this to save the file if needed)
    'xlWorkBook.Close False
    
    ' Release Excel objects
    'Set UsedRange = Nothing
    'Set xlWorkSheet = Nothing
    'Set xlWorkBook = Nothing
    
    ' Quit Excel Application
    'xlApp.Quit
    'Set xlApp = Nothing
End Sub


Public Function GetFieldsToBeExported(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    
    sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        Dim AllowNull: AllowNull = rs.fields("AllowNull")
        Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
        Dim TableName: TableName = rs.fields("TableName")
        If AllowNull Then
            lines.Add "nz([" & TableName & "].[" & DatabaseFieldName & "],""NULL"") AS " & DatabaseFieldName
        Else
            lines.Add DatabaseFieldName
        End If
        
        rs.MoveNext
    Loop
    
    If Not isFalse(SlugField) Then
        lines.Add "slug"
    End If
    
    If lines.count > 0 Then
        GetFieldsToBeExported = lines.JoinArr(",")
        ''GetFieldsToBeExported = GetGeneratedByFunctionSnippet(GetFieldsToBeExported, "GetFieldsToBeExported")
        CopyToClipboard GetFieldsToBeExported
    End If

End Function


Public Function GetResetSerialAutonumber(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetResetSerialAutonumber = GetReplacedTemplate(rs, "Reset Serial autonumber")
    ''GetResetSerialAutonumber = GetGeneratedByFunctionSnippet(GetResetSerialAutonumber, "GetResetSerialAutonumber", "Reset Serial autonumber")
    CopyToClipboard GetResetSerialAutonumber
    
End Function

Public Function GetAllIndextStatementByFilter(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFilterID FROM tblSeqModelFilters WHERE SeqModelID = " & SeqModelID & _
        " AND NOT SeqModelFieldID IS NULL AND FilterQueryName <> ""q"" ORDER BY SeqModelFilterID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        lines.Add GetIndextStatementByFilter(frm, SeqModelFilterID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllIndextStatementByFilter = lines.JoinArr(vbNewLine)
        ''GetAllIndextStatementByFilter = GetGeneratedByFunctionSnippet(GetAllIndextStatementByFilter, "GetAllIndextStatementByFilter")
        CopyToClipboard GetAllIndextStatementByFilter
    End If

End Function

Public Function WriteToHookPostRoutePerModel(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelHookID FROM tblSeqModelHooks WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelHookID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelHookID: SeqModelHookID = rs.fields("SeqModelHookID")
        lines.Add WriteToHookPostRoute(frm, SeqModelHookID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        WriteToHookPostRoutePerModel = lines.JoinArr(vbNewLine)
        WriteToHookPostRoutePerModel = GetGeneratedByFunctionSnippet(WriteToHookPostRoutePerModel, "WriteToHookPostRoutePerModel")
        CopyToClipboard WriteToHookPostRoutePerModel
    End If

End Function

Public Function AddSlugColumnSql(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    AddSlugColumnSql = GetReplacedTemplate(rs, "Add Slug Column")
    ''AddSlugColumnSql = GetGeneratedByFunctionSnippet(AddSlugColumnSql, "AddSlugColumnSql", "Add Slug Column")
    CopyToClipboard AddSlugColumnSql
    
End Function

Public Function WriteToAfterupdatecomponent_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function

    WriteToAfterupdatecomponent_tsx = GetReplacedTemplate(rs, "AfterUpdateComponent.tsx")
    WriteToAfterupdatecomponent_tsx = GetGeneratedByFunctionSnippet(WriteToAfterupdatecomponent_tsx, "WriteToAfterupdatecomponent_tsx", "AfterUpdateComponent.tsx")
    CopyToClipboard WriteToAfterupdatecomponent_tsx
    
    ''C:\Users\User\Desktop\Web Development\personal-finance-next-13\src\components\journal-entries\AddOrUpdateCashInWallet.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\" & ModelName & "AfterUpdateComponent.ts"
    WriteToFile filePath, WriteToAfterupdatecomponent_tsx, SeqModelID, "WriteToAfterupdatecomponent_tsx"

End Function

Public Function WriteToValidatemodelname_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    WriteToValidatemodelname_tsx = GetReplacedTemplate(rs, "validateModelName.tsx")
    WriteToValidatemodelname_tsx = GetGeneratedByFunctionSnippet(WriteToValidatemodelname_tsx, "WriteToValidatemodelname_tsx", "validateModelName.tsx")
    CopyToClipboard WriteToValidatemodelname_tsx
    
    ''C:\Users\User\Desktop\Web Development\personal-finance-next-13\src\lib\validation\validateJournalEntry.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\validation\validate" & ModelName & ".ts"
    WriteToFile filePath, WriteToValidatemodelname_tsx, SeqModelID, "WriteToValidatemodelname_tsx"
    
End Function

Public Function AddForeignKeyConstraintPostgres(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim replacement: replacement = GetAllPostgreSQLCreateTableRelationships(frm, SeqModelID)
    replacement = replace(replacement, "CONSTRAINT", "ADD CONSTRAINT")
    AddForeignKeyConstraintPostgres = GetReplacedTemplate(rs, "PostgreSQL add foreign key constraints")
    AddForeignKeyConstraintPostgres = replace(AddForeignKeyConstraintPostgres, "[GetAllPostgreSQLCreateTableRelationships]", replacement)
    ''AddForeignKeyConstraintPostgres = GetGeneratedByFunctionSnippet(AddForeignKeyConstraintPostgres, "AddForeignKeyConstraintPostgres", "PostgreSQL add foreign key constraints")
    CopyToClipboard AddForeignKeyConstraintPostgres
    
End Function

Public Function GetAllPostgreSQLCreateUniqueIndexes(frm As Object, Optional SeqModelID = "") As String

    RunCommandSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND NOT UniqueWith IS NULL ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetPostgreSQLCreateUniqueIndexes(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllPostgreSQLCreateUniqueIndexes = lines.JoinArr(vbNewLine)
        ''GetAllPostgreSQLCreateUniqueIndexes = GetGeneratedByFunctionSnippet(GetAllPostgreSQLCreateUniqueIndexes, "GetAllPostgreSQLCreateUniqueIndexes")
        CopyToClipboard GetAllPostgreSQLCreateUniqueIndexes
    End If

End Function

''Command Name: Create PostgreSQL View From Model
Public Function CreatePostgresqlViewFromModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    CreatePostgresqlViewFromModel = GetReplacedTemplate(rs, "PostgreSQL View creator")
    
    CreatePostgresqlViewFromModel = replace(CreatePostgresqlViewFromModel, "[GetPostgreSQLViewFields]", GetPostgreSQLViewFields(SeqModelID))
    CreatePostgresqlViewFromModel = replace(CreatePostgresqlViewFromModel, "[GetPostgreSQLRelatedJoins]", GetPostgreSQLRelatedJoins(SeqModelID))
    
    DoCmd.OpenForm "frmClipboardForms"
    Forms("frmClipboardForms")("Snippet") = CreatePostgresqlViewFromModel
    CopyFieldContent Forms("frmClipboardForms"), "Snippet"
    
End Function


Public Function GetPostgreSQLViewFields(SeqModelID) As String

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SlugField: SlugField = rs.fields("SlugField")
    Dim Timestamps: Timestamps = rs.fields("Timestamps")
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND Not IsGeneratedField ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    
    
    Dim letters As New clsArray: letters.arr = "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z"
    Dim uniqueFields As New clsArray
    
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetPostgreSQLViewField(SeqModelFieldID, "a", uniqueFields)
        rs.MoveNext
    Loop
    
    If Not isFalse(SlugField) Then lines.Add "a.slug"
    If Timestamps Then
        lines.Add "a.created_at"
        lines.Add "a.updated_at"
    End If
    
    sqlStr = "SELECT SeqModelFieldID, DatabaseFieldName FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " And AddEmbedding ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
        lines.Add "a." & DatabaseFieldName & "_embedding"
        rs.MoveNext
    Loop
    
    Dim i As Integer: i = 1
    
    sqlStr = "SELECT RightModelID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    
    ''Get the relateds where SeqModelID is a left
    Do Until rs.EOF
        Dim RightModelID: RightModelID = rs.fields("RightModelID"): If ExitIfTrue(isFalse(RightModelID), "RightModelID is empty..") Then Exit Function
        sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & RightModelID & _
        " AND Not IsGeneratedField AND Not PrimaryKey ORDER BY SeqModelFieldID"
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
        Do Until rs2.EOF
            SeqModelFieldID = rs2.fields("SeqModelFieldID")
            lines.Add GetPostgreSQLViewField(SeqModelFieldID, letters.arr(i), uniqueFields)
            rs2.MoveNext
        Loop
        i = i + 1
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetPostgreSQLViewFields = lines.JoinArr("," & vbNewLine)
        ''GetPostgreSQLViewFields = GetGeneratedByFunctionSnippet(GetPostgreSQLViewFields, "GetPostgreSQLViewFields")
        ''\CopyToClipboard GetPostgreSQLViewFields
    End If

End Function

Public Function GetPostgreSQLRelatedJoins(SeqModelID) As String


    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim letters As New clsArray: letters.arr = "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z"
    
    sqlStr = "SELECT SeqModelRelationshipID FROM tblSeqModelRelationships WHERE LeftModelID = " & SeqModelID & _
        " ORDER BY SeqModelRelationshipID"
    Set rs = ReturnRecordset(sqlStr)
    Dim i As Integer: i = 1
    Do Until rs.EOF
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        lines.Add GetPostgreSQLRelatedJoin(SeqModelRelationshipID, letters.arr(i))
        i = i + 1
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetPostgreSQLRelatedJoins = lines.JoinArr(vbNewLine)
        ''GetPostgreSQLRelatedJoins = GetGeneratedByFunctionSnippet(GetPostgreSQLRelatedJoins, "GetPostgreSQLRelatedJoins")
        CopyToClipboard GetPostgreSQLRelatedJoins
    End If

End Function

''Command Name: Write to getModelFormGridTemplateAreas.ts
Public Function WriteToGetmodelformgridtemplateareas_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    WriteToGetmodelformgridtemplateareas_ts = GetReplacedTemplate(rs, "getModelFormGridTemplateAreas")
    WriteToGetmodelformgridtemplateareas_ts = GetGeneratedByFunctionSnippet(WriteToGetmodelformgridtemplateareas_ts, "WriteToGetmodelformgridtemplateareas_ts", "getModelFormGridTemplateAreas")
    CopyToClipboard WriteToGetmodelformgridtemplateareas_ts
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getSaleFormGridTemplateAreas.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\get" & ModelName & "FormGridTemplateAreas.ts"
    WriteToFile filePath, WriteToGetmodelformgridtemplateareas_ts, SeqModelID, "WriteToGetmodelformgridtemplateareas_ts"
    
End Function

''Command Name: Write to getFormikControlsOnChange.ts
Public Function WriteToGetformikcontrolsonchange_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetformikcontrolsonchange_ts = GetReplacedTemplate(rs, "getFormikControlsOnChange")
    WriteToGetformikcontrolsonchange_ts = GetGeneratedByFunctionSnippet(WriteToGetformikcontrolsonchange_ts, "WriteToGetformikcontrolsonchange_ts", "getFormikControlsOnChange")
    CopyToClipboard WriteToGetformikcontrolsonchange_ts
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getFormikControlsOnChange.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getFormikControlsOnChange.ts"
    WriteToFile filePath, WriteToGetformikcontrolsonchange_ts, SeqModelID, "WriteToGetformikcontrolsonchange_ts"
    
End Function

''Command Name: Write to CustomModelFormElements.tsx
Public Function WriteToCustommodelformelements_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    WriteToCustommodelformelements_tsx = GetReplacedTemplate(rs, "CustomModelFormElements.tsx")
    WriteToCustommodelformelements_tsx = GetGeneratedByFunctionSnippet(WriteToCustommodelformelements_tsx, "WriteToCustommodelformelements_tsx", "CustomModelFormElements.tsx")
    CopyToClipboard WriteToCustommodelformelements_tsx
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\components\sales\CustomSaleFormElements.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\Custom" & ModelName & "FormElements.tsx"
    WriteToFile filePath, WriteToCustommodelformelements_tsx, SeqModelID, "WriteToCustommodelformelements_tsx"
    
End Function

''Command Name: Write to getModelFormContainerStyles.ts
Public Function WriteToGetmodelformcontainerstyles_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    WriteToGetmodelformcontainerstyles_ts = GetReplacedTemplate(rs, "getModelFormContainerStyles")
    WriteToGetmodelformcontainerstyles_ts = GetGeneratedByFunctionSnippet(WriteToGetmodelformcontainerstyles_ts, "WriteToGetmodelformcontainerstyles_ts", "getModelFormContainerStyles")
    CopyToClipboard WriteToGetmodelformcontainerstyles_ts
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getSaleFormContainerStyles.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\get" & ModelName & "FormContainerStyles.ts"
    WriteToFile filePath, WriteToGetmodelformcontainerstyles_ts, SeqModelID, "WriteToGetmodelformcontainerstyles_ts"
    
End Function

''Command Name: Write to getModelFormOtherCommonProps.ts
Public Function WriteToGetmodelformothercommonprops_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    WriteToGetmodelformothercommonprops_ts = GetReplacedTemplate(rs, "getModelFormOtherCommonProps.ts")
    WriteToGetmodelformothercommonprops_ts = GetGeneratedByFunctionSnippet(WriteToGetmodelformothercommonprops_ts, "WriteToGetmodelformothercommonprops_ts", "getModelFormOtherCommonProps.ts")
    CopyToClipboard WriteToGetmodelformothercommonprops_ts
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getSaleFormOtherCommonProps.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\get" & ModelName & "FormOtherCommonProps.ts"
    WriteToFile filePath, WriteToGetmodelformothercommonprops_ts, SeqModelID, "WriteToGetmodelformothercommonprops_ts"
End Function

''Command Name: Write to getFormikSubformGeneratorOptions.ts
Public Function WriteToGetformiksubformgeneratoroptions_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetformiksubformgeneratoroptions_ts = GetReplacedTemplate(rs, "getFormikSubformGeneratorOptions.ts")
    WriteToGetformiksubformgeneratoroptions_ts = GetGeneratedByFunctionSnippet(WriteToGetformiksubformgeneratoroptions_ts, "WriteToGetformiksubformgeneratoroptions_ts", "getFormikSubformGeneratorOptions.ts")
    CopyToClipboard WriteToGetformiksubformgeneratoroptions_ts
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getFormikSubformGeneratorOptions.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getFormikSubformGeneratorOptions.ts"
    WriteToFile filePath, WriteToGetformiksubformgeneratoroptions_ts, SeqModelID, "WriteToGetformiksubformgeneratoroptions_ts"
    
End Function

''Command Name: Get Model Formik Controls On Change
Public Function GetModelFormikControlsOnChange(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelFormikControlsOnChange = GetReplacedTemplate(rs, "GetModelFormikControlsOnChange")
    GetModelFormikControlsOnChange = GetGeneratedByFunctionSnippet(GetModelFormikControlsOnChange, "GetModelFormikControlsOnChange", "GetModelFormikControlsOnChange")
    CopyToClipboard GetModelFormikControlsOnChange
    
End Function

''Command Name: Get Model Columns To Be Overriden
Public Function GetModelColumnsToBeOverriden(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelColumnsToBeOverriden = GetReplacedTemplate(rs, "GetModelColumnsToBeOverriden")
    GetModelColumnsToBeOverriden = GetGeneratedByFunctionSnippet(GetModelColumnsToBeOverriden, "GetModelColumnsToBeOverriden", "GetModelColumnsToBeOverriden")
    CopyToClipboard GetModelColumnsToBeOverriden
    
End Function

''Command Name: Write to getModifiedInitialValues.ts
Public Function WriteToGetmodifiedinitialvalues_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetmodifiedinitialvalues_ts = GetReplacedTemplate(rs, "getModifiedInitialValues")
    WriteToGetmodifiedinitialvalues_ts = GetGeneratedByFunctionSnippet(WriteToGetmodifiedinitialvalues_ts, "WriteToGetmodifiedinitialvalues_ts", "getModifiedInitialValues")
    CopyToClipboard WriteToGetmodifiedinitialvalues_ts
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales-items\getModifiedInitialValues.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getModifiedInitialValues.ts"
    WriteToFile filePath, WriteToGetmodifiedinitialvalues_ts, SeqModelID, "WriteToGetmodifiedinitialvalues_ts"
    
End Function

''Command Name: Get Model Formik Controls On Blurs
Public Function GetModelFormikControlsOnBlurs(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelFormikControlsOnBlurs = GetReplacedTemplate(rs, "GetModelFormikControlsOnBlurs")
    GetModelFormikControlsOnBlurs = GetGeneratedByFunctionSnippet(GetModelFormikControlsOnBlurs, "GetModelFormikControlsOnBlurs", "GetModelFormikControlsOnBlurs")
    CopyToClipboard GetModelFormikControlsOnBlurs
    
End Function

''Command Name: Write to getFormikControlsOnBlur.ts
Public Function WriteToGetformikcontrolsonblur(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetformikcontrolsonblur = GetReplacedTemplate(rs, "getFormikControlsOnBlur")
    WriteToGetformikcontrolsonblur = GetGeneratedByFunctionSnippet(WriteToGetformikcontrolsonblur, "WriteToGetformikcontrolsonblur", "getFormikControlsOnBlur")
    CopyToClipboard WriteToGetformikcontrolsonblur
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales-item-sizes\getFormikControlsOnBlur.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getFormikControlsOnBlur.ts"
    WriteToFile filePath, WriteToGetformikcontrolsonblur, SeqModelID, "WriteToGetformikcontrolsonblur"
    
End Function


''Command Name: Write to getFormikControlsOnTabKeyDown
Public Function WriteToGetformikcontrolsontabkeydown(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetformikcontrolsontabkeydown = GetReplacedTemplate(rs, "getFormikControlsOnTabKeyDown")
    WriteToGetformikcontrolsontabkeydown = GetGeneratedByFunctionSnippet(WriteToGetformikcontrolsontabkeydown, "WriteToGetformikcontrolsontabkeydown", "getFormikControlsOnTabKeyDown")
    CopyToClipboard WriteToGetformikcontrolsontabkeydown
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getFormikControlsOnTabKeyDown.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getFormikControlsOnTabKeyDown.ts"
    WriteToFile filePath, WriteToGetformikcontrolsontabkeydown, SeqModelID, "WriteToGetformikcontrolsontabkeydown"
    
    
End Function

''Command Name: Write to getColumnVisibilityToOverride
Public Function WriteToGetcolumnvisibilitytooverride(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetcolumnvisibilitytooverride = GetReplacedTemplate(rs, "getColumnVisibilityToOverride")
    WriteToGetcolumnvisibilitytooverride = GetGeneratedByFunctionSnippet(WriteToGetcolumnvisibilitytooverride, "WriteToGetcolumnvisibilitytooverride", "getColumnVisibilityToOverride")
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getColumnVisibilityToOverride.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getColumnVisibilityToOverride.tsx"
    WriteToFile filePath, WriteToGetcolumnvisibilitytooverride, SeqModelID, "WriteToGetcolumnvisibilitytooverride"

End Function

''Command Name: Write to getColumnWidthToOverride
Public Function WriteToGetcolumnwidthtooverride(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetcolumnwidthtooverride = GetReplacedTemplate(rs, "getColumnWidthToOverride")
    WriteToGetcolumnwidthtooverride = GetGeneratedByFunctionSnippet(WriteToGetcolumnwidthtooverride, "WriteToGetcolumnwidthtooverride", "getColumnWidthToOverride")
    CopyToClipboard WriteToGetcolumnwidthtooverride
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getColumnWidthToOverride.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getColumnWidthToOverride.tsx"
    WriteToFile filePath, WriteToGetcolumnwidthtooverride, SeqModelID, "WriteToGetcolumnwidthtooverride"
    
End Function

''Command Name: Write to useModifiedRequiredList
Public Function WriteToUsemodifiedrequiredlist(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToUsemodifiedrequiredlist = GetReplacedTemplate(rs, "useModifiedRequiredList")
    WriteToUsemodifiedrequiredlist = GetGeneratedByFunctionSnippet(WriteToUsemodifiedrequiredlist, "WriteToUsemodifiedrequiredlist", "useModifiedRequiredList")
    CopyToClipboard WriteToUsemodifiedrequiredlist
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\useModifiedRequiredList.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\useModifiedRequiredList.ts"
    WriteToFile filePath, WriteToUsemodifiedrequiredlist, SeqModelID, "WriteToUsemodifiedrequiredlist"
End Function

''Command Name: Write to getCustomizedListElements
Public Function WriteToGetcustomizedlistelements(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetcustomizedlistelements = GetReplacedTemplate(rs, "getCustomizedListElements")
    WriteToGetcustomizedlistelements = GetGeneratedByFunctionSnippet(WriteToGetcustomizedlistelements, "WriteToGetcustomizedlistelements", "getCustomizedListElements")
    CopyToClipboard WriteToGetcustomizedlistelements
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getCustomizedListElements.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getCustomizedListElements.tsx"
    WriteToFile filePath, WriteToGetcustomizedlistelements, SeqModelID, "WriteToGetcustomizedlistelements"
    
End Function

''Command Name: Write to getModelActions
Public Function WriteToGetmodelactions(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetmodelactions = GetReplacedTemplate(rs, "getModelActions")
    WriteToGetmodelactions = GetGeneratedByFunctionSnippet(WriteToGetmodelactions, "WriteToGetmodelactions", "getModelActions")
    CopyToClipboard WriteToGetmodelactions
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getModelActions.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getModelActions.tsx"
    WriteToFile filePath, WriteToGetmodelactions, SeqModelID, "WriteToGetmodelactions"
    
End Function

''Command Name: Get Model Action
Public Function GetModelAction(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelAction = GetReplacedTemplate(rs, "GetModelAction")
    GetModelAction = GetGeneratedByFunctionSnippet(GetModelAction, "GetModelAction", "GetModelAction")
    CopyToClipboard GetModelAction
    
End Function

''Command Name: Get Use Modified Required List
Public Function GetUseModifiedRequiredList(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetUseModifiedRequiredList = GetReplacedTemplate(rs, "GetUseModifiedRequiredList")
    GetUseModifiedRequiredList = GetGeneratedByFunctionSnippet(GetUseModifiedRequiredList, "GetUseModifiedRequiredList", "GetUseModifiedRequiredList")
    CopyToClipboard GetUseModifiedRequiredList
End Function

''Command Name: Write to getRequiredListRowFilter
Public Function WriteToGetrequiredlistrowfilter(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetrequiredlistrowfilter = GetReplacedTemplate(rs, "getRequiredListRowFilter")
    WriteToGetrequiredlistrowfilter = GetGeneratedByFunctionSnippet(WriteToGetrequiredlistrowfilter, "WriteToGetrequiredlistrowfilter", "getRequiredListRowFilter")
    CopyToClipboard WriteToGetrequiredlistrowfilter
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\payment-sales\getRequiredListRowFilter.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getRequiredListRowFilter.tsx"
    WriteToFile filePath, WriteToGetrequiredlistrowfilter, SeqModelID, "WriteToGetrequiredlistrowfilter"
        
End Function

''Command Name: Get Get Required List Row Filter
Public Function GetGetRequiredListRowFilter(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetGetRequiredListRowFilter = GetReplacedTemplate(rs, "GetGetRequiredListRowFilter")
    GetGetRequiredListRowFilter = GetGeneratedByFunctionSnippet(GetGetRequiredListRowFilter, "GetGetRequiredListRowFilter", "GetGetRequiredListRowFilter")
    CopyToClipboard GetGetRequiredListRowFilter
    
End Function

''Command Name: Get Customized List Elements Option
Public Function GetCustomizedListElementsOption(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetCustomizedListElementsOption = GetReplacedTemplate(rs, "GetCustomizedListElementsOption")
    GetCustomizedListElementsOption = GetGeneratedByFunctionSnippet(GetCustomizedListElementsOption, "GetCustomizedListElementsOption", "GetCustomizedListElementsOption")
    CopyToClipboard GetCustomizedListElementsOption
    
End Function

''Command Name: Get Modified Initial Values Line
Public Function GetModifiedInitialValuesLine(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModifiedInitialValuesLine = GetReplacedTemplate(rs, "GetModifiedInitialValuesLine")
    GetModifiedInitialValuesLine = GetGeneratedByFunctionSnippet(GetModifiedInitialValuesLine, "GetModifiedInitialValuesLine", "GetModifiedInitialValuesLine")
    CopyToClipboard GetModifiedInitialValuesLine
    
End Function

''Command Name: Get getModelWideActions
Public Function Get_getModelWideActions(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Get_getModelWideActions = GetReplacedTemplate(rs, "getModelWideActions")
    Get_getModelWideActions = GetGeneratedByFunctionSnippet(Get_getModelWideActions, "Get_getModelWideActions", "getModelWideActions")
    CopyToClipboard Get_getModelWideActions
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\payments\getModelWideActions.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getModelWideActions.tsx"
    WriteToFile filePath, Get_getModelWideActions, SeqModelID, "Get_getModelWideActions"
    
End Function

''Command Name: Get Column Width To Override Option
Public Function GetColumnWidthToOverrideOption(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetColumnWidthToOverrideOption = GetReplacedTemplate(rs, "GetColumnWidthToOverrideOption")
    GetColumnWidthToOverrideOption = GetGeneratedByFunctionSnippet(GetColumnWidthToOverrideOption, "GetColumnWidthToOverrideOption", "GetColumnWidthToOverrideOption")
    CopyToClipboard GetColumnWidthToOverrideOption
    
End Function
''Command Name: Get Column Visibility To Override Option
Public Function GetColumnVisibilityToOverrideOption(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetColumnVisibilityToOverrideOption = GetReplacedTemplate(rs, "GetColumnVisibilityToOverrideOption")
    GetColumnVisibilityToOverrideOption = GetGeneratedByFunctionSnippet(GetColumnVisibilityToOverrideOption, "GetColumnVisibilityToOverrideOption", "GetColumnWidthToOverrideOption")
    CopyToClipboard GetColumnVisibilityToOverrideOption
End Function
''Command Name: Get Column Order To Override
Public Function GetColumnOrderToOverride(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetColumnOrderToOverride = GetReplacedTemplate(rs, "GetColumnOrderToOverride")
    GetColumnOrderToOverride = GetGeneratedByFunctionSnippet(GetColumnOrderToOverride, "GetColumnOrderToOverride", "GetColumnWidthToOverrideOption")
    CopyToClipboard GetColumnOrderToOverride
End Function

''Command Name: Write to getColumnOrderToOverride
Public Function WriteToGetcolumnordertooverride(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetcolumnordertooverride = GetReplacedTemplate(rs, "getColumnOrderToOverride.ts")
    WriteToGetcolumnordertooverride = GetGeneratedByFunctionSnippet(WriteToGetcolumnordertooverride, "WriteToGetcolumnordertooverride", "getColumnOrderToOverride")
    CopyToClipboard WriteToGetcolumnordertooverride
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales-items\getColumnOrderToOverride.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getColumnOrderToOverride.ts"
    WriteToFile filePath, WriteToGetcolumnordertooverride, SeqModelID, "WriteToGetcolumnordertooverride"
    
End Function

''Command Name: Write to getModelPermissions.ts
Public Function WriteToGetmodelpermissions_ts(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetmodelpermissions_ts = GetReplacedTemplate(rs, "getModelPermissions.ts")
    WriteToGetmodelpermissions_ts = GetGeneratedByFunctionSnippet(WriteToGetmodelpermissions_ts, "WriteToGetmodelpermissions_ts", "getModelPermissions.ts")
    CopyToClipboard WriteToGetmodelpermissions_ts
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\sales\getModelPermissions.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getModelPermissions.ts"
    WriteToFile filePath, WriteToGetmodelpermissions_ts, SeqModelID, "WriteToGetmodelpermissions_ts"
    
End Function

''Command Name: Get Model Permission
Public Function GetModelPermission(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelPermission = GetReplacedTemplate(rs, "GetModelPermission")
    GetModelPermission = GetGeneratedByFunctionSnippet(GetModelPermission, "GetModelPermission", "GetModelPermission")
    CopyToClipboard GetModelPermission
End Function

''Command Name: Get Add_updated_at_ Column
Public Function GetAdd_updated_at_Column(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetAdd_updated_at_Column = GetReplacedTemplate(rs, "Add updated_at column")
    'GetAdd_updated_at_Column = GetGeneratedByFunctionSnippet(GetAdd_updated_at_Column, "GetAdd_updated_at_Column", "Add updated_at column")
    CopyToClipboard GetAdd_updated_at_Column
    
End Function

Public Function GetAdd_created_at_Column(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetAdd_created_at_Column = GetReplacedTemplate(rs, "Add created_at column")
    'GetAdd_created_at_Column = GetGeneratedByFunctionSnippet(GetAdd_created_at_Column, "GetAdd_created_at_Column", "Add created_at column")
    CopyToClipboard GetAdd_created_at_Column
    
End Function

''Command Name: Get Set_updated_at_to_ N U L L
Public Function GetSet_updated_at_to_NULL(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSet_updated_at_to_NULL = GetReplacedTemplate(rs, "Set updated_at to NULL")
    ''GetSet_updated_at_to_NULL = GetGeneratedByFunctionSnippet(GetSet_updated_at_to_NULL, "GetSet_updated_at_to_NULL", "Set updated_at to NULL")
    CopyToClipboard GetSet_updated_at_to_NULL
    
End Function

''Command Name: Get Created And Updated Index
Public Function GetCreatedAndUpdatedIndex(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetCreatedAndUpdatedIndex = GetReplacedTemplate(rs, "create composite index")
    ''GetCreatedAndUpdatedIndex = GetGeneratedByFunctionSnippet(GetCreatedAndUpdatedIndex, "GetCreatedAndUpdatedIndex", "create composite index")
    CopyToClipboard GetCreatedAndUpdatedIndex
    
End Function

''Command Name: Get Delete Trigger
Public Function GetDeleteTrigger(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim PrimaryKey: PrimaryKey = ELookup("qrySeqModelFields", "PrimaryKey AND SeqModelID = " & SeqModelID, "DatabaseFieldName")
    GetDeleteTrigger = GetReplacedTemplate(rs, "Delete Trigger")
    GetDeleteTrigger = replace(GetDeleteTrigger, "[PrimaryKey]", PrimaryKey)
    ''GetDeleteTrigger = GetGeneratedByFunctionSnippet(GetDeleteTrigger, "GetDeleteTrigger", "Delete Trigger")
    CopyToClipboard GetDeleteTrigger
    
End Function

''Command Name: Get Update Timestamp to 2024-01-01
Public Function GetUpdateTimestampTo2024_01_01(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetUpdateTimestampTo2024_01_01 = GetReplacedTemplate(rs, "Update Timestamp to 2024-01-01")
    ''GetUpdateTimestampTo2024_01_01 = GetGeneratedByFunctionSnippet(GetUpdateTimestampTo2024_01_01, "GetUpdateTimestampTo2024_01_01", "Update Timestamp to 2024-01-01")
    
    CopyToClipboard GetUpdateTimestampTo2024_01_01
End Function

''Command Name: Get Model Form Grid Template Areas Option
Public Function GetModelFormGridTemplateAreasOption(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelFormGridTemplateAreasOption = GetReplacedTemplate(rs, "GetModelFormGridTemplateAreasOption")
    GetModelFormGridTemplateAreasOption = GetGeneratedByFunctionSnippet(GetModelFormGridTemplateAreasOption, "GetModelFormGridTemplateAreasOption", "GetModelFormGridTemplateAreasOption")
    CopyToClipboard GetModelFormGridTemplateAreasOption
    
End Function

''Command Name: Get Custom Model Form Elements
Public Function GetCustomModelFormElements(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetCustomModelFormElements = GetReplacedTemplate(rs, "GetCustomModelFormElements")
    GetCustomModelFormElements = GetGeneratedByFunctionSnippet(GetCustomModelFormElements, "GetCustomModelFormElements", "GetCustomModelFormElements")
    CopyToClipboard GetCustomModelFormElements
    
End Function

''Command Name: Get Formik Controls On Tab Key Downs
Public Function GetFormikControlsOnTabKeyDowns(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetFormikControlsOnTabKeyDowns = GetReplacedTemplate(rs, "GetFormikControlsOnTabKeyDowns")
    GetFormikControlsOnTabKeyDowns = GetGeneratedByFunctionSnippet(GetFormikControlsOnTabKeyDowns, "GetFormikControlsOnTabKeyDowns", "GetFormikControlsOnTabKeyDowns")
    CopyToClipboard GetFormikControlsOnTabKeyDowns
    
End Function
''Command Name: Get Model Form Container Styles Option
Public Function GetModelFormContainerStylesOption(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetModelFormContainerStylesOption = GetReplacedTemplate(rs, "GetModelFormContainerStylesOption")
    GetModelFormContainerStylesOption = GetGeneratedByFunctionSnippet(GetModelFormContainerStylesOption, "GetModelFormContainerStylesOption", "GetModelFormContainerStylesOption")
    CopyToClipboard GetModelFormContainerStylesOption
End Function
''Command Name: Get Use Model Form Template Row
Public Function GetUseModelFormTemplateRow(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ModelName: ModelName = rs.fields("ModelName")
    
    GetUseModelFormTemplateRow = GetReplacedTemplate(rs, "useModelFormTemplateRow.ts")
    GetUseModelFormTemplateRow = GetGeneratedByFunctionSnippet(GetUseModelFormTemplateRow, "GetUseModelFormTemplateRow", "useModelFormTemplateRow.ts")
    CopyToClipboard GetUseModelFormTemplateRow
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13-supabase\src\lib\cards\useModelFormTemplateRow.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\get" & ModelName & "FormTemplateRow.ts"
    WriteToFile filePath, GetUseModelFormTemplateRow, SeqModelID, "GetUseModelFormTemplateRow"
    
End Function

Public Function GetAllAddColumnPostgresStatements(frm As Object, Optional SeqModelID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SanitizedAppName: SanitizedAppName = rs.fields("SanitizedAppName"): If ExitIfTrue(isFalse(SanitizedAppName), "SanitizedAppName is empty..") Then Exit Function
    Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Function
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetAddColumnPostgresStatement(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If isPresent("qrySeqModelEmbeddings", "SeqModelID = " & SeqModelID) Then
        lines.Add "ADD COLUMN ""embedding"" vector(1536)"
    End If
    
    If lines.count > 0 Then
        GetAllAddColumnPostgresStatements = lines.JoinArr("," & vbCrLf)
        GetAllAddColumnPostgresStatements = "ALTER TABLE " & SanitizedAppName & "." & TableName & vbCrLf & GetAllAddColumnPostgresStatements & ";"
        CopyToClipboard GetAllAddColumnPostgresStatements
    End If

End Function

Public Function GetAllSeqModelSpecPrompts(frm As Object, Optional SeqModelID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SanitizedAppName: SanitizedAppName = rs.fields("SanitizedAppName"): If ExitIfTrue(isFalse(SanitizedAppName), "SanitizedAppName is empty..") Then Exit Function
    Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Function
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetSeqModelSpecPrompt(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        lines.Add "created_at,TIMESTAMP,default current timestamp"
        GetAllSeqModelSpecPrompts = "write the sql statement that creates a table in PostgreSQL under schema " & _
            Esc(SanitizedAppName) & " named " & Esc(TableName) & " with the following fields:" & vbCrLf & lines.JoinArr(vbNewLine)
        CopyToClipboard GetAllSeqModelSpecPrompts
    End If

End Function

''Command Name: Get CustomModelTableComponent
Public Function GetCustommodeltablecomponent(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetCustommodeltablecomponent = GetReplacedTemplate(rs, "CustomModelTableComponent")
    GetCustommodeltablecomponent = GetGeneratedByFunctionSnippet(GetCustommodeltablecomponent, "GetCustommodeltablecomponent", "CustomModelTableComponent")
    CopyToClipboard GetCustommodeltablecomponent
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\components\sales\CustomModelTableComponent.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\" & ModelPath & "\CustomModelTableComponent.tsx"
    WriteToFile filePath, GetCustommodeltablecomponent, SeqModelID, "GetCustommodeltablecomponent"
    
End Function

''Command Name: Get customAddToRequiredList.tsx
Public Function GetCustomaddtorequiredlist_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetCustomaddtorequiredlist_tsx = GetReplacedTemplate(rs, "customAddToRequiredList.tsx")
    GetCustomaddtorequiredlist_tsx = GetGeneratedByFunctionSnippet(GetCustomaddtorequiredlist_tsx, "GetCustomaddtorequiredlist_tsx", "customAddToRequiredList.tsx")
    CopyToClipboard GetCustomaddtorequiredlist_tsx
    
    ''C:\Users\User\Desktop\Web Development\vibram-sales\src\lib\inventoriable-hardsoles\customAddToRequiredList.tsx
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\customAddToRequiredList.tsx"
    WriteToFile filePath, GetCustomaddtorequiredlist_tsx, SeqModelID, "GetCustomaddtorequiredlist_tsx"
    
End Function

''Command Name: Validate SeqModel
Public Function ValidateSeqmodel(frm As Object, Optional SeqModelID = "") As Boolean
    
    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    If ECount("qrySeqModelSorts", "SeqModelID = " & SeqModelID) = 0 Then
        ValidateSeqmodel = False
        Dim rs2 As Recordset: Set rs2 = ReturnRecordset("SELECT * FROM qrySeqModelFields WHERE SeqModelID = " & SeqModelID & " ORDER BY SeqModelFieldID ASC")
        Dim SeqModelFieldID: SeqModelFieldID = rs2.fields("SeqModelFieldID")
        Dim VerboseFieldName: VerboseFieldName = rs2.fields("VerboseFieldName")
        RunSQL "INSERT INTO tblSeqModelSorts (SeqModelID,SeqModelFieldID,ModelFieldCaption,ModelSortOrder) VALUES (" & _
            SeqModelID & "," & SeqModelFieldID & "," & Esc(VerboseFieldName) & ",1)"
        MsgBox Esc(ModelName) & " doesn't have a valid sort."
        Exit Function
    End If
    
    ''Delete any invalid model filters
    Dim toBeDeleteds As New clsArray
    
    sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND ((SeqModelFieldModelID IS NULL AND FilterOperator <> ""isPresent"") " & _
        " OR NOT SeqModelFieldModelID = " & SeqModelID & " )"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFilterID: SeqModelFilterID = rs.fields("SeqModelFilterID")
        toBeDeleteds.Add SeqModelFilterID
        rs.MoveNext
    Loop
    
    If toBeDeleteds.count > 0 Then
        Dim item
        For Each item In toBeDeleteds.arr
            RunSQL "DELETE FROM tblSeqModelFilters WHERE SeqModelFilterID = " & item
        Next item
    End If
    
    ''Set the null WebControlTypeID to the ID of the Text control
    Dim WebControlTypeID: WebControlTypeID = ELookup("tblWebControlTypes", "WebControlType = " & Esc("Text"), "WebControlTypeID")
    
    Dim fields As New clsArray: fields.arr = "WebControlTypeID"
    Dim fieldValues As New clsArray
    Set fieldValues = New clsArray
    fieldValues.Add WebControlTypeID
    
    UpsertRecord "tblSeqModelFields", fields, fieldValues, "SeqModelID = " & SeqModelID & " AND WebControlTypeID IS NULL"
    
    ''For the filters
    Dim SeqModelFilterOperatorID: SeqModelFilterOperatorID = ELookup("tblSeqModelFilterOperators", "FilterOperator = " & _
        Esc("Equal"), "SeqModelFilterOperatorID")
    
    Set fields = New clsArray: fields.arr = "WebControlTypeID"
    Set fieldValues = New clsArray
    fieldValues.Add WebControlTypeID
    
    UpsertRecord "tblSeqModelFilters", fields, fieldValues, "SeqModelID = " & SeqModelID & " AND WebControlTypeID IS NULL"
    
    Set fields = New clsArray: fields.arr = "SeqModelFilterOperatorID"
    Set fieldValues = New clsArray
    fieldValues.Add SeqModelFilterOperatorID
    
    UpsertRecord "tblSeqModelFilters", fields, fieldValues, "SeqModelID = " & SeqModelID & " AND SeqModelFilterOperatorID IS NULL"
    
    RunSQL "UPDATE tblSeqModels SET Timestamps = -1 WHERE Timestamps = 0"
    
    If Not Validate_isPresentFilters(SeqModelID) Then Exit Function
    
    ValidateSeqmodel = True
    
End Function

Private Function Validate_isPresentFilters(SeqModelID) As Boolean
    
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelID = " & SeqModelID & " AND FilterOperator = ""isPresent"""
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Do Until rs.EOF
        Dim ModelName: ModelName = rs.fields("ModelName")
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs.fields("SeqModelRelationshipID")
        Dim FilterQueryName: FilterQueryName = rs.fields("FilterQueryName")
        Dim LeftModelID: LeftModelID = rs.fields("LeftModelID")
        Dim LeftModelName: LeftModelName = rs.fields("LeftModelName")
        Dim IsMultiple: IsMultiple = rs.fields("IsMultiple")
        Dim RightModelName: RightModelName = rs.fields("RightModelName")
        
        If isFalse(SeqModelRelationshipID) Then
            ShowError Esc(ModelName) & " has no SeqModelRelationshipID."
            Exit Function
        End If
        
        If FilterQueryName = "q" Then
            If Not isPresent("qrySeqModelFilters", "SeqModelID = " & LeftModelID & " AND FilterQueryName = " & Esc(FilterQueryName)) Then
                ShowError "There is no ""q"" filter for " & Esc(LeftModelName) & " model. The parent model is " & Esc(RightModelName)
                Exit Function
            End If
            GoTo NextRecord
        End If
        
        If IsMultiple And Not isPresent("qrySeqModelFilters", "SeqModelID = " & LeftModelID & " AND FilterQueryName = " & Esc(FilterQueryName)) Then
            ShowError Esc(FilterQueryName) & " filter is not present from " & Esc(LeftModelName) & " model."
            Exit Function
        End If
        
        If Not IsMultiple And isPresent("qrySeqModelFilters", "SeqModelID = " & LeftModelID & " AND FilterQueryName = " & Esc(FilterQueryName)) Then
            ShowError Esc(FilterQueryName) & " filter should be removed from " & Esc(LeftModelName) & " model as it is not a multiple filter. The parent model is " & Esc(RightModelName)
            Exit Function
        End If
NextRecord:
        rs.MoveNext
    Loop
    
    Validate_isPresentFilters = True
    
End Function

Public Function GetAllEmbeddingTypeDeclaration(frm As Object, Optional SeqModelID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND AddEmbedding ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetEmbeddingTypeDeclaration(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllEmbeddingTypeDeclaration = lines.JoinArr(vbNewLine)
        ''Remove the comma at the end
        GetAllEmbeddingTypeDeclaration = Left(GetAllEmbeddingTypeDeclaration, Len(GetAllEmbeddingTypeDeclaration) - 1)
        CopyToClipboard GetAllEmbeddingTypeDeclaration
    End If

End Function

Public Function GetAllAddEmbeddingDataField(frm As Object, Optional SeqModelID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " And AddEmbedding ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetAddEmbeddingDataField(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllAddEmbeddingDataField = lines.JoinArr("," & vbNewLine)
        CopyToClipboard GetAllAddEmbeddingDataField
    End If

End Function

Public Function RelatedFunctionsAndTriggers(frm As Object, Optional SeqModelID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fields: fields = "RelatedModelID,SeqModelRelationshipID,RelatedTableName,ForeignKey,TableName,SanitizedAppName"
    sqlStr = "SELECT " & fields & _
        " FROM qrySeqModelEmbeddings WHERE SeqModelID = " & SeqModelID & _
        " AND NOT RelatedModelID IS NULL GROUP BY " & fields & " ORDER BY RelatedModelID"
    Set rs = ReturnRecordset(sqlStr)
    
    Dim content
    Do Until rs.EOF
        Dim RelatedModelID: RelatedModelID = rs.fields("RelatedModelID"): If ExitIfTrue(isFalse(RelatedModelID), "RelatedModelID is empty..") Then Exit Function
        content = GetReplacedTemplate(rs, "related function and trigger for updated_embeddings")
        content = replace(content, "[FieldConditions]", GetAllEmbeddingConditions(frm, SeqModelID, RelatedModelID))
        lines.Add content
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        RelatedFunctionsAndTriggers = lines.JoinArr(vbNewLine)
        CopyToClipboard RelatedFunctionsAndTriggers
    End If

End Function

''Command Name: Get Functions and Triggers for Embedding
Public Function GetFunctionsAndTriggersForEmbedding(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetFunctionsAndTriggersForEmbedding = GetReplacedTemplate(rs, "function and trigger for updated_embeddings")
    
    Dim PrimaryKey: PrimaryKey = GetPrimaryKeyField(frm, SeqModelID)
    
    GetFunctionsAndTriggersForEmbedding = replace(GetFunctionsAndTriggersForEmbedding, "[PrimaryKey]", PrimaryKey)
    GetFunctionsAndTriggersForEmbedding = replace(GetFunctionsAndTriggersForEmbedding, "[FieldConditions]", GetAllEmbeddingConditions(frm, SeqModelID))
    GetFunctionsAndTriggersForEmbedding = replace(GetFunctionsAndTriggersForEmbedding, "[RelatedFunctionsAndTriggers]", RelatedFunctionsAndTriggers(frm, SeqModelID))
    CopyToClipboard GetFunctionsAndTriggersForEmbedding
    
End Function

Public Function GetAllEmbeddingConditions(frm As Object, Optional SeqModelID = "", Optional RelatedModelID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim whereArray As New clsArray
    whereArray.Add "SeqModelID = " & SeqModelID
    
    If Not isFalse(RelatedModelID) Then
        whereArray.Add "RelatedModelID = " & RelatedModelID
    Else
        whereArray.Add "RelatedModelID IS NULL"
    End If
    
    sqlStr = "SELECT SeqModelEmbeddingID FROM tblSeqModelEmbeddings WHERE " & whereArray.JoinArr(" AND ") & "  ORDER BY SeqModelEmbeddingID"
        
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelEmbeddingID: SeqModelEmbeddingID = rs.fields("SeqModelEmbeddingID")
        lines.Add GetEmbeddingCondition(frm, SeqModelEmbeddingID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllEmbeddingConditions = lines.JoinArr(" OR ")
        CopyToClipboard GetAllEmbeddingConditions
    End If

End Function

''Command Name: Write to model update-embeddings
''Note: This function will create a POST route file related to the SeqModelID that will update the embeddings for that specific model.
Public Function WriteToModel_update_embeddings(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SeqModelEmbeddingID: SeqModelEmbeddingID = ELookup("tblSeqModelEmbeddings", "SeqModelID = " & SeqModelID, "SeqModelEmbeddingID")
    
    If isFalse(SeqModelEmbeddingID) Then Exit Function
    
    WriteToModel_update_embeddings = GetReplacedTemplate(rs, "model update-embeddings route")
    WriteToModel_update_embeddings = GetGeneratedByFunctionSnippet(WriteToModel_update_embeddings, "WriteToModel_update_embeddings", "model update-embeddings route")
    CopyToClipboard WriteToModel_update_embeddings
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13-supabase\src\app\api\cards\update-embeddings\route.ts
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\app\api\" & ModelPath & "\update-embeddings\route.ts"
    WriteToFile filePath, WriteToModel_update_embeddings, SeqModelID, "WriteToModel_update_embeddings"
    
End Function

''Command Name: Write to Prompt component for Model
Public Function WriteToPromptComponentForModel(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToPromptComponentForModel = GetReplacedTemplate(rs, "model Prompt component")
    WriteToPromptComponentForModel = GetGeneratedByFunctionSnippet(WriteToPromptComponentForModel, "WriteToPromptComponentForModel", "model Prompt component")
    CopyToClipboard WriteToPromptComponentForModel
    
    ''C:\Users\User\Desktop\Web Development\marvel-duel-next-13-supabase\src\app\(protected)\prompt\_components\Prompt.tsx
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\app\(protected)\prompt\_components\Prompt.tsx"
    WriteToFile filePath, WriteToPromptComponentForModel, SeqModelID, "WriteToPromptComponentForModel"
    
End Function

''Command Name: Delete Unnecessary Model File
Public Function DeleteUnnecessaryModelFile(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim rs2 As Recordset
    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModelButtons"
        .AddFilter "tblFunctionChainItems.Note LIKE " & Esc("*Special File*") & " AND NOT FilePathTemplate IS NULL"
        .fields = "FilePathTemplate"
        .joins.Add GenerateJoinObj("tblFunctionChainItems", "ModelButtonID")
        .OrderBy = "FunctionOrder"
        Set rs2 = .Recordset
    End With
    
    Do Until rs2.EOF
        Dim FilePathTemplate: FilePathTemplate = rs2.fields("FilePathTemplate")
        Dim ProcessedFilePath: ProcessedFilePath = GetReplacedTemplate(rs, "", , FilePathTemplate)
        
        Dim IsProtected: IsProtected = isPresent("qryProtectedModelFiles", "SeqModelID = " & SeqModelID & " AND IsProtected " & _
            " AND filePath = " & Esc(ProcessedFilePath))
            
        Dim doesFileExists: doesFileExists = fileExists(ProcessedFilePath)
        
        If Not IsProtected And doesFileExists Then
            RunSQL "DELETE FROM tblSeqModelFiles WHERE filePath = " & Esc(ProcessedFilePath)
            Kill ProcessedFilePath
        End If
        
        rs2.MoveNext
    Loop
    
End Function

''Command Name: Write to getModelProps.tsx
Public Function WriteToGetmodelprops_tsx(frm As Object, Optional SeqModelID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    WriteToGetmodelprops_tsx = GetReplacedTemplate(rs, "getModelProps.tsx")
    WriteToGetmodelprops_tsx = GetGeneratedByFunctionSnippet(WriteToGetmodelprops_tsx, "WriteToGetmodelprops_tsx", "getModelProps.tsx")
    CopyToClipboard WriteToGetmodelprops_tsx
    
    Dim ModelPath: ModelPath = rs.fields("ModelPath"): If ExitIfTrue(isFalse(ModelPath), "ModelPath is empty..") Then Exit Function
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\lib\" & ModelPath & "\getModelProps.tsx"
    WriteToFile filePath, WriteToGetmodelprops_tsx, SeqModelID, "WriteToGetmodelprops_tsx"
    
End Function


Public Function GetAllAlterFieldDefaultPerSeqModel(frm As Form, Optional SeqModelID = "") As String

    DoCmd.RunCommand acCmdSaveRecord
     
    If isFalse(SeqModelID) Then
        SeqModelID = frm("SeqModelID")
        If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModels WHERE SeqModelID = " & SeqModelID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim BackendProjectID: BackendProjectID = rs.fields("BackendProjectID"): If ExitIfTrue(isFalse(BackendProjectID), """BackendProjectID"" is empty..") Then Exit Function
    
    Dim templateString: templateString = "ALTER TABLE [SanitizedAppName].[TableName] [GetAlterFieldDefault];"
    GetAllAlterFieldDefaultPerSeqModel = GetReplacedTemplate(rs, "None", , templateString)
    
    sqlStr = "SELECT SeqModelFieldID FROM tblSeqModelFields WHERE SeqModelID = " & SeqModelID & _
        " AND NOT DefaultValue IS NULL ORDER BY SeqModelFieldID"
    Set rs = ReturnRecordset(sqlStr)
    Do Until rs.EOF
        Dim SeqModelFieldID: SeqModelFieldID = rs.fields("SeqModelFieldID")
        lines.Add GetAlterFieldDefault(frm, SeqModelFieldID)
        rs.MoveNext
    Loop
    
    If lines.count > 0 Then
        GetAllAlterFieldDefaultPerSeqModel = replace(GetAllAlterFieldDefaultPerSeqModel, "[GetAlterFieldDefault]", lines.JoinArr(","))
        CopyToClipboard GetAllAlterFieldDefaultPerSeqModel
    End If

End Function

