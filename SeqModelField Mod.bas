Attribute VB_Name = "SeqModelField Mod"
Option Compare Database
Option Explicit

Public Function SeqModelFieldCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4, 5: ''Data Entry Form
            frm("FieldName").AfterUpdate = "=SeqModelFieldFieldNameAfterUpdate([Form])"
            frm("SeqDataTypeID").AfterUpdate = "=SeqModelFieldSeqDataTypeIDAfterUpdate([Form])"
            frm("FieldOrder").AfterUpdate = "=SeqModelFieldFieldOrderAfterUpdate([Form])"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function OpenfrmSeqDataTypes(frm As Form)

    Dim SeqDataTypeID: SeqDataTypeID = frm("SeqDataTypeID")
    If isFalse(SeqDataTypeID) Then Exit Function
    
    DoCmd.OpenForm "frmSeqDataTypes", , , "SeqDataTypeID = " & SeqDataTypeID
    
End Function

Public Function CopyModelFieldObject(frm As Form)
    
    Dim SeqModelFieldID: SeqModelFieldID = frm("SeqModelFieldID")
    
    If isFalse(SeqModelFieldID) Then Exit Function
    
    Dim ModelFieldDict As clsDictionary: Set ModelFieldDict = GetModelFieldDict(SeqModelFieldID)
    CopyToClipboard ModelFieldDict.ToFormatString(True)
    
End Function

Public Function GetModelFieldDict(SeqModelFieldID, Optional migrateMode As Boolean = False) As clsDictionary
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: tblSeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption
    
    Dim Unique, fieldName, SeqDataTypeID, Autoincrement, PrimaryKey, AllowNull, SeqModelID, DataTypeOption, DataType, DatabaseFieldName
    Unique = rs.fields("Unique")
    fieldName = rs.fields("FieldName")
    SeqDataTypeID = rs.fields("SeqDataTypeID")
    Autoincrement = rs.fields("Autoincrement")
    PrimaryKey = rs.fields("PrimaryKey")
    AllowNull = rs.fields("AllowNull")
    SeqModelID = rs.fields("SeqModelID")
    DataTypeOption = rs.fields("DataTypeOption")
    DataType = rs.fields("DataType")
    DatabaseFieldName = rs.fields("DatabaseFieldName")
    
    Dim dict As New clsDictionary
    dict.Add "type", "DataTypes." & DataType & DataTypeOption
    If Autoincrement Then dict.Add "autoIncrement", "true"
    If PrimaryKey Then dict.Add "primaryKey", "true"
    If AllowNull Then dict.Add "allowNull", "true"
    If Unique Then dict.Add "unique", "true"
    If DataType = "BOOLEAN" Then dict.Add "defaultValue", "false"
    If Not IsNull(DatabaseFieldName) Then dict.Add "field", Esc(DatabaseFieldName)
    
    ''Check for references
    If migrateMode Then
        sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND LeftForeignKey = " & Esc(fieldName)
        Set rs = ReturnRecordset(sqlStr)
        Do Until rs.EOF
            Dim RightModelID: RightModelID = rs.fields("RightModelID"): If ExitIfTrue(isFalse(RightModelID), "RightModelID is empty..") Then Exit Function
            Dim TableName: TableName = ELookup("tblSeqModels", "SeqModelID = " & RightModelID, "TableName")
            Dim references: references = "{model: {tableName:'" & TableName & "'},key:'" & DatabaseFieldName & "',onDelete: 'CASCADE',onUpdate: 'CASCADE'}"
            dict.Add "references", references
            rs.MoveNext
        Loop
    End If
        
    Dim dict2 As New clsDictionary
    dict2.Add fieldName, dict
    
    If migrateMode Then
        Set GetModelFieldDict = dict
    Else
        Set GetModelFieldDict = dict2
    End If
    
    
End Function


Public Function SeqModelFieldFieldNameAfterUpdate(frm As Form)
    
    Dim fieldName As String: fieldName = frm("FieldName")
    
    Dim VerboseFieldName: VerboseFieldName = ConvertToVerboseCaption(fieldName)
    frm("VerboseFieldName") = VerboseFieldName
    
    If isFalse(frm("DatabaseFieldName")) Then
        frm("DatabaseFieldName") = replace(LCase(VerboseFieldName), " ", "_")
    End If
    
End Function

Public Function SeqModelFieldSeqDataTypeIDAfterUpdate(frm As Form)
    
    Dim DataType As String: DataType = frm("SeqDataTypeID").Column(1)
    
    If DataType = "DECIMAL" Then
        frm("DataTypeOption") = "(10,2)"
    ElseIf DataType = "STRING" Then
        frm("DataTypeOption") = "(50)"
    End If
    
End Function

Public Function SeqModelFieldFieldOrderAfterUpdate(frm As Form)
    
    Dim FieldOrder As String: FieldOrder = frm("FieldOrder")
    
    Dim TableOrder: TableOrder = frm("TableOrder")
    
    If IsNull(TableOrder) Or TableOrder = 0 Then
            frm("TableOrder") = FieldOrder
    End If
    
End Function

Public Function CopyEnumConstantDeclaration(frm As Object, Optional SeqModelFieldID = "") As String
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
    End If
    
    If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If ExitIfTrue(rs.EOF, "There is no such record") Then Exit Function
    
    Dim DataType: DataType = rs.fields("DataType")
    Dim AllowedOptions: AllowedOptions = rs.fields("AllowedOptions")
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface")
    
    If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    If ExitIfTrue(DataType <> "ENUM" And isFalse(AllowedOptions), "Field is not an ENUM type or the field doesn't have AllowedOptions.") Then Exit Function
    
    Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption")
    If ExitIfTrue(DataType = "ENUM" And isFalse(DataTypeOption), "Data Type Option is empty..") Then Exit Function
    
    Dim options As New clsArray
    If DataType = "ENUM" Then
        Set options = ConvertEnumToArray(DataTypeOption)
    Else
        options.arr = AllowedOptions
    End If
    
    Dim PluralizedFieldName: PluralizedFieldName = rs.fields("PluralizedFieldName")
    
    If ExitIfTrue(isFalse(PluralizedFieldName), "Pluralized Field Name is empty..") Then Exit Function
    
    ''Must produce --> const battleStyles: BasicModel[] =[{id:"Attack",name:"Attack"},{id:"Guardian",name:"Guardian"},{id:"Support",name:"Support"}];
    ''Or --> const costs: number[] = [2, 3, 4, 5, 6];
    Dim objectArr As New clsArray
    Dim dict As New clsDictionary
    Dim item
    
    For Each item In options.arr
        If DataType = "ENUM" Then
            objectArr.Add "{id: " & item & ",name: " & item & "}"
        End If
    Next item
    
    If DataType = "ENUM" Then
        CopyEnumConstantDeclaration = "const " & PluralizedFieldName & ": BasicModel[]=[" & objectArr.JoinArr(",") & "];"
    Else
        If DataTypeInterface = "string" Then
            options.EscapeItems
        End If
        CopyEnumConstantDeclaration = "const " & PluralizedFieldName & ": " & DataTypeInterface & "[] = [" & options.JoinArr(",") & "];"
    End If
    
    CopyEnumConstantDeclaration = CopyEnumConstantDeclaration
    CopyEnumConstantDeclaration = GetGeneratedByFunctionSnippet(CopyEnumConstantDeclaration, "CopyEnumConstantDeclaration")
    
    CopyToClipboard CopyEnumConstantDeclaration
    
End Function

Public Function CreateFieldFormInitialValues(frm As Object, Optional SeqModelFieldID = "") As String
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
    End If
    
    If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID")
    Dim DataType: DataType = rs.fields("DataType")
    Dim AllowedOptions: AllowedOptions = rs.fields("AllowedOptions")
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption")
    Dim PluralizedFieldName: PluralizedFieldName = rs.fields("PluralizedFieldName")
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim PrimaryKey: PrimaryKey = rs.fields("PrimaryKey")
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND LeftForeignKey = " & Esc(fieldName)
    Set rs = ReturnRecordset(sqlStr)
    
    Dim options As New clsArray
    If DataType = "ENUM" Then
        If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
        Set options = ConvertEnumToArray(DataTypeOption)
        CreateFieldFormInitialValues = fieldName & ": " & IIf(AllowNull, """""", options.arr(0))
    ElseIf DataType = "DATEONLY" Then
        CreateFieldFormInitialValues = fieldName & ": " & IIf(AllowNull, """""", "convertDateToYYYYMMDD(new Date())")
    ElseIf Not IsNull(AllowedOptions) Then
        options.arr = AllowedOptions
        CreateFieldFormInitialValues = fieldName & ": " & Esc(options.arr(0))
    ElseIf fieldName Like "*month*" Then
        CreateFieldFormInitialValues = fieldName & ": " & IIf(AllowNull, """""", "getCurrentMonthNumber().toString()")
    ElseIf fieldName Like "*year*" Then
        CreateFieldFormInitialValues = fieldName & ": " & IIf(AllowNull, """""", "getCurrentYear().toString()")
    ElseIf Not rs.EOF Then
        Dim RightVariablePluralName: RightVariablePluralName = rs.fields("RightVariablePluralName"): If ExitIfTrue(isFalse(RightVariablePluralName), "RightVariablePluralName is empty..") Then Exit Function
        CreateFieldFormInitialValues = fieldName & ": " & RightVariablePluralName & " && " & RightVariablePluralName & ".length > 0 ? " & RightVariablePluralName & "[0].id.toString() : """""
    Else
        CreateFieldFormInitialValues = fieldName & ": """""
    End If
    
    CopyToClipboard CreateFieldFormInitialValues

End Function

Public Function CreateFieldSchema(frm As Object, Optional SeqModelFieldID = "", Optional IsArrayValidation As Boolean = False) As String
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
    End If
    
    If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    
    Dim VerboseFieldName: VerboseFieldName = rs.fields("VerboseFieldName"): If ExitIfTrue(isFalse(VerboseFieldName), "VerboseFieldName is empty..") Then Exit Function
    
    Dim chains As New clsArray, chainStr
    
    If Not AllowNull And Not DataTypeInterface = "boolean" Then
    
        chains.Add "required(""" & VerboseFieldName & " is a required field."")"
        
        If DataType = "ENUM" Then
            Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
            Dim Enums As New clsArray: Set Enums = ConvertEnumToArray(DataTypeOption)
            
            chains.Add "oneOf([" & Enums.JoinArr & "], """ & VerboseFieldName & " is invalid"")"
        End If
        
        chainStr = IIf(chains.count > 0, "." & chains.JoinArr("."), "")
        
        
    End If
    
    If IsArrayValidation Then
        If AllowNull Then
            CreateFieldSchema = fieldName & ": Yup." & DataTypeInterface & "().nullable()"
        Else
            CreateFieldSchema = fieldName & ": Yup." & DataTypeInterface & "().when(""touched"", ([touched], schema) => touched ? schema" & chainStr & " : schema.notRequired())"
        End If
    Else
        If AllowNull Then
            CreateFieldSchema = fieldName & ": Yup." & DataTypeInterface & "().nullable().transform((value, originalValue) =>" & _
                    "originalValue && originalValue !== """" ? value : null)" & chainStr & ","
        Else
            CreateFieldSchema = fieldName & ": Yup." & DataTypeInterface & "()" & chainStr & ","
        End If
        
    End If
    
    CopyToClipboard CreateFieldSchema
    
End Function

Private Function IsControlFirst(SeqModelFieldID)
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim FieldOrder: FieldOrder = rs.fields("FieldOrder"): If ExitIfTrue(isFalse(FieldOrder), "FieldOrder is empty..") Then Exit Function
    Dim FirstFieldOrder: FirstFieldOrder = Emin("qrySeqModelFields", "Not PrimaryKey", "FieldOrder")
    
    IsControlFirst = FieldOrder = FirstFieldOrder
    
End Function

Public Function GenerateIndividualFormControl(frm As Object, Optional SeqModelFieldID = Null)
    
    RunCommandSaveRecord
    If IsNull(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
    End If
    
    If isFalse(SeqModelFieldID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: qrySeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName|PluralizedFieldName
    ''AllowedOptions|ControlType|FieldOrder|VerboseFieldName|DataType|DataTypeInterface

    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    Dim VerboseFieldName: VerboseFieldName = rs.fields("VerboseFieldName"): If ExitIfTrue(isFalse(VerboseFieldName), "VerboseFieldName is empty..") Then Exit Function
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    Dim FieldOrder: FieldOrder = rs.fields("FieldOrder"): If ExitIfTrue(isFalse(FieldOrder), "FieldOrder is empty..") Then Exit Function
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    Dim PluralizedFieldName: PluralizedFieldName = rs.fields("PluralizedFieldName")
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim NoneStr As String
    If AllowNull Then NoneStr = "{ value: """", label: ""None"" }"
    
    ''Check if setFocusOnLoad will be rendered true, should be the first FieldOrder of this qrySeqModelFields using the SeqModelID
    Dim setFocusOnLoad: setFocusOnLoad = IIf(IsControlFirst(SeqModelFieldID), "setFocusOnLoad={true}", "")
    ''Peek relationship
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND LeftForeignKey = " & Esc(fieldName)
    Set rs = ReturnRecordset(sqlStr)
    Dim RightVariablePluralName
    Dim TemplateContent, replacedContent
    
    Select Case ControlType
    Case "Text":
        If DataType = "DATEONLY" Then
            GenerateIndividualFormControl = "<MUIDatePicker label=" & Esc(fieldName) & " name=" & Esc(fieldName) & " " & setFocusOnLoad & " />"
        Else
            Dim typeString As String
            If DataTypeInterface = "number" Then typeString = "type=""number"""
            GenerateIndividualFormControl = "<MUIText label=" & Esc(VerboseFieldName) & " name=" & Esc(fieldName) & " " & typeString & " " & setFocusOnLoad & "/>"
        End If
    
    Case "Hidden"
        GenerateIndividualFormControl = "<Field name=" & Esc(fieldName) & " type=""hidden"" />"
    
    Case "Textarea"
        GenerateIndividualFormControl = "<MUIText label=" & Esc(VerboseFieldName) & " name=" & Esc(fieldName) & " multiline minRows={9} " & setFocusOnLoad & "/>"
    
    Case "Option":
        
        If DataType = "ENUM" Then
            If ExitIfTrue(isFalse(PluralizedFieldName), "PluralizedFieldName is empty..") Then Exit Function
            TemplateContent = GetTemplateContent("Form Control Option")
            replacedContent = replace(TemplateContent, "[None]", NoneStr)
            replacedContent = replace(replacedContent, "[List]", PluralizedFieldName)
        ElseIf Not rs.EOF Then
            RightVariablePluralName = rs.fields("RightVariablePluralName")
            If ExitIfTrue(isFalse(RightVariablePluralName), "RightVariablePluralName is empty..") Then Exit Function
            TemplateContent = GetTemplateContent("Form Control Option")
            replacedContent = replace(TemplateContent, "[None]", NoneStr)
            replacedContent = replace(replacedContent, "[List]", RightVariablePluralName)
        ElseIf DataTypeInterface = "number" Then
            If ExitIfTrue(isFalse(PluralizedFieldName), "PluralizedFieldName is empty..") Then Exit Function
            TemplateContent = GetTemplateContent("Form Control Option - Numeric")
            replacedContent = replace(TemplateContent, "[None]", NoneStr)
            replacedContent = replace(replacedContent, "[List]", PluralizedFieldName)
        End If
        
        replacedContent = replace(replacedContent, "[Caption]", VerboseFieldName)
        replacedContent = replace(replacedContent, "[FieldName]", fieldName)
        GenerateIndividualFormControl = replacedContent
    
    Case "Autocomplete":
        
        If Not rs.EOF Then
            RightVariablePluralName = rs.fields("RightVariablePluralName")
            If ExitIfTrue(isFalse(RightVariablePluralName), "RightVariablePluralName is empty..") Then Exit Function
            TemplateContent = GetTemplateContent("Form Control Autocomplete")
            replacedContent = replace(TemplateContent, "[List]", RightVariablePluralName)
        Else
            If ExitIfTrue(isFalse(PluralizedFieldName), "PluralizedFieldName is empty..") Then Exit Function
            TemplateContent = GetTemplateContent("Form Control Autocomplete")
            replacedContent = replace(TemplateContent, "[List]", PluralizedFieldName)
        End If
        
        replacedContent = replace(replacedContent, "[Caption]", VerboseFieldName)
        replacedContent = replace(replacedContent, "[FieldName]", fieldName)
        GenerateIndividualFormControl = replacedContent
    
    Case "Switch":
        Dim FilterCaption: FilterCaption = rs.fields("FilterCaption"): If ExitIfTrue(isFalse(FilterCaption), "FilterCaption is empty..") Then Exit Function
        GenerateIndividualFormControl = "<MUISwitch label=" & Esc(VerboseFieldName) & " name=" & Esc(fieldName) & " />"
        
    End Select
    
    GenerateIndividualFormControl = GetGeneratedByFunctionSnippet(GenerateIndividualFormControl, "GenerateIndividualFormControl", True)
    
    CopyToClipboard GenerateIndividualFormControl
    
End Function

Public Function GenerateIndividualFormControlArray(frm As Object, Optional SeqModelFieldID = Null)
    
    RunCommandSaveRecord
    If IsNull(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
    End If
    
    If isFalse(SeqModelFieldID) Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''TABLE: qrySeqModelFields Fields: SeqModelFieldID|Unique|FieldName|SeqDataTypeID|Autoincrement|PrimaryKey
    ''AllowNull|Timestamp|CreatedBy|RecordImportID|SeqModelID|DataTypeOption|DatabaseFieldName|PluralizedFieldName
    ''AllowedOptions|ControlType|FieldOrder|VerboseFieldName|DataType|DataTypeInterface|PluralizedModelName
    ''ModelName
    
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    Dim VerboseFieldName: VerboseFieldName = rs.fields("VerboseFieldName"): If ExitIfTrue(isFalse(VerboseFieldName), "VerboseFieldName is empty..") Then Exit Function
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    Dim FieldOrder: FieldOrder = rs.fields("FieldOrder"): If ExitIfTrue(isFalse(FieldOrder), "FieldOrder is empty..") Then Exit Function
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    Dim PluralizedModelName: PluralizedModelName = rs.fields("PluralizedModelName"): If ExitIfTrue(isFalse(PluralizedModelName), "PluralizedModelName is empty..") Then Exit Function
    Dim PluralizedFieldName: PluralizedFieldName = rs.fields("PluralizedFieldName")
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim NoneStr As String
    If AllowNull Then NoneStr = "{ value: """", label: ""None"" }"
    
    ''Declare the OnKeyDown For List Form Template
    Dim OnKeyDown: OnKeyDown = GetReplacedTemplate(rs, "OnKeyDown For List Form")
    
    ''Declare the inputRef
    Dim inputRef: inputRef = GetReplacedTemplate(rs, "inputRef for List Form")
        
    ''Peek relationship
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftModelID = " & SeqModelID & " AND LeftForeignKey = " & Esc(fieldName)
    Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
    
    Dim RightVariablePluralName, LeftForeignKey
    Dim TemplateContent As String, replacedContent
    
    If ControlType = "Text" And DataType = "DECIMAL" Then
        TemplateContent = GetReplacedTemplate(rs, "MUINumber For List Form")
    ElseIf ControlType = "Text" Then
        Dim typeString As String
        If DataTypeInterface = "number" Then typeString = "type=""number"""
        TemplateContent = GetReplacedTemplate(rs, "MUIText For List Form")
        TemplateContent = replace(TemplateContent, "[type]", typeString)
    ElseIf ControlType = "Hidden" Then
        GenerateIndividualFormControlArray = "<Field name={`" & PluralizedModelName & ".${index}." & fieldName & "`} type=""hidden"" />"
    ElseIf ControlType = "Textarea" Then
        TemplateContent = GetReplacedTemplate(rs, "Textarea For List Form")
    ElseIf ControlType = "Option" And DataType = "ENUM" Then
        
        TemplateContent = GetReplacedTemplate(rs, "MUIRadio For List Form")
        TemplateContent = replace(TemplateContent, "[NoneStr]", NoneStr)
        
    ElseIf ControlType = "Option" And Not rs2.EOF Then
        
        RightVariablePluralName = rs2.fields("RightVariablePluralName"): If ExitIfTrue(isFalse(RightVariablePluralName), "RightVariablePluralName is empty..") Then Exit Function
        
        TemplateContent = GetReplacedTemplate(rs, "MUI Radio for List Form using Related")
        TemplateContent = replace(TemplateContent, "[NoneStr]", NoneStr)
        TemplateContent = replace(TemplateContent, "[List]", RightVariablePluralName)
        
    ElseIf ControlType = "Option" And DataTypeInterface = "number" Then
    
        If ExitIfTrue(isFalse(PluralizedFieldName), "PluralizedFieldName is empty..") Then Exit Function
        TemplateContent = GetTemplateContent("Form Control Option - Numeric")
        TemplateContent = replace(TemplateContent, "[None]", NoneStr)
        TemplateContent = replace(TemplateContent, "[List]", PluralizedFieldName)
        TemplateContent = replace(TemplateContent, "[Caption]", VerboseFieldName)
        TemplateContent = replace(TemplateContent, "[FieldName]", fieldName)
    
        
    ElseIf ControlType = "Autocomplete" And Not rs2.EOF Then
        
        RightVariablePluralName = rs2.fields("RightVariablePluralName"): If ExitIfTrue(isFalse(RightVariablePluralName), "RightVariablePluralName is empty..") Then Exit Function
        LeftForeignKey = rs.fields("LeftForeignKey"): If ExitIfTrue(isFalse(LeftForeignKey), "LeftForeignKey is empty..") Then Exit Function
        TemplateContent = GetTemplateContent("Form Control Array Autocomplete")
        TemplateContent = replace(TemplateContent, "[RightVariablePluralName]", RightVariablePluralName)
        TemplateContent = replace(TemplateContent, "[PluralizedModelName]", PluralizedModelName)
        TemplateContent = replace(TemplateContent, "[FieldName]", fieldName)
        
    ElseIf ControlType = "Autocomplete" Then
        
        If ExitIfTrue(isFalse(PluralizedFieldName), "PluralizedFieldName is empty..") Then Exit Function
        TemplateContent = GetTemplateContent("Form Control Array Autocomplete")
        TemplateContent = replace(TemplateContent, "[RightVariablePluralName]", PluralizedFieldName)
        TemplateContent = replace(TemplateContent, "[FieldName]", fieldName)
        TemplateContent = replace(TemplateContent, "[PluralizedModelName]", PluralizedModelName)
        
        
    ElseIf ControlType = "Switch" Then
        TemplateContent = GetReplacedTemplate(rs, "Switch For List Form")
    End If
    
    ModifyTemplateContent TemplateContent, FieldOrder, SeqModelID, inputRef, OnKeyDown, LeftForeignKey
    GenerateIndividualFormControlArray = TemplateContent
    GenerateIndividualFormControlArray = GetGeneratedByFunctionSnippet(GenerateIndividualFormControlArray, "GenerateIndividualFormControlArray", True)
    CopyToClipboard GenerateIndividualFormControlArray
    
End Function

Private Function ModifyTemplateContent(ByRef TemplateContent As String, ByVal FieldOrder As Double, ByVal SeqModelID As Long, ByVal inputRef As String, ByVal OnKeyDown As String, ByVal ForeignKey) As String
    Dim filterStr As String
    Dim minFieldOrder As Double
    Dim maxFieldOrder As Double
    
    ' Get the minimum and maximum FieldOrder values for the SeqModelID.
    filterStr = "SeqModelID = " & SeqModelID & " AND NOT PrimaryKey"
    If Not isFalse(ForeignKey) Then
        filterStr = filterStr & " AND FieldName <> " & Esc(ForeignKey)
    End If
    minFieldOrder = Emin("qrySeqModelFields", filterStr, "FieldOrder")
    maxFieldOrder = Emax("qrySeqModelFields", filterStr, "FieldOrder")
    
    ' Replace the [inputRef] and [OnKeyDown] placeholders in the TemplateContent string.
    TemplateContent = replace(TemplateContent, "[inputRef]", IIf(FieldOrder = minFieldOrder, inputRef, ""))
    TemplateContent = replace(TemplateContent, "[OnKeyDown]", IIf(FieldOrder = maxFieldOrder, OnKeyDown, ""))
End Function

''can be center | right | left
Private Function GetTableCellAlignment(SeqModelFieldID) As String
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID)
    
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    If DataType = "DECIMAL" Then
        GetTableCellAlignment = "align = ""right"""
    End If
    
End Function

Public Function GenerateFieldTableCellHeader(frm As Object, Optional SeqModelFieldID = "") As String
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
    End If
    
    If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim VerboseFieldName: VerboseFieldName = rs.fields("VerboseFieldName"): If ExitIfTrue(isFalse(VerboseFieldName), "VerboseFieldName is empty..") Then Exit Function
    
    Dim FieldWidth: FieldWidth = rs.fields("FieldWidth")
    Dim FieldWidthPart: FieldWidthPart = IIf(isFalse(FieldWidth), "", "sx={{ width: " & Esc(FieldWidth & "px") & " }}")
    
    Dim TableCellAlignment: TableCellAlignment = GetTableCellAlignment(SeqModelFieldID)
    
    GenerateFieldTableCellHeader = "<TableCell " & TableCellAlignment & " " & FieldWidthPart & ">" & VerboseFieldName & "</TableCell>"
    GenerateFieldTableCellHeader = GetGeneratedByFunctionSnippet(GenerateFieldTableCellHeader, "GenerateFieldTableCellHeader", True)
    CopyToClipboard GenerateFieldTableCellHeader
    
End Function

Public Function GenerateEnumValidationForSingleField(frm As Object, Optional SeqModelFieldID = "") As String
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
    End If
    
    If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
    
    Dim options As New clsArray: Set options = ConvertEnumToArray(DataTypeOption)
    ''options.EscapeItems
    
    GenerateEnumValidationForSingleField = GetReplacedTemplate(rs, "Enum Validation")
    GenerateEnumValidationForSingleField = replace(GenerateEnumValidationForSingleField, "[Enums]", "[" & options.JoinArr(",") & "]")
    GenerateEnumValidationForSingleField = GenerateEnumValidationForSingleField & " //Generated by GenerateEnumValidationForSingleField"
    CopyToClipboard GenerateEnumValidationForSingleField
    
End Function

Public Function GenerateCreateUpdateFieldAsChild(frm As Object, Optional SeqModelFieldID = "") As String
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
    End If
    
    If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''FieldName, DataTypeInterface
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    Dim ModelName: ModelName = rs.fields("ModelName"): If ExitIfTrue(isFalse(ModelName), "ModelName is empty..") Then Exit Function
    
    Dim fieldValue: fieldValue = "item." & fieldName
    If DataTypeInterface = "number" Then
        fieldValue = "parseInt(" & fieldValue & " as string)"
    End If
    
'    If DataType = "DECIMAL" Then
'        fieldValue = "convertStringToFloat(" & fieldValue & ")"
'    End If
    
    If DataType = "ENUM" Then
        fieldValue = fieldValue & " as " & ModelName & "Model[" & Esc(fieldName) & "]"
    End If
    
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    If AllowNull Then
        fieldValue = "item." & fieldName & " ? " & fieldValue & " : null"
    End If
    
    GenerateCreateUpdateFieldAsChild = fieldName & ": " & fieldValue
    CopyToClipboard GenerateCreateUpdateFieldAsChild
    
End Function

Public Function GenerateCreateUpdateField(frm As Object, Optional SeqModelFieldID = "") As String
    
    RunCommandSaveRecord
    
    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
    End If
    
    If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    ''FieldName, DataTypeInterface
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    
    Dim fieldValue: fieldValue = fieldName & "!"
    
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    If DataType = "ENUM" Then
        Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
        Dim Enums As New clsArray: Set Enums = ConvertEnumToArray(DataTypeOption)
        fieldValue = fieldValue & " as " & Enums.JoinArr(" | ")
    ElseIf DataType = "DATEONLY" Then
        fieldValue = "convertDateStringToYYYYMMDD(" & fieldValue & ")"
    End If
    
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    If AllowNull Then
        fieldValue = fieldName & " === """" ? null : " & fieldValue
    End If
    
    GenerateCreateUpdateField = fieldName & ": " & fieldValue
    CopyToClipboard GenerateCreateUpdateField
    
End Function


Public Function Get_updateModelsEnumValidation(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    
    Dim options As New clsArray: Set options = ConvertEnumToArray(DataTypeOption)
    Dim Enums: Enums = options.JoinArr(",")
    
    ''updateModelsEnumValidation Null Allowed
    Get_updateModelsEnumValidation = GetReplacedTemplate(rs, IIf(AllowNull, "updateModelsEnumValidation Null Allowed", "updateModelsEnumValidation"))
    Get_updateModelsEnumValidation = replace(Get_updateModelsEnumValidation, "[Enums]", Enums)
    
    Get_updateModelsEnumValidation = GetGeneratedByFunctionSnippet(Get_updateModelsEnumValidation, "Get_updateModelsEnumValidation")
    
    CopyToClipboard Get_updateModelsEnumValidation
    
End Function


Public Function GetBackendModelRequiredSnippet(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetBackendModelRequiredSnippet = GetReplacedTemplate(rs, "Model Field Required")
    GetBackendModelRequiredSnippet = GetGeneratedByFunctionSnippet(GetBackendModelRequiredSnippet, "GetBackendModelRequiredSnippet")
    CopyToClipboard GetBackendModelRequiredSnippet
    
End Function

Public Function GenerateSequelizeMigrateAddColumn(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Function
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    
    Dim FieldOptions, fieldDicts As New clsArray
    Dim dict As New clsDictionary: Set dict = GetModelFieldDict(SeqModelFieldID, True)
    Dim ModelFieldDict: ModelFieldDict = dict.ToFormatString(True)
    fieldDicts.Add ModelFieldDict
    
    ''should generate something like this
'    {
'      type: Sequelize.STRING,
'      allowNull: true,
'    }
    FieldOptions = "{" & fieldDicts.JoinArr("," & vbCrLf) & "}"
    
    GenerateSequelizeMigrateAddColumn = GetReplacedTemplate(rs, "Sequelize Migrate Add Column")
    GenerateSequelizeMigrateAddColumn = replace(GenerateSequelizeMigrateAddColumn, "[FieldOptions]", FieldOptions)
    GenerateSequelizeMigrateAddColumn = replace(GenerateSequelizeMigrateAddColumn, "DataTypes.", "Sequelize.")
    GenerateSequelizeMigrateAddColumn = GetGeneratedByFunctionSnippet(GenerateSequelizeMigrateAddColumn, "GenerateSequelizeMigrateAddColumn", "Sequelize Migrate Add Column")
    CopyToClipboard GenerateSequelizeMigrateAddColumn
    
    ''npx sequelize-cli migration:generate --name add-column-to-posts
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim fileName: fileName = ConvertToCustomTimestamp & "-add_" & DatabaseFieldName & "_column_to_" & TableName & ".js"
    Dim filePath: filePath = ProjectPath & "src\migrations\" & fileName
    WriteToFile filePath, GenerateSequelizeMigrateAddColumn, SeqModelID
    
End Function

Public Function GetModelFieldType(frm As Object, Optional SeqModelFieldID = "", Optional forceString As Boolean = False, Optional forceOptional As Boolean = False)

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    Dim PrimaryKey: PrimaryKey = rs.fields("PrimaryKey")
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    
    Dim interface
    If forceString Then
        interface = "string"
    ElseIf DataType = "DECIMAL" Then
        interface = "number"
    ElseIf DataType = "ENUM" Then
        Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
        Dim EnumOptionStr: EnumOptionStr = ConvertEnumToArray(DataTypeOption).JoinArr(" | ")
        interface = EnumOptionStr
    ElseIf DataTypeInterface = "number" Then
        interface = "number | string"
    Else
        interface = DataTypeInterface
    End If
    
    GetModelFieldType = fieldName & IIf(forceOptional, "?", "") & ": " & interface & IIf(AllowNull, " | null", "") & ";"
    GetModelFieldType = GetGeneratedByFunctionSnippet(GetModelFieldType, "GetModelFieldType", "", , True)
    CopyToClipboard GetModelFieldType
    
End Function

Public Function GetTableFieldCellInput(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    Dim FieldWidth: FieldWidth = rs.fields("FieldWidth")
    Dim isExpression: isExpression = rs.fields("isExpression")
    
    Dim enableSorting: enableSorting = IIf(ControlType = "Textarea" Or isExpression Or ControlType = "FileInput", "false", "true")
    
    Dim alignment: alignment = ""
    
    If DataTypeInterface = "boolean" Or DatabaseFieldName = "number" Then
        alignment = "center"
    ElseIf isPresent("tblSeqModelRelationships", "LeftForeignKey = " & Esc(DatabaseFieldName) & " AND " & _
        "LeftModelID = " & SeqModelID) Then
        alignment = ""
    End If
    
    If Not isFalse(alignment) Then
        alignment = "alignment: " & Esc(alignment) & ","
    End If
    
    ''Options
    Dim options As String
    If DataType = "ENUM" Then
        ''options={CONTROL_OPTIONS["type"]}
        options = "options={CONTROL_OPTIONS[" & Esc(fieldName) & "]}"
    End If
    
    ''isNumeric and isWholeNumber
    Dim isNumeric As String, isWholeNumber As String
    
    If DataType = "INTEGER" And ControlType = "Text" Then
        isNumeric = "isNumeric: true,"
        isWholeNumber = "isWholeNumber: true,"
    End If
    
    sqlStr = "SELECT * FROM qrySeqModelRelationships WHERE LeftForeignKey = " & Esc(DatabaseFieldName) & " AND " & _
        "LeftModelID = " & SeqModelID
    Dim rs2 As Recordset: Set rs2 = ReturnRecordset(sqlStr)
    
    Dim GetListName As String
    Dim GetCellValue: GetCellValue = "cell.getValue()"
    
    If Not rs2.EOF Then
    
        Dim SeqModelRelationshipID: SeqModelRelationshipID = rs2.fields("SeqModelRelationshipID")
        Dim ExcludeInRequiredList: ExcludeInRequiredList = rs2.fields("ExcludeInRequiredList")
        GetListName = IIf(ExcludeInRequiredList, "", Get_listNameFromRelationship(frm, SeqModelRelationshipID))
        
        Dim RightVariableName: RightVariableName = rs2.fields("RightVariableName"): If ExitIfTrue(isFalse(RightVariableName), "RightVariableName is empty..") Then Exit Function
        Dim RightModelName: RightModelName = rs2.fields("RightModelName"): If ExitIfTrue(isFalse(RightModelName), "RightModelName is empty..") Then Exit Function
        Dim RightModelID: RightModelID = rs2.fields("RightModelID"): If ExitIfTrue(isFalse(RightModelID), "RightModelID is empty..") Then Exit Function
        
        ''Get the most unique name to be used when this is a relationship
        Dim UniqueName: UniqueName = ELookup("tblSeqModelFields", "SeqModelID = " & RightModelID & " AND Not PrimaryKey AND Unique", "FieldName")
        If isFalse(UniqueName) Then
            UniqueName = ELookup("tblSeqModelFields", "SeqModelID = " & RightModelID & " AND PrimaryKey", "FieldName") & "?.toString()"
        End If
        GetCellValue = "cell.row.original." & RightModelName & "." & UniqueName
        
    End If
    
    If Not isFalse(GetListName) Then
        ''options={cell.table.options.meta?.options?.heroList || []}
        options = "options={cell.table.options.meta?.options?." & RightVariableName & "List || []}"
    End If
    
    If DataType = "BOOLEAN" Then
        'Boolean column cell value
        GetCellValue = GetReplacedTemplate(rs, "Boolean column cell value")
    End If
    
    If DataType = "DATE" Then
        'date time column cell value
        GetCellValue = GetReplacedTemplate(rs, "DateTime format data column")
    End If
    
    If DataType = "DATEONLY" Then
        'date column cell value
        GetCellValue = GetReplacedTemplate(rs, "Date format data column")
    End If
    
    GetTableFieldCellInput = GetReplacedTemplate(rs, "Editable Table Cell")
    GetTableFieldCellInput = replace(GetTableFieldCellInput, "[Width]", IIf(isFalse(FieldWidth), "", "width: " & FieldWidth & ","))
    GetTableFieldCellInput = replace(GetTableFieldCellInput, "[enableSorting]", enableSorting)
    GetTableFieldCellInput = replace(GetTableFieldCellInput, "[Alignment]", alignment)
    GetTableFieldCellInput = replace(GetTableFieldCellInput, "[Options]", options)
    GetTableFieldCellInput = replace(GetTableFieldCellInput, "[GetListName]", GetListName)
    GetTableFieldCellInput = replace(GetTableFieldCellInput, "[isNumeric]", isNumeric)
    GetTableFieldCellInput = replace(GetTableFieldCellInput, "[isWholeNumber]", isWholeNumber)
    GetTableFieldCellInput = replace(GetTableFieldCellInput, "[GetCellValue]", GetCellValue)

    GetTableFieldCellInput = GetGeneratedByFunctionSnippet(GetTableFieldCellInput, "GetTableFieldCellInput", "Editable Table Cell")
    CopyToClipboard GetTableFieldCellInput
    
End Function

Public Function GetFormDefaultValue(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    Dim AllowedOptions: AllowedOptions = rs.fields("AllowedOptions")
    Dim options As New clsArray
    
    Dim DefaultValue: DefaultValue = Esc("")
    
    If AllowNull Then DefaultValue = "null"
    
    If DataTypeInterface = "boolean" Then
        DefaultValue = "false"
    ElseIf DataType = "ENUM" Then
        Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption"): If ExitIfTrue(isFalse(DataTypeOption), "DataTypeOption is empty..") Then Exit Function
        Set options = ConvertEnumToArray(DataTypeOption)
        ''Get the first item
        If Not AllowNull Then
            DefaultValue = options.arr(0) & " as const"
        Else
            DefaultValue = "null"
        End If
    ElseIf DataType = "Decimal" Then
        If Not AllowNull Then
            DefaultValue = Esc("0.00")
        End If
    ElseIf DataType = "DATEONLY" Then
        If Not AllowNull Then
            DefaultValue = "convertDateToYYYYMMDD(new Date())"
        End If
    ElseIf Not isFalse(AllowedOptions) Then
        options.arr = AllowedOptions
        options.EscapeItems
        ''Get the first item
        If Not AllowNull Then
            DefaultValue = options.arr(0) & " as const"
        Else
            DefaultValue = "null"
        End If
    End If
    
    GetFormDefaultValue = fieldName & ": " & DefaultValue & ","
    GetFormDefaultValue = GetGeneratedByFunctionSnippet(GetFormDefaultValue, "GetFormDefaultValue", "", , True)
    CopyToClipboard GetFormDefaultValue
    
End Function

Public Function GenerateAQFilterField(frm As Object, Optional SeqModelFilterID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFilterID) Then
        SeqModelFilterID = frm("SeqModelFilterID")
        If ExitIfTrue(isFalse(SeqModelFilterID), "SeqModelFilterID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFilters WHERE SeqModelFilterID = " & SeqModelFilterID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GenerateAQFilterField = GetReplacedTemplate(rs, "GenerateAQFilterField")
    GenerateAQFilterField = GetGeneratedByFunctionSnippet(GenerateAQFilterField, "GenerateAQFilterField", "GenerateAQFilterField")
    CopyToClipboard GenerateAQFilterField
    
End Function

Public Function GetModelAttribute(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    
    GetModelAttribute = Esc(fieldName)
    
End Function

Public Function GetFieldsToUpdate(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), "DataTypeInterface is empty..") Then Exit Function
    Dim VariableName: VariableName = rs.fields("VariableName"): If ExitIfTrue(isFalse(VariableName), "VariableName is empty..") Then Exit Function
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    
    Dim fieldValue: fieldValue = VariableName & "." & fieldName & "!"
    
    
    GetFieldsToUpdate = fieldName & ": " & VariableName & "." & fieldName & "!"
    
    If DataType = "DECIMAL" Then
        fieldValue = VariableName & "." & fieldName & "!"
    ElseIf DataTypeInterface = "number" Then
        fieldValue = "parseInt(" & VariableName & "." & fieldName & " as string)"
    Else
        fieldValue = VariableName & "." & fieldName & "!"
    End If
    
    If AllowNull Then
        fieldValue = VariableName & "." & fieldName & " ? " & fieldValue & " : null"
    End If
    
    GetFieldsToUpdate = fieldName & ": " & fieldValue
     
    CopyToClipboard GetFieldsToUpdate
    
End Function

Public Function GetUniqueField(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetUniqueField = GetReplacedTemplate(rs, "GetUniqueField")
    CopyToClipboard GetUniqueField
    
End Function

Public Function GetRequiredField(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM tblSeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetRequiredField = GetReplacedTemplate(rs, "Get Required Field")
    GetRequiredField = GetGeneratedByFunctionSnippet(GetRequiredField, "GetRequiredField", "Get Required Field", , True)
    CopyToClipboard GetRequiredField
    
End Function

Public Function GetSimpleJoinFields(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetSimpleJoinFields = GetReplacedTemplate(rs, "Simple Join Fields")
    CopyToClipboard GetSimpleJoinFields
    
End Function

Public Function GetModelFieldName(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    
    GetModelFieldName = fieldName
    
End Function

Public Function GetFieldOptions(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption")
    Dim AllowedOptions: AllowedOptions = rs.fields("AllowedOptions")
    Dim options As New clsArray, item, choices As New clsArray
    
    If Not isFalse(DataTypeOption) Then
        Set options = ConvertEnumToArray(DataTypeOption)
    End If
    
    If Not isFalse(AllowedOptions) Then
        options.arr = AllowedOptions
        options.EscapeItems
    End If
    
    For Each item In options.arr
        choices.Add "{id: " & item & ",name: " & item & "}"
    Next item
    
    GetFieldOptions = fieldName & ": [" & choices.JoinArr & "],"
    GetFieldOptions = GetGeneratedByFunctionSnippet(GetFieldOptions, "GetFieldOptions", "", , True)
    
    
End Function

Public Function GetExpressionLiteralField(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetExpressionLiteralField = GetReplacedTemplate(rs, "ExpressionLiteralField")
    GetExpressionLiteralField = GetGeneratedByFunctionSnippet(GetExpressionLiteralField, "GetExpressionLiteralField")
    CopyToClipboard GetExpressionLiteralField
    
End Function

Public Function GetConstantFieldDictionary(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim VerboseFieldName: VerboseFieldName = rs.fields("VerboseFieldName"): If ExitIfTrue(isFalse(VerboseFieldName), "VerboseFieldName is empty..") Then Exit Function
    
    Dim GetOrderFromModelSorts: GetOrderFromModelSorts = ELookup("tblSeqModelSorts", "SeqModelFieldID = " & SeqModelFieldID, "ModelSortOrder")
    If isFalse(GetOrderFromModelSorts) Then GetOrderFromModelSorts = "0"
    
    Dim GetVerboseSortFieldName: GetVerboseSortFieldName = ELookup("tblSeqModelSorts", "SeqModelFieldID = " & SeqModelFieldID, "ModelFieldCaption")
    If GetVerboseSortFieldName = VerboseFieldName Then
        GetVerboseSortFieldName = ""
    Else
        GetVerboseSortFieldName = "verboseSortFieldName: " & Esc(GetVerboseSortFieldName) & ","
    End If
    
    GetConstantFieldDictionary = GetReplacedTemplate(rs, "Constant Field Dictionary") & ","
    GetConstantFieldDictionary = replace(GetConstantFieldDictionary, "[GetOrderFromModelSorts]", GetOrderFromModelSorts)
    GetConstantFieldDictionary = replace(GetConstantFieldDictionary, "[GetVerboseSortFieldName]", GetVerboseSortFieldName)
    GetConstantFieldDictionary = GetGeneratedByFunctionSnippet(GetConstantFieldDictionary, "GetConstantFieldDictionary", "Constant Field Dictionary", , True)
    CopyToClipboard GetConstantFieldDictionary
    
End Function

Public Function GetFilterSwitchFormControl(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetFilterSwitchFormControl = GetReplacedTemplate(rs, "Filter Switch Form Control")
    GetFilterSwitchFormControl = GetGeneratedByFunctionSnippet(GetFilterSwitchFormControl, "GetFilterSwitchFormControl", "Filter Switch Form Control", True)
    CopyToClipboard GetFilterSwitchFormControl
    
End Function

Public Function GetFacetedFormControl(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetFacetedFormControl = GetReplacedTemplate(rs, "GetFacetedFormControl")
    GetFacetedFormControl = GetGeneratedByFunctionSnippet(GetFacetedFormControl, "GetFacetedFormControl", "GetFacetedFormControl", True)
    CopyToClipboard GetFacetedFormControl
    
End Function

Public Function GetComboBoxFormControl(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    
    GetComboBoxFormControl = GetReplacedTemplate(rs, "GetComboBoxFormControl")
    GetComboBoxFormControl = replace(GetComboBoxFormControl, "[GetFormOptionOrModeList]", GetFormOptionOrModeList(frm, SeqModelFieldID))
    GetComboBoxFormControl = GetGeneratedByFunctionSnippet(GetComboBoxFormControl, "GetComboBoxFormControl", "GetComboBoxFormControl", True)
    CopyToClipboard GetComboBoxFormControl
    
End Function

Public Function GetSelectFormControl(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim GetAllowBlank: GetAllowBlank = IIf(AllowNull, "true", "false")
    
    GetSelectFormControl = GetReplacedTemplate(rs, "Select Form Control")
    GetSelectFormControl = replace(GetSelectFormControl, "[GetFormOptionOrModeList]", GetFormOptionOrModeList(frm, SeqModelFieldID))
    GetSelectFormControl = replace(GetSelectFormControl, "[GetAllowBlank]", GetAllowBlank)
    GetSelectFormControl = GetGeneratedByFunctionSnippet(GetSelectFormControl, "GetSelectFormControl", "Select Form Control", True)
    CopyToClipboard GetSelectFormControl
    
End Function

Public Function GetInputFormControl(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim SeqModelID: SeqModelID = rs.fields("SeqModelID"): If ExitIfTrue(isFalse(SeqModelID), "SeqModelID is empty..") Then Exit Function
    Dim fieldName: fieldName = rs.fields("FieldName"): If ExitIfTrue(isFalse(fieldName), "FieldName is empty..") Then Exit Function
    Dim FirstFieldName: FirstFieldName = ELookup("tblSeqModelFields", "SeqModelID = " & SeqModelID & " AND NOT PrimaryKey", "FieldName", "FieldOrder")
    Dim GetFirstControlInForm As String
    
    If FirstFieldName = fieldName Then
        GetFirstControlInForm = "ref={ref}" & vbNewLine & "setFocusOnLoad={true}"
    End If
    
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim IsNullAllowed: IsNullAllowed = IIf(AllowNull, "true", "false")
    
    GetInputFormControl = GetReplacedTemplate(rs, "Input Form Control")
    GetInputFormControl = replace(GetInputFormControl, "[GetFirstControlInForm]", GetFirstControlInForm)
    GetInputFormControl = replace(GetInputFormControl, "[IsNullAllowed]", IsNullAllowed)
    GetInputFormControl = GetGeneratedByFunctionSnippet(GetInputFormControl, "GetInputFormControl", "Input Form Control", True)
    CopyToClipboard GetInputFormControl
    
End Function

Public Function GetFormikControl(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ControlType: ControlType = rs.fields("ControlType"): If ExitIfTrue(isFalse(ControlType), "ControlType is empty..") Then Exit Function
    
    If ControlType = "Switch" Then
        GetFormikControl = GetFilterSwitchFormControl(frm, SeqModelFieldID)
    ElseIf ControlType = "FacetedControl" Then
        GetFormikControl = GetFacetedFormControl(frm, SeqModelFieldID)
    ElseIf ControlType = "Combobox" Then
        GetFormikControl = GetComboBoxFormControl(frm, SeqModelFieldID)
    ElseIf ControlType = "Select" Then
        GetFormikControl = GetSelectFormControl(frm, SeqModelFieldID)
    Else
        GetFormikControl = GetInputFormControl(frm, SeqModelFieldID)
    End If
   
EndThisFunction:
    CopyToClipboard GetFormikControl
    
End Function

Public Function GetUniqueKeysOption(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetUniqueKeysOption = GetReplacedTemplate(rs, "GetUniqueKeysOption")
    GetUniqueKeysOption = GetGeneratedByFunctionSnippet(GetUniqueKeysOption, "GetUniqueKeysOption", "GetUniqueKeysOption", , True)
    CopyToClipboard GetUniqueKeysOption
    
End Function

Public Function GetHiddenColumns(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetHiddenColumns = GetReplacedTemplate(rs, "GetHiddenColumns")
    GetHiddenColumns = GetGeneratedByFunctionSnippet(GetHiddenColumns, "GetHiddenColumns", "GetHiddenColumns", , True)
    CopyToClipboard GetHiddenColumns
    
End Function
Public Function GetPlaceholderModelFromField(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetPlaceholderModelFromField = GetReplacedTemplate(rs, "GetPlaceholderModelFromField")
    GetPlaceholderModelFromField = GetGeneratedByFunctionSnippet(GetPlaceholderModelFromField, "GetPlaceholderModelFromField", "GetPlaceholderModelFromField")
    CopyToClipboard GetPlaceholderModelFromField
End Function

Public Function GetRelatedModelListFromField(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim RelatedModelID: RelatedModelID = rs.fields("RelatedModelID"): If ExitIfTrue(isFalse(RelatedModelID), "RelatedModelID is empty..") Then Exit Function

    GetRelatedModelListFromField = GetReplacedTemplate(rs, "GetRelatedModelListFromField")
    GetRelatedModelListFromField = replace(GetRelatedModelListFromField, "[GetAllPlaceholderModelFromField]", GetAllPlaceholderModelFromField(frm, RelatedModelID))
    GetRelatedModelListFromField = GetGeneratedByFunctionSnippet(GetRelatedModelListFromField, "GetRelatedModelListFromField", "GetRelatedModelListFromField")
    CopyToClipboard GetRelatedModelListFromField
End Function

Public Function GetRequiredListImportForRelatedModelFromField(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetRequiredListImportForRelatedModelFromField = GetReplacedTemplate(rs, "GetRequiredListImportForRelatedModelFromField")
    GetRequiredListImportForRelatedModelFromField = GetGeneratedByFunctionSnippet(GetRequiredListImportForRelatedModelFromField, "GetRequiredListImportForRelatedModelFromField", "GetRequiredListImportForRelatedModelFromField")
    CopyToClipboard GetRequiredListImportForRelatedModelFromField
End Function

Public Function CreateSequelizeMigrationForIndexAddition(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    
    CreateSequelizeMigrationForIndexAddition = GetReplacedTemplate(rs, "GetAddIndexMigrationFile")
    CreateSequelizeMigrationForIndexAddition = GetGeneratedByFunctionSnippet(CreateSequelizeMigrationForIndexAddition, "CreateSequelizeMigrationForIndexAddition", "GetAddIndexMigrationFile")
    CopyToClipboard CreateSequelizeMigrationForIndexAddition
    
    Dim ProjectPath: ProjectPath = rs.fields("ProjectPath"): If ExitIfTrue(isFalse(ProjectPath), "ProjectPath is empty..") Then Exit Function
    Dim fileName: fileName = ConvertToCustomTimestamp & "-add_index_" & DatabaseFieldName & ".js"
    Dim filePath: filePath = ProjectPath & "src\migrations\" & fileName
    WriteToFile filePath, CreateSequelizeMigrationForIndexAddition
    
End Function
Public Function GetOptionalField(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetOptionalField = GetReplacedTemplate(rs, "GetOptionalField")
    GetOptionalField = GetGeneratedByFunctionSnippet(GetOptionalField, "GetOptionalField", "GetOptionalField")
    CopyToClipboard GetOptionalField
End Function

Public Function GetOptionalFieldType(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetOptionalFieldType = GetReplacedTemplate(rs, "GetOptionalFieldType")
    GetOptionalFieldType = GetGeneratedByFunctionSnippet(GetOptionalFieldType, "GetOptionalFieldType", "GetOptionalFieldType")
    CopyToClipboard GetOptionalFieldType
    
End Function

Public Function GetSeqModelFieldKeys(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    GetSeqModelFieldKeys = "{" & vbNewLine & GetKVPairs("qrySeqModelFields", rs) & vbNewLine & "},"
    
End Function

Public Function GetPostgreSQLCreateTableField(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetPostgreSQLCreateTableField = GetReplacedTemplate(rs, "GetPostgreSQLCreateTableField")
    GetPostgreSQLCreateTableField = replace(GetPostgreSQLCreateTableField, "[GetPostgreSQLFieldDeclaration]", GetPostgreSQLFieldDeclaration(rs))
    ''GetPostgreSQLCreateTableField = GetGeneratedByFunctionSnippet(GetPostgreSQLCreateTableField, "GetPostgreSQLCreateTableField", "GetPostgreSQLCreateTableField")
    CopyToClipboard GetPostgreSQLCreateTableField
End Function

Private Function GetPostgreSQLFieldDeclaration(rs As Recordset)
    
    Dim DataType: DataType = rs.fields("DataType")
    Dim PrimaryKey: PrimaryKey = rs.fields("PrimaryKey")
    Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption")
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim Unique: Unique = rs.fields("Unique")
    Dim DefaultValue: DefaultValue = rs.fields("DefaultValue")
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    
    Select Case DataType
        Case "BIGINT"
            If PrimaryKey Then
                GetPostgreSQLFieldDeclaration = "BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY"
            Else
                GetPostgreSQLFieldDeclaration = "BIGINT"
            End If
        Case "STRING"
            GetPostgreSQLFieldDeclaration = "citext"
            If Not isFalse(DataTypeOption) Then
                Dim strLength: strLength = RemoveFirstAndLastCharacter(DataTypeOption)
                GetPostgreSQLFieldDeclaration = GetPostgreSQLFieldDeclaration & " CHECK (CHAR_LENGTH(" & DatabaseFieldName & ") <= " & strLength & ")"
                ''GetPostgreSQLFieldDeclaration = GetPostgreSQLFieldDeclaration & DataTypeOption
            End If
            
            If PrimaryKey Then
                GetPostgreSQLFieldDeclaration = GetPostgreSQLFieldDeclaration & " PRIMARY KEY"
            End If
            
        Case "INTEGER"
            If PrimaryKey Then
                GetPostgreSQLFieldDeclaration = "INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY"
            Else
                GetPostgreSQLFieldDeclaration = "INTEGER"
            End If
        Case "TEXT"
            GetPostgreSQLFieldDeclaration = "TEXT"
        Case "DECIMAL"
            GetPostgreSQLFieldDeclaration = "NUMERIC"
            If Not isFalse(DataTypeOption) Then
                GetPostgreSQLFieldDeclaration = GetPostgreSQLFieldDeclaration & DataTypeOption
            End If
        Case "DATEONLY"
            GetPostgreSQLFieldDeclaration = "DATE"
            If PrimaryKey Then
                GetPostgreSQLFieldDeclaration = GetPostgreSQLFieldDeclaration & " PRIMARY KEY"
            End If
        Case "DATE"
            GetPostgreSQLFieldDeclaration = "TIMESTAMP"
        Case "BOOLEAN"
            GetPostgreSQLFieldDeclaration = "BOOLEAN default false"
        Case "JSONB"
            GetPostgreSQLFieldDeclaration = "JSONB"
        Case Else
            ' Handle any other data types or provide a default value
            GetPostgreSQLFieldDeclaration = "UNKNOWN"
    End Select
    
    If Not isFalse(DefaultValue) And Not PrimaryKey And DataType <> "BOOLEAN" Then
        If DataType = "TEXT" Or DataType = "STRING" Then
            DefaultValue = "'" & DefaultValue & "'"
        End If
        GetPostgreSQLFieldDeclaration = GetPostgreSQLFieldDeclaration & " DEFAULT " & DefaultValue
    End If

    If Not AllowNull And Not PrimaryKey Then
        GetPostgreSQLFieldDeclaration = GetPostgreSQLFieldDeclaration & " NOT NULL"
    End If
    
    If Unique Then GetPostgreSQLFieldDeclaration = GetPostgreSQLFieldDeclaration & " UNIQUE"

End Function

Public Function GetPostgreSQLCreateUniqueIndexes(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetPostgreSQLCreateUniqueIndexes = GetReplacedTemplate(rs, "Create Unique Index PostgreSQL")
    ''GetPostgreSQLCreateUniqueIndexes = GetGeneratedByFunctionSnippet(GetPostgreSQLCreateUniqueIndexes, "GetPostgreSQLCreateUniqueIndexes", "Create Unique Index PostgreSQL")
    CopyToClipboard GetPostgreSQLCreateUniqueIndexes
    
End Function

''Command Name: Get PostgreSQL View Field
Public Function GetPostgreSQLViewField(SeqModelFieldID, Alias, uniqueFields As clsArray)

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim TableName: TableName = rs.fields("TableName"): If ExitIfTrue(isFalse(TableName), "TableName is empty..") Then Exit Function
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    ''s.dr_number
    
    GetPostgreSQLViewField = Alias & "." & DatabaseFieldName
    If (uniqueFields.InArray(DatabaseFieldName)) Then
        Dim FieldAlias: FieldAlias = TableName & "_" & DatabaseFieldName
        GetPostgreSQLViewField = GetPostgreSQLViewField & " AS " & FieldAlias
        uniqueFields.Add FieldAlias
    Else
        GetPostgreSQLViewField = GetPostgreSQLViewField
        uniqueFields.Add DatabaseFieldName
    End If
    
End Function

''Command Name: Get Add Column Postgres Statement
Public Function GetAddColumnPostgresStatement(frm As Object, Optional SeqModelFieldID = "") As String

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    Dim DefaultValue: DefaultValue = rs.fields("DefaultValue")
    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption")
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    
    Select Case DataType
        Case "BOOLEAN":
            GetAddColumnPostgresStatement = "ADD COLUMN " & DatabaseFieldName & " BOOLEAN DEFAULT " & IIf(DefaultValue <> "0", "TRUE", "FALSE")
        Case "DECIMAL":
            GetAddColumnPostgresStatement = "ADD COLUMN " & DatabaseFieldName & " DECIMAL"
            If Not isFalse(DataTypeOption) Then
                GetAddColumnPostgresStatement = GetAddColumnPostgresStatement & DataTypeOption
            End If
            If Not isFalse(DefaultValue) Then
                GetAddColumnPostgresStatement = GetAddColumnPostgresStatement & " DEFAULT " & DefaultValue
            End If
        Case Else:
            ''Default is String
            GetAddColumnPostgresStatement = "ADD COLUMN " & DatabaseFieldName & " " & DataType
            If Not isFalse(DataTypeOption) Then
                GetAddColumnPostgresStatement = GetAddColumnPostgresStatement & DataTypeOption
            End If
            If Not isFalse(DefaultValue) Then
                GetAddColumnPostgresStatement = GetAddColumnPostgresStatement & " '" & DefaultValue & "'"
            End If
    End Select
    
    If Not AllowNull And DataType <> "BOOLEAN" Then
        GetAddColumnPostgresStatement = GetAddColumnPostgresStatement & " NOT NULL"
    End If
    
End Function

''Command Name: Get Seq Model Spec Prompt
Public Function GetSeqModelSpecPrompt(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    Dim DatabaseFieldName: DatabaseFieldName = rs.fields("DatabaseFieldName"): If ExitIfTrue(isFalse(DatabaseFieldName), "DatabaseFieldName is empty..") Then Exit Function
    Dim DataType: DataType = rs.fields("DataType"): If ExitIfTrue(isFalse(DataType), "DataType is empty..") Then Exit Function
    Dim DataTypeOption: DataTypeOption = rs.fields("DataTypeOption")
    Dim DefaultValue: DefaultValue = rs.fields("DefaultValue")
    Dim AllowNull: AllowNull = rs.fields("AllowNull")
    Dim Unique: Unique = rs.fields("Unique")
    Dim PrimaryKey: PrimaryKey = rs.fields("PrimaryKey")
    Dim Autoincrement: Autoincrement = rs.fields("Autoincrement")
    
    Dim stringParts As New clsArray
    
    stringParts.Add DatabaseFieldName
    stringParts.Add DataType & IIf(isFalse(DataTypeOption), "", DataTypeOption)
    
    If PrimaryKey Then
        stringParts.Add "Primary Key"
    End If
    
    If Autoincrement Then
        stringParts.Add "Autoincrementing"
    End If
    
    If DataType = "BOOLEAN" Then
        stringParts.Add "Default: " & IIf(isFalse(DefaultValue), "false", "true")
    Else
    
        If Not isFalse(DefaultValue) Then
            stringParts.Add "Default: " & DefaultValue
        End If
        
        stringParts.Add IIf(AllowNull, "", "NOT ") & "NULLABLE"
    
        If Unique Then
            stringParts.Add "SHOULD BE UNIQUE"
        End If
    End If
    
    GetSeqModelSpecPrompt = stringParts.JoinArr(",")
    
End Function

''Command Name: Get Embedding Type Declaration
Public Function GetEmbeddingTypeDeclaration(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetEmbeddingTypeDeclaration = GetReplacedTemplate(rs, "Embedding Type Declaration")
    CopyToClipboard GetEmbeddingTypeDeclaration
    
End Function

''Command Name: Get Add Embedding Data Field
Public Function GetAddEmbeddingDataField(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)

    GetAddEmbeddingDataField = GetReplacedTemplate(rs, "Add Embedding Data Field")
    CopyToClipboard GetAddEmbeddingDataField
    
End Function

''Command Name: Get Alter Field Default
Public Function GetAlterFieldDefault(frm As Object, Optional SeqModelFieldID = "")

    RunCommandSaveRecord

    If isFalse(SeqModelFieldID) Then
        SeqModelFieldID = frm("SeqModelFieldID")
        If ExitIfTrue(isFalse(SeqModelFieldID), "SeqModelFieldID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qrySeqModelFields WHERE SeqModelFieldID = " & SeqModelFieldID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    
    Dim DefaultValue: DefaultValue = rs.fields("DefaultValue"): If ExitIfTrue(isFalse(DefaultValue), """DefaultValue"" is empty..") Then Exit Function
    Dim DataTypeInterface: DataTypeInterface = rs.fields("DataTypeInterface"): If ExitIfTrue(isFalse(DataTypeInterface), """DataTypeInterface"" is empty..") Then Exit Function
    
    If DataTypeInterface = "string" Then
        DefaultValue = "$$" & DefaultValue & "$$"
    End If
    
    GetAlterFieldDefault = GetReplacedTemplate(rs, "None", , "ALTER COLUMN [DatabaseFieldName] SET DEFAULT [DefaultValue]")
    CopyToClipboard GetAlterFieldDefault
    
    
End Function
