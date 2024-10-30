Attribute VB_Name = "ModelField Mod"
Option Compare Database
Option Explicit

Public Function ModelFieldCreate(frm As Object, FormTypeID)
    
    If FormTypeID = 5 Then ModelFieldDSCreate frm
    Select Case FormTypeID
        Case 4, 5:
            AttachFunctions frm
            'frm("ModelID").RowSource = "SELECT ModelID, Model FROM tblModels WHERE UserQueryFields = 0"
    End Select
    
End Function

Private Sub ModelFieldDSCreate(frm As Form)
    
    frm.OrderBy = "FieldOrder ASC, ModelFieldID ASC"
    frm.SubPageOrder.DefaultValue = ""
    
End Sub

Private Sub AttachFunctions(frm As Form)

    frm.ParentModelID.AfterUpdate = "=ModelFieldParentModelIDChange([Form])"

End Sub

Private Sub UpdateFields(frm As Object, ctl As control)
    
    Dim ParentModelID: ParentModelID = frm("ParentModelID")
    Dim Model: Model = frm("ParentModelID").Column(1)
    
    Dim ModelField: ModelField = Model & "ID"
    
    Dim PrimaryKey: PrimaryKey = ELookup("tblModels", "ModelID = " & ParentModelID, "PrimaryKey")
    
    Dim FieldTypeID: FieldTypeID = dbLong
    
    If Not isFalse(PrimaryKey) Then
        ModelField = PrimaryKey
        FieldTypeID = ELookup("tblModelFields", "ModelField = " & Esc(ModelField) & " AND ModelID = " & ParentModelID, "FieldTypeID")
    End If
    
    frm("VerboseName") = AddSpaces(ModelField)
    frm("ForeignKey") = ModelField
    frm("IsIndexed") = -1
    frm("FieldTypeID") = FieldTypeID
    frm("ModelField") = ModelField
    
End Sub

Private Function IsParentModelIDNull(frm As Form) As Boolean
    Dim ParentModelID As Variant
    ParentModelID = frm("ParentModelID")
    IsParentModelIDNull = IsNull(ParentModelID)
End Function


Public Function ModelFieldParentModelIDChange(frm As Form)
    
    Dim ctl As control
    Set ctl = frm("ParentModelID")
    
    If Not IsParentModelIDNull(frm) Then
        UpdateFields frm, ctl
    End If

End Function
