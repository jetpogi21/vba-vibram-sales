Attribute VB_Name = "ControlCreationOptionGroup Mod"
Option Compare Database
Option Explicit
Private rs As Recordset


Public Function ControlCreationOptionGroupCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SetOptionGroupValue(frm As Object, FieldToUse, ogName)
    
    Dim controlValue: controlValue = frm(FieldToUse)
    
    If Not IsNull(controlValue) Then
        controlValue = CStr(controlValue)
        ''Get the equivalent OptionValue from the option button
        Dim ctl As control
        For Each ctl In frm.controls
            If ctl.Name Like ogName & "_*" Then
                
                If ctl.Tag = controlValue Then
                    controlValue = ctl.optionValue
                    frm(ogName) = controlValue
                    Exit For
                End If
            End If
        Next ctl
        
    Else
        frm(ogName) = -2
    End If
    
End Function

Public Function OGControlAfterUpdate(frm As Object, FieldToUse, ogName)
    
    Dim ogValue: ogValue = frm(ogName)
    
    If ogValue = -2 Then
        frm(FieldToUse) = Null
        Exit Function
    End If
    
    Dim ctl As control, controlVal
    For Each ctl In frm.controls
        If ctl.Name Like ogName & "_*" Then
            If ctl.optionValue = ogValue Then
                controlVal = ctl.Tag
                frm(FieldToUse) = controlVal
            End If
        End If
    Next ctl
            
End Function

''There will be 2 possible directions one is horizontal and the other is vertical
Public Function CreateOptionGroup(ControlCreationHelperID)
    
    ''TABLE: tblControlCreationHelper Fields: ControlCreationHelperID|CustomControlTypeID|Model|PrimaryKey
    ''MainField|FieldToUse|Direction|Timestamp|CreatedBy|RecordImportID|Width|ParentModel|IsDynamicList|FieldCaption
    ''MiddleTable|PossibleValues
    Set rs = ReturnRecordset("SELECT * FROM tblControlCreationHelper Where ControlCreationHelperID = " & ControlCreationHelperID)
    
    Dim frm As Form: Set frm = CreateForm
    Dim FieldToUse: FieldToUse = rs.fields("FieldToUse")
    Dim ControlNames As New clsArray
    Dim x: x = 100
    Dim firstObName, lblName
    
    Dim ogName: ogName = "og" & FieldToUse
    RenderOGLabel frm, ogName, x, lblName
    RenderOg frm, ogName, x
    
    Dim possibleValues: possibleValues = rs.fields("PossibleValues")
    If Not IsNull(possibleValues) Then
        RenderOGBasedOnPossibleValues frm, ogName, ControlNames, firstObName
        RepositionOptionGroup frm, ControlNames, x, lblName
        RepositionOptionOptionGroupParent frm, ogName, firstObName
    Else
        
    End If
    
End Function

Private Sub RepositionOptionOptionGroupParent(frm As Object, ogName, firstObName)
    
    Dim Top, Left
    
    ''Top = frm(firstObName).Top
    Top = GetControlPosition(frm, firstObName, "Top")
    Left = GetControlPosition(frm, firstObName, "Left")
    
    frm(ogName).Top = Top
    frm(ogName).Left = Left
    frm(ogName).Height = 0
    frm(ogName).Width = 0
    
End Sub

Private Sub RepositionOptionGroup(frm As Object, ControlNames As clsArray, x, lblName)
    
    Dim Direction: Direction = rs.fields("Direction")
    Dim Width: Width = rs.fields("Width")
    Dim PairSize: PairSize = rs.fields("PairSize")
    Dim AsInline: AsInline = rs.fields("AsInline")
    Dim proportions As New clsArray
    Dim ControlName
    
    Dim PairSizes As New clsArray: PairSizes.arr = "1,1"
    If Not isFalse(PairSize) Then PairSizes.arr = PairSize
   
    Dim i As Integer: i = 1
    If Direction = "Horizontal" Then
        
        For Each ControlName In ControlNames.arr
            If i = 2 Then
                proportions.Add PairSizes.items(1)
                i = 1
            Else
                proportions.Add PairSizes.items(0)
                i = i + 1
            End If
        Next ControlName
        
        If AsInline Then
            proportions.arr = PairSizes.items(2) & "," & proportions.JoinArr(",")
            ControlNames.arr = lblName & "," & ControlNames.JoinArr(",")
        End If
        
        RepositionControlsInRow frm, proportions.JoinArr(","), ControlNames.JoinArr(","), x, Width
    Else
        Dim tempControlNames As New clsArray
        For Each ControlName In ControlNames.arr
            If i = 2 Then
                proportions.Add PairSizes.items(1)
                tempControlNames.Add ControlName
                RepositionControlsInRow frm, proportions.JoinArr(","), tempControlNames.JoinArr(","), x, Width, , , True
                i = 1
            Else
                proportions.clearArr
                tempControlNames.clearArr
                proportions.Add PairSizes.items(0)
                tempControlNames.Add ControlName
                i = i + 1
            End If
        Next ControlName
    End If
    
End Sub

Private Sub RenderOGBasedOnPossibleValues(frm As Object, ogName, ControlNames As clsArray, firstObName)
    
    ''ogName is the option group name
    ''TABLE: tblControlCreationHelper Fields: ControlCreationHelperID|CustomControlTypeID|Model|PrimaryKey
    ''MainField|FieldToUse|Direction|Timestamp|CreatedBy|RecordImportID|Width|ParentModel|IsDynamicList|FieldCaption
    ''MiddleTable|PossibleValues
    Dim possibleValues: possibleValues = rs.fields("PossibleValues")
    Dim ParentModel: ParentModel = rs.fields("ParentModel")
    Dim NoneValue: NoneValue = rs.fields("NoneValue")
    
    Dim opt, options As New clsArray: options.arr = possibleValues
    Dim i As Integer: i = 1
    
    
    If Not isFalse(NoneValue) Then
        RenderIndividualOptionButton frm, NoneValue, ogName, -2, ControlNames, firstObName
    End If
    
    For Each opt In options.arr
    
        RenderIndividualOptionButton frm, opt, ogName, i, ControlNames, firstObName
        
    Next opt
    
End Sub

Private Sub RenderIndividualOptionButton(frm As Object, opt, ogName, optionValue, ControlNames As clsArray, firstObName)
    
    Dim obName, lblName, ctl As control
    
    obName = ogName & "_" & opt
    lblName = "lbl" & obName

    ''Render the checkbox
    Set ctl = CreateControl(frm.Name, acOptionButton, , ogName, , 0, 0, 0)
    ctl.Name = obName
    ctl.optionValue = optionValue
    optionValue = optionValue + 1
    ctl.Tag = opt
    
    FilterControlSetCommonProperties ctl, frm
    
    ''Render the Label
    Set ctl = CreateControl(frm.Name, acLabel, , obName, , 0, 0, 0)
    ctl.Name = lblName
    ctl.Caption = opt
    CopyProperties frm, lblName, "LabelControl", False
    
    ControlNames.Add obName
    ControlNames.Add lblName
    
    If isFalse(firstObName) Then
        firstObName = obName
    End If
        
End Sub

Private Sub RenderOg(frm As Object, ogName, x)
    
    Dim Width: Width = rs.fields("Width")
    Dim FieldToUse: FieldToUse = rs.fields("FieldToUse")
    ''Create the option group control
    Dim ctl As control: Set ctl = CreateControl(frm.Name, acOptionGroup, , , , 0, 0, 0, 0)
    ctl.Name = ogName
    ''RepositionControlsInRow frm, "1", ogName, x, Width
    ctl.BorderStyle = 0
    ctl.SpecialEffect = 0
    ctl.HorizontalAnchor = acHorizontalAnchorRight
    ctl.AfterUpdate = "=OGControlAfterUpdate([Form]," & Esc(FieldToUse) & "," & Esc(ogName) & ")"
    ''frm As Form, FieldToUse, ogName
End Sub

Private Sub RenderOGLabel(frm As Object, ogName, x, lblName)
    
    Dim Width: Width = rs.fields("Width")
    ''Get the caption
    Dim Caption
    Dim Model: Model = rs.fields("Model")
    Dim FieldToUse: FieldToUse = rs.fields("FieldToUse")
    Dim FieldCaption: FieldCaption = rs.fields("FieldCaption")
    Dim AsInline: AsInline = rs.fields("AsInline")
    
    If isFalse(Model) Then
        Caption = FieldCaption
        If isFalse(Caption) Then
            Caption = FieldToUse
        End If
    Else
        Dim tblName: tblName = GetTableName(Model)
        Caption = GetCaptionPropertyFromTable(tblName, FieldToUse)
    End If
    
    ''Filter caption
    Dim ctl As control: Set ctl = CreateControl(frm.Name, acLabel, , , , 0, 0, 0)
    lblName = "lbl" & ogName
    ctl.Name = lblName
    ctl.Caption = Caption
    CopyProperties frm, ctl.Name, "LabelControl", False
    If Not AsInline Then
        RepositionControlsInRow frm, "1", lblName, x, Width, , , True
    Else
        ctl.Caption = ctl.Caption & " :"
    End If
    
End Sub






