Attribute VB_Name = "FilterField Mod"
Option Compare Database
Option Explicit

Public Function FilterFieldCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function FilterFieldOnLoad(frm As Form)

    SetDefaultUserID frm
    
    Dim sqlStr
    sqlStr = "SELECT ModelFieldID, ModelField FROM tblModelFields"
    ''Check if there is a parent
    On Error Resume Next
    If DoesObjectExists(frm.parent) Then
        If frm.parent.Name = "frmModels" Then
            sqlStr = sqlStr & " WHERE ModelID = " & frm.parent.ModelID
        End If
    End If
    
    sqlStr = sqlStr & " ORDER BY ModelField"
    
    frm.ModelFieldID.RowSource = sqlStr
    frm.ModelFieldID.Requery
    
End Function

Public Function FilterContOptionOnChange(frm As Object, tblName, Optional FieldToUse)
    
    If DoesPropertyExists(frm, "Parent") Then
        frm.parent.Form.Dirty = True
    End If
    
    Dim Selected: Selected = frm("Selected")
    Dim id: id = frm("ID")
    
    SetFormSingleValue frm, FieldToUse
    ''Assumed to be option button
    ''Unselect everyrecord except the current ID
    RunSQL "UPDATE " & tblName & " SET Selected = 0 WHERE ID <> " & id
    frm.Requery
    
End Function

Public Function FilterContFormOnLoad(frm As Object, sqlStr As String, tblName As String)

    ''Delete the existing record using the tblName
    RunSQL "DELETE FROM " & tblName
    ''Re-insert based on the sqlStr important is the value, label from the sqlStr
    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = tblName
          .fields = "[Value], FilterLabel"
          .insertSQL = sqlStr
          .InsertUseAsPlain = True
          rowsAffected = .Run
    End With
    
    frm.Requery
   
End Function

Public Function SetSubformCaption(frm As Object, tblName, subformLabelName)
        
    ''Skip this part if the parent form is a mainform -> to be used only on DE forms
    If frm.Name Like "main*" Then Exit Function
    
    Dim FilterLabel, filterLabels As New clsArray
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT FilterLabel FROM " & tblName & " WHERE Selected")
    Do Until rs.EOF
        filterLabels.Add rs.fields(0)
        rs.MoveNext
    Loop
    
    FilterLabel = filterLabels.JoinArr(" | ")
    If isFalse(FilterLabel) Then
        FilterLabel = "None"
    End If
On Error Resume Next
    frm(subformLabelName).Caption = FilterLabel
    
End Function


Public Function ExtractFilterContFormOnLoadParams(str) As clsArray
    
    Dim matches As New clsArray
    Dim regex As New RegExp
    regex.pattern = "=FilterContFormOnLoad\(\[.*?\],\""(.*)\"",\""(.*)\""\)"
    regex.Global = True
    
    Dim match As Object
    For Each match In regex.Execute(str)
        matches.Add match.SubMatches(0) ' "SELECT Name As [Value], Name AS Label From tblTemplateControls GROUP BY Name ORDER BY Name"
        matches.Add match.SubMatches(1) ' "tblfltrName"
    Next
    
    Set ExtractFilterContFormOnLoadParams = matches
    
End Function

Public Function ExtractFilterValue(str) As String
    Dim regex As New RegExp
    regex.pattern = "fltrName_(.*)"
    regex.Global = True
    
    Dim match As Object
    For Each match In regex.Execute(str)
        ExtractFilterValue = match.SubMatches(0)  ' ButtonControl
    Next
End Function

Public Function ToggleFilterCB(frm As Object, tblName, Optional FieldToUse)

    Dim Selected: Selected = frm("Selected")
    
    frm("Selected") = Not Selected
    
    If frm("Selected").ControlType = acOptionButton Then
        FilterContOptionOnChange frm, tblName, FieldToUse
        Exit Function
    End If
    
    ''Get subformLabelName
On Error Resume Next
    Dim subformLabelName, ctl As control: subformLabelName = ""
    Dim parentFrm As Form: Set parentFrm = frm.parent.Form
    For Each ctl In parentFrm.controls
        If ctl.ControlType = acSubform Then
            Dim chldFrm As Form: Set chldFrm = ctl.Form
            Dim recordSource: recordSource = chldFrm.recordSource
            If tblName = recordSource Then
                subformLabelName = "lbl" & ctl.Name
                Exit For
            End If
        End If
    Next ctl
    
    RunCommandSaveRecord
    SetSubformCaption frm.parent.Form, tblName, subformLabelName
    
End Function


