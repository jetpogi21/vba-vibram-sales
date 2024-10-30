Attribute VB_Name = "Object Navigator"
Option Compare Database
Option Explicit

Public Function OpenModule(moduleName, ProcedureName)

On Error GoTo Err_Handler:
    
    If isFalse(moduleName) And isFalse(ProcedureName) Then Exit Function
    
    If Not isFalse(moduleName) Then
        DoCmd.OpenModule moduleName, ProcedureName
    Else
        DoCmd.OpenModule , ProcedureName
    End If
    Exit Function
    
Err_Handler:
    
    If moduleName Like "Form_*" Then
        DoCmd.OpenForm replace(moduleName, "Form_", ""), acDesign
        Exit Function
    End If
    
    If moduleName Like "Report_*" Then
        DoCmd.OpenForm replace(moduleName, "Report_", ""), acDesign
        Exit Function
    End If
    
    ShowError Esc(ProcedureName) & " does not exist."
    
    ''DoCmd.OpenModule ModuleName, ProcedureName
    
End Function

Public Function ListProcedures(frm As Form)

    Dim VBProj As VBIDE.VBProject, vbComp As VBIDE.VBComponent, CodeMod As VBIDE.CodeModule
    Dim lineNum As Long, NumLines As Long, ProcName As String, ProcKind As VBIDE.vbext_ProcKind
    Dim ProcKindStr As String
    
    Set VBProj = Application.VBE.ActiveVBProject
    
    Dim fields(2) As String, fieldValues(2) As String
    fields(0) = "CustomVBAFunction"
    fields(1) = "ModuleName"
    fields(2) = "DeclarationType"
    
    For Each vbComp In VBProj.VBComponents
        fieldValues(1) = "'" & vbComp.Name & "'"
        Set vbComp = VBProj.VBComponents(vbComp.Name)
        Set CodeMod = vbComp.CodeModule
        With CodeMod
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .CountOfLines
                ProcName = .ProcOfLine(lineNum, ProcKind)
                ProcKindStr = ProcKindString(ProcKind)
                fieldValues(0) = "'" & ProcName & "'"
                fieldValues(2) = "'" & ProcKindStr & "'"
                If Not isPresent("tblCustomVBAFunctions", "CustomVBAFunction = " & fieldValues(0) & " And " & _
                    "ModuleName = " & fieldValues(1)) Then
                    InsertAndLog "tblCustomVBAFunctions", fields, fieldValues
                    
                End If
                lineNum = .ProcStartLine(ProcName, ProcKind) + .ProcCountLines(ProcName, ProcKind) + 1
            Loop
        End With
        
    Next vbComp
    
    frm.Requery
    
End Function

Private Function ProcKindString(ProcKind As VBIDE.vbext_ProcKind) As String
    Select Case ProcKind
        Case vbext_pk_Get
            ProcKindString = "Property Get"
        Case vbext_pk_Let
            ProcKindString = "Property Let"
        Case vbext_pk_Set
            ProcKindString = "Property Set"
        Case vbext_pk_Proc
            ProcKindString = "Sub Or Function"
        Case Else
            ProcKindString = "Unknown Type: " & CStr(ProcKind)
    End Select
End Function

Public Function ImportAllObjects(frm As Form)

    DeleteObjectFromRecord
    
    Dim objTypes(4) As Integer, objType As Variant
    objTypes(0) = 0
    objTypes(1) = 1
    objTypes(2) = 2
    objTypes(3) = 3
    objTypes(4) = 5
    
    For Each objType In objTypes
        SearchAccessObjects objType
    Next objType
    
    frm.subform.Requery
    
End Function

Public Function DeleteObjectFromRecord()

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblAccObjs")
     
    Do Until rs.EOF
        If Not DoesObjectExist(rs.fields("AccObj"), rs.fields("AccObjTypeID")) Then
            rs.Delete
        End If
        rs.MoveNext
    Loop
    
End Function

Private Function DoesObjectExist(AccObj As String, AccObjTypeID As Integer) As Boolean
    Select Case AccObjTypeID
        Case 0:
            DoesObjectExist = DoesPropertyExists(Application.CurrentData.AllTables, AccObj)
        Case 1:
            DoesObjectExist = DoesPropertyExists(Application.CurrentData.AllQueries, AccObj)
        Case 2:
            DoesObjectExist = DoesPropertyExists(Application.CurrentProject.AllForms, AccObj)
        Case 3:
            DoesObjectExist = DoesPropertyExists(Application.CurrentProject.AllReports, AccObj)
        Case 5:
            DoesObjectExist = DoesPropertyExists(Application.CurrentProject.AllModules, AccObj)
     End Select
End Function

Public Function SearchAccessObjects(objType As Variant)
    
    Dim obj As Object, objs As Object, db As Object
    
    Select Case objType
         Case 0:
             Set objs = Application.CurrentData.AllTables
         Case 1:
             Set objs = Application.CurrentData.AllQueries
         Case 2:
             Set objs = Application.CurrentProject.AllForms
         Case 3:
             Set objs = Application.CurrentProject.AllReports
         Case 5:
             Set objs = Application.CurrentProject.AllModules
     End Select
     
     Dim fields(1) As String, fieldValues(1) As String
     
     fields(0) = "AccObj"
     fields(1) = "AccObjTypeID"
     
     fieldValues(1) = objType
     
     For Each obj In objs
        If Not obj.Name Like "*MSys*" And Not obj.Name Like "*AccObj*" Then
            If Not isPresent("tblAccObjs", "AccObj = '" & obj.Name & "'") Then
                fieldValues(0) = "'" & obj.Name & "'"
                InsertAndLog "tblAccObjs", fields, fieldValues
            End If
        End If
     Next obj
     
     
End Function

Public Function OpenDesignView(AccObj As Variant, AccObjTypeID As Variant)
    
    Select Case AccObjTypeID
        Case 0:
            DoCmd.OpenTable AccObj, acViewDesign
        Case 1:
            DoCmd.OpenQuery AccObj, acViewDesign
        Case 2:
            DoCmd.OpenForm AccObj, acViewDesign
        Case 3:
            DoCmd.OpenReport AccObj, acViewDesign
        Case 5:
            DoCmd.OpenModule AccObj
        Case Else:
            ShowError "Please select a record.."
            Exit Function
     End Select
     
End Function

Public Function OpenNormalView(AccObj As Variant, AccObjTypeID As Variant)
    
    Select Case AccObjTypeID
        Case 0:
            DoCmd.OpenTable AccObj, acViewNormal
        Case 1:
            DoCmd.OpenQuery AccObj, acViewNormal
        Case 2:
            DoCmd.OpenForm AccObj, acViewNormal
        Case 3:
            DoCmd.OpenReport AccObj, acViewNormal
        Case 5:
            DoCmd.OpenModule AccObj
        Case Else:
            ShowError "Please select a record.."
            Exit Function
     End Select
     
End Function

