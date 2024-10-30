Attribute VB_Name = "Page Mod"
Option Compare Database
Option Explicit

Public Function PageCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function PagePageNameAfterUpdate(frm As Form)

    Dim PageName: PageName = frm("PageName")
    Dim ComponentName
    
    If Not isFalse(frm("ComponentName")) Then
        Exit Function
    End If
    
    If PageName Like "*-*" Then
        ComponentName = replace(StrConv(replace(PageName, "-", " "), vbProperCase), " ", "")
    Else
        ComponentName = StrConv(ConvertToVerboseCaption(PageName), vbProperCase)
    End If
    
    frm("ComponentName") = replace(ComponentName, " ", "")
    
End Function
Public Function GeneratePageFile(frm As Object, Optional PageID = "")

    RunCommandSaveRecord

    If isFalse(PageID) Then
        PageID = frm("PageID")
        If ExitIfTrue(isFalse(PageID), "PageID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryPages WHERE PageID = " & PageID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim PageName: PageName = rs.fields("PageName")
    Dim ComponentName: ComponentName = rs.fields("ComponentName"): If ExitIfTrue(isFalse(ComponentName), "ComponentName is empty..") Then Exit Function
    Dim IsComponent: IsComponent = rs.fields("IsComponent")
    
    GeneratePageFile = GetReplacedTemplate(rs, "basic page file")
    GeneratePageFile = GetGeneratedByFunctionSnippet(GeneratePageFile, "GeneratePageFile")
    CopyToClipboard GeneratePageFile
    
    Dim filePath
    If IsComponent Then
        filePath = ClientPath & "src\components\"
        If Not isFalse(PageName) Then
            filePath = filePath & PageName & "\"
        End If
    Else
        filePath = ClientPath & "src\app\"
        If Not isFalse(PageName) Then
            filePath = filePath & PageName & "\"
        End If
    End If
    
    filePath = filePath & IIf(IsComponent, ComponentName, "page") & ".tsx"
    
    WriteToFile filePath, GeneratePageFile
    
End Function
