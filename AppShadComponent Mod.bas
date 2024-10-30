Attribute VB_Name = "AppShadComponent Mod"
Option Compare Database
Option Explicit

Public Function AppShadComponentCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function InstallShadComponent(frm As Object, Optional AppShadComponentID = "")

    RunCommandSaveRecord

    If isFalse(AppShadComponentID) Then
        AppShadComponentID = frm("AppShadComponentID")
        If ExitIfTrue(isFalse(AppShadComponentID), "AppShadComponentID is empty..") Then Exit Function
    End If

    Dim lines As New clsArray
    Dim sqlStr: sqlStr = "SELECT * FROM qryAppShadComponents WHERE AppShadComponentID = " & AppShadComponentID
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim ShadComponent: ShadComponent = rs.fields("ShadComponent"): If ExitIfTrue(isFalse(ShadComponent), "ShadComponent is empty..") Then Exit Function
    Dim ShadComponentFile: ShadComponentFile = rs.fields("ShadComponentFile"): If ExitIfTrue(isFalse(ShadComponentFile), "ShadComponentFile is empty..") Then Exit Function
    
    InstallShadComponent = GetReplacedTemplate(rs, "shad " & ShadComponent)
    InstallShadComponent = GetGeneratedByFunctionSnippet(InstallShadComponent, "InstallShadComponent")
    CopyToClipboard InstallShadComponent
    
    ''C:\Users\User\Desktop\Web Development\next-13-tutorial\src\components\ui\Button.tsx
    Dim ClientPath: ClientPath = rs.fields("ClientPath"): If ExitIfTrue(isFalse(ClientPath), "ClientPath is empty..") Then Exit Function
    Dim filePath: filePath = ClientPath & "src\components\ui\" & ShadComponentFile & ".tsx"
    WriteToFile filePath, InstallShadComponent
    
    ''Install the required packages
    Dim RequiredPackages: RequiredPackages = rs.fields("RequiredPackages")
    
    If Not isFalse(RequiredPackages) Then
        Dim shellArr As New clsArray
    
        shellArr.Add "cmd.exe /k cd /d " & Esc(ClientPath) ''CD into the ClienthPath
        shellArr.Add "npm install " & RequiredPackages
        
        Call Shell(shellArr.JoinArr(" & "))
    End If
    
    ''The hook file
    ''C:\Users\User\Desktop\Web Development\next-13-tutorial\src\hooks\use-toast.tsx
    Dim HookName: HookName = rs.fields("HookName")
    If Not isFalse(HookName) Then
    
        InstallShadComponent = GetReplacedTemplate(rs, "shad " & HookName)
        InstallShadComponent = GetGeneratedByFunctionSnippet(InstallShadComponent, "InstallShadComponent")
        CopyToClipboard InstallShadComponent
        
        filePath = ClientPath & "src\hooks\" & HookName & ".ts"
        WriteToFile filePath, InstallShadComponent
    End If
    
    
End Function
