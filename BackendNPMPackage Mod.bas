Attribute VB_Name = "BackendNPMPackage Mod"
Option Compare Database
Option Explicit

Public Function BackendNPMPackageCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Function ExtractPackageNames(jsonString) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Define the regular expression pattern to match package names
    regex.pattern = """([^""]+)"""

    Dim matches As Object
    Set matches = regex.Execute(jsonString)
    
    Dim packageNames As String
    packageNames = ""
    
    ' Iterate through the matches and concatenate package names
    Dim match As Object
    For Each match In matches
        If packageNames <> "" Then
            packageNames = packageNames & "," & match.SubMatches(0)
        Else
            packageNames = match.SubMatches(0)
        End If
    Next match
    
    ExtractPackageNames = packageNames
End Function



