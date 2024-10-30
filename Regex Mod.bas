Attribute VB_Name = "Regex Mod"
Option Compare Database
Option Explicit

Public Function RegexCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

'Public Function GetPropertyID(toMatch) As String
'
'    ''https://rpp.rpdata.com/rpp/property/detail.html?propertyId=16118042
'    Dim pattern: pattern = ".+propertyID=(\d+)"
'    ''get the ID=
'
'    Dim regex As Object
'    Dim theMatches As Object
'    Dim Match As Object
'    Set regex = New RegExp
'
'    regex.pattern = pattern
'    regex.Global = False
'    regex.IgnoreCase = True
'
'    If regex.Test(toMatch) Then
'        TestRegex = regex.Replace(toMatch, "$1")
'    End If
'
'End Function

''GetSanitizedName("Sai - Enquired 50 Pechey Street Chermside")
Public Function GetSanitizedName(toMatch) As String
    
    GetSanitizedName = toMatch
    Dim pattern: pattern = "^([a-zA-Z0-9 ]+)[ -_]+Enq.*"
    
    Dim regex As Object
    Dim theMatches As Object
    Dim match As Object
    Set regex = New RegExp
    
    regex.pattern = pattern
    regex.Global = False
    regex.IgnoreCase = True

    If regex.Test(toMatch) Then
        GetSanitizedName = regex.replace(toMatch, "$1")
    End If
    

End Function

Public Function GetMatchedPatterns(Text, pattern As String, Optional DebugMode As Boolean = False) As clsArray

    Dim regex As Object
    Set regex = New RegExp

    regex.pattern = pattern
    regex.Global = True
    regex.IgnoreCase = True
    
    Dim theMatches As Object, match As Object
    Set theMatches = regex.Execute(Text)
    
    Dim MatchArr As New clsArray
    
    For Each match In theMatches
      MatchArr.Add match.value
      If DebugMode Then Debug.Print match.value
    Next
    
    Set GetMatchedPatterns = MatchArr

End Function

Public Function RemoveMatchedPattern(Text, pattern As String, Optional DebugMode As Boolean = False) As String

    Dim regex As Object
    Set regex = New RegExp

    regex.pattern = pattern
    regex.Global = True
    regex.IgnoreCase = True
    
    If regex.Test(Text) Then
        RemoveMatchedPattern = regex.replace(Text, "")
        Exit Function
    End If
    
    RemoveMatchedPattern = Text
    
End Function

Public Function ReplaceMatchedPattern(Text, pattern As String, replacement As String, Optional DebugMode As Boolean = False) As String
    
    If pattern = ",\n\}" Then
        Debug.Print
    End If
    
    Dim regex As Object
    Set regex = New RegExp

    regex.pattern = pattern
    regex.Global = True
    regex.IgnoreCase = False
    
    If regex.Test(Text) Then
        ReplaceMatchedPattern = regex.replace(Text, replacement)
        Exit Function
    End If
    
    ReplaceMatchedPattern = Text
    
End Function


Public Function GetMatchedGroup(Text, pattern As String, Optional DebugMode As Boolean = False) As String
    
    Dim regex As Object
    Set regex = New RegExp

    regex.pattern = pattern
    regex.Global = True
    regex.IgnoreCase = True
    
    Dim theMatches As Object, match As Object
    Set theMatches = regex.Execute(Text)
    
    Dim MatchArr As New clsArray
    
    For Each match In theMatches
        Dim MatchValue: MatchValue = match.value
        GetMatchedGroup = regex.replace(MatchValue, "$1")
        Exit Function
        If DebugMode Then Debug.Print match.value
    Next

End Function

Public Function GetAllMatchedGroup(Text, pattern As String, Optional DebugMode As Boolean = False) As clsArray

    Dim regex As Object
    Set regex = New RegExp

    regex.pattern = pattern
    regex.Global = True
    regex.IgnoreCase = True
    
    Dim theMatches As Object, match As Object
    Set theMatches = regex.Execute(Text)
    
    Dim MatchArr As New clsArray
    
    For Each match In theMatches
       
        MatchArr.Add match.value
        If DebugMode Then Debug.Print match.value
        
    Next
    
    Set GetAllMatchedGroup = MatchArr

End Function
