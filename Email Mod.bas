Attribute VB_Name = "Email Mod"
Option Compare Database
Option Explicit



Public Function ComposeMail(frm As Object, Optional TextBoxName = "EmailAddress")

    Dim EmailAddress
    EmailAddress = frm(TextBoxName)
    
    If ExitIfTrue(isFalse(EmailAddress), "Email address is empty...") Then Exit Function
    
    
    Dim olApp As Object, olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    With olMail
        .Display
        .To = EmailAddress
        '.SentOnBehalfOfName = "jet_pradas@yahoo.com"
        '.Subject = Subject
        '.HTMLBody = htmlTxt
        '.Send
    End With

End Function

Public Function ComposeMailContacts(frm As Form)

    Dim EmailAddress, ContactCategoryName, PropertyListID
    EmailAddress = frm("EmailAddress")
    ContactCategoryName = frm("ContactCategoryName")
    PropertyListID = frm("PropertyListID")
    
    If ExitIfTrue(isFalse(EmailAddress), "Email address is empty...") Then Exit Function
    
    Dim recipientArr As New clsArray
    
    If ContactCategoryName = "Solicitor" Then
        Dim rs As Recordset
        Set rs = ReturnRecordset("SELECT * FROM qryPropertyContacts WHERE PropertyListID = " & PropertyListID & " AND ContactCategoryName = 'Solicitor'")
        
        Do Until rs.EOF
            EmailAddress = rs.fields("EmailAddress")
            If Not isFalse(EmailAddress) Then recipientArr.Add EmailAddress
            rs.MoveNext
        Loop
        
    Else
        recipientArr.Add EmailAddress
    End If
    
    
    If ExitIfTrue(recipientArr.count = 0, "Email address is empty...") Then Exit Function
    
    Dim olApp As Object, olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    With olMail
        .Display
        .To = recipientArr.JoinArr(";")
        '.SentOnBehalfOfName = "jet_pradas@yahoo.com"
        '.Subject = Subject
        '.HTMLBody = htmlTxt
        '.Send
    End With
    
    
End Function


