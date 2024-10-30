Attribute VB_Name = "Ribbon Helper"
Option Compare Database
Option Explicit

Public MyRibbon As IRibbonUI

Public Function CreateModelRibbon(frm As Form)
    
    ''Initiate the form data
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns
    Dim SetFocus, IsKeyVisible, QueryName, OnFormCreate, SubformName, UserQueryFields, PrimaryKey
    
    ModelID = frm("ModelID")
    Model = frm("Model")
    VerboseName = frm("VerboseName")
    VerbosePlural = frm("VerbosePlural")
    MainField = frm("MainField")
    TableWideValidation = frm("TableWideValidation")
    FormColumns = frm("FormColumns")
    SetFocus = frm("SetFocus")
    IsKeyVisible = frm("IsKeyVisible")
    QueryName = frm("QueryName")
    OnFormCreate = frm("OnFormCreate")
    SubformName = frm("SubformName")
    UserQueryFields = frm("UserQueryFields")
    PrimaryKey = frm("PrimaryKey")

    Dim xmlStr
    xmlStr = "<customUI onLoad=""MyAddInInitialize"" loadImage=""LoadRibbonImage"" xmlns=""http://schemas.microsoft.com/office/2006/01/customui""><ribbon startFromScratch=""false""><tabs>"
    ''Create the tab tag
    xmlStr = xmlStr & "<tab id=" & EscapeString(Model & "Tab") & " label=" & EscapeString(GetFieldCaption(VerboseName, Model)) & " visible=""true"">"
    ''Create Group
    xmlStr = xmlStr & _
            "<group id=" & EscapeString(Model & "Group1") & " label=""Record"">" & _
                "<button id=" & EscapeString(Model & "Add") & " label=""Add New"" imageMso=""FileNewDefault"" size=""large""/>" & _
                "<button id=" & EscapeString(Model & "Edit") & " label=""Edit"" imageMso=""ViewsFormView"" size=""large""/>" & _
                "<button id=" & EscapeString(Model & "Delete") & " label=""Delete"" imageMso=""DataFormDeleteRecord"" size=""large""/>" & _
            "</group>" & _
            "<group idMso=""GroupClipboard""></group>" & _
            "<group idMso=""GroupSortAndFilter""></group>"
    xmlStr = xmlStr & CreateFormSpecificControls(ModelID)
    ''Close the xmlStr
    xmlStr = xmlStr & "</tab></tabs></ribbon></customUI>"
    
    If isPresent("USysRibbons", "RibbonName = " & EscapeString(Model & "Ribbon")) Then
        RunSQL "UPDATE USysRibbons SET RibbonXML = '" & xmlStr & "' WHERE RibbonName = " & EscapeString(Model & "Ribbon")
    Else
        RunSQL "INSERT INTO USysRibbons (RibbonName,RibbonXML) VALUES (" & EscapeString(Model & "Ribbon") & ",'" & xmlStr & "')"
    End If
     
End Function

Public Function CreateFormSpecificControls(ModelID) As String
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblModelButtons WHERE ModelID = " & ModelID & _
                                    " AND HideOnMain = 0")
    Dim ModelButtonID, ModelButton, FunctionName, TableWideFunction, ModelButtonOrder, HideOnMain, HideOnForm
    
    Dim buttonArr As New clsArray, buttonID
    Do Until rs.EOF
    
        ModelButtonID = rs.fields("ModelButtonID")
        ModelButton = rs.fields("ModelButton")
        FunctionName = rs.fields("FunctionName")
        TableWideFunction = rs.fields("TableWideFunction")
        ModelButtonOrder = rs.fields("ModelButtonOrder")
        HideOnMain = rs.fields("HideOnMain")
        HideOnForm = rs.fields("HideOnForm")
        
        buttonID = "customButton" & ModelButtonID

        buttonArr.Add "<button id=" & EscapeString(buttonID) & " label=" & EscapeString(ModelButton) & " onAction=""RunRibbonFunction""/>"
            
        rs.MoveNext
    Loop
    
    If buttonArr.count > 0 Then
         CreateFormSpecificControls = "<group id=" & EscapeString("ModelID") & " label=""Form Specific"">" & buttonArr.JoinArr("") & "</group>"
    End If
    
End Function

Public Function RunRibbonFunction(ctl As IRibbonControl)
    
    Dim frm As Form
    Set frm = Screen.ActiveControl.parent.Form
    
    Dim ModelButtonID
    ModelButtonID = replace(ctl.id, "customButton", "")
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblModelButtons WHERE ModelButtonID = " & ModelButtonID)
    
    Dim ModelButton, FunctionName, TableWideFunction, ModelButtonOrder, HideOnMain, HideOnForm
    
    ModelButton = rs.fields("ModelButton")
    FunctionName = rs.fields("FunctionName")
    TableWideFunction = rs.fields("TableWideFunction")
    ModelButtonOrder = rs.fields("ModelButtonOrder")
    
    If TableWideFunction Then
        Run FunctionName
    Else
        Run FunctionName, frm
    End If
    
End Function

Public Function MyAddInInitialize(Ribbon As IRibbonUI)

    Set MyRibbon = Ribbon
    
End Function

Public Sub LoadRibbonImage(strImage As String, ByRef image)

    Dim strImagePath  As String
    
    strImagePath = CurrentProject.path & "\Assets\Ribbon\" & strImage
    Set image = LoadPicture(strImagePath)

End Sub

Public Sub RefreshRibbon()

    MyRibbon.Invalidate
    
End Sub
