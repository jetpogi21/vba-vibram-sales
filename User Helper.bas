Attribute VB_Name = "User Helper"
Option Compare Database
Option Explicit

Public Function OpenAccessibleMenus(frm As Form)

    Dim UserGroup, UserGroupID
    UserGroup = frm.subform.Form("UserGroup")
    UserGroupID = frm.subform.Form("UserGroupID")
    
    If isFalse(UserGroup) Then
        ShowError "Please select a user group record.."
        Exit Function
    End If
    
    If UserGroup = "Administrator" Then
        ShowError "[Administrator] groups access don't need to be set.."
        Exit Function
    End If
    
    DoCmd.OpenForm "mainUserGroupButtons", , , "UserGroupID = " & UserGroupID
    
End Function

Public Function SetGlobalsAidan()

    UpdateGlobalSetting "rptPackSheets_filePath", "\\nas\ERP-Staging\assets\Reports\Logistics\Pack Sheets"
    UpdateGlobalSetting "rptPickSheets_filePath", "\\nas\ERP-Staging\assets\Reports\Logistics\Pick Sheets"
    UpdateGlobalSetting "rptIntermediateLabels_filePath", "\\nas\ERP-Staging\assets\Labels\Intermediate"
    UpdateGlobalSetting "rptPrintH_filePath", "\\nas\ERP-Staging\assets\Labels\Heavy Weight"
    UpdateGlobalSetting "rptPurchaseOrderCSV_filePath", "\\nas\ERP-Staging\assets\Reports\Purchasing\csv"
    UpdateGlobalSetting "rptPurchaseOrderProforma_filePath", "\\nas\ERP-Staging\assets\Reports\Purchasing\Proforma"
    UpdateGlobalSetting "rptHighValueLabels_filePath", "\\nas\ERP-Staging\assets\Labels\High Value"
    UpdateGlobalSetting "rptInternationalLabels_filePath", "\\nas\ERP-Staging\assets\Labels\International"
    UpdateGlobalSetting "rptShippingMethodLabels_filePath", "\\nas\ERP-Staging\assets\Labels\Shipping Method"
    UpdateGlobalSetting "rptSplitOrders_filepath", "\\nas\ERP-Staging\assets\Reports\Logistics\Split Orders"
    UpdateGlobalSetting "Application_EmailTemplatesFilePath", "\\nas\ERP-Staging\assets\Media\System"
    UpdateGlobalSetting "Application_ImportCSV_filePath", "\\nas\ERP-Staging\assets\Imports\"
    UpdateGlobalSetting "rptShelfLocationLabels_filePath", "\\nas\ERP-Staging\assets\Labels\Shelf Locations\"
    UpdateGlobalSetting "systemProductImages_filePath", "\\nas\ERP-Staging\assets\Media\Products\"
    UpdateGlobalSetting "SystemAdminEmailAddress", "coreelectronicssystem@gmail.com"

End Function

Private Sub UpdateGlobalSetting(ByVal SettingName As String, ByVal SettingValue As String)

    Dim sqlStr As String
    
    sqlStr = "UPDATE tblGlobalSettings SET tblGlobalSettings.GlobalSettingValue = '" & SettingValue & "' WHERE (((tblGlobalSettings.GlobalSetting)='" & SettingName & "'));"
    
    CurrentDb.Execute (sqlStr)

End Sub


Public Function RefreshObjects(frm As Form)

    Dim UserID: UserID = frm("UserID")
    If IsNull(UserID) Then Exit Function
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    ''TABLE: tblFormForRights Fields: FormForRightsID|ModelName|FormName|FormType|Timestamp|CreatedBy|RecordImportID|ModelCaption
    ''TABLE: tblUserRights Fields: UserRightID|User|ModelName|CanView|CanEdit|CanAdd|CanDelete|Timestamp|CreatedBy|RecordImportID
    
    ''Get all the unqiue ModelName from tblFormForRights
    ''Insert into tblUserRights, User = UserID, ModelName = ModelName
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblFormForRights"
        .fields = UserID & " AS User, ModelName"
        .GroupBy = "ModelName"
        sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblUserRights"
        .fields = "User,ModelName"
        .insertSQL = sqlStr
        .InsertFilterField = "User,ModelName"
        rowsAffected = .Run
    End With
    
    frm("subUserUserRights").Form.Requery
    
End Function


Public Function CanViewChange(frm As Form)

    Dim CanView, CanAdd, CanEdit, CanDelete
    
    CanView = frm("CanView")
    
    If Not CanView Then
    
        'frm("CanAdd") = 0
        frm("CanEdit") = 0
        frm("CanDelete") = 0
        
    End If
    
End Function

Public Function OtherCansChange(frm As Form)

    Dim CanView, CanAdd, CanEdit, CanDelete
    Dim OtherCan
    
    OtherCan = frm("CanEdit")
    
    If Not OtherCan Then OtherCan = frm("CanDelete")
    
    If OtherCan Then
    
        frm("CanView") = -1
        
    End If
    
End Function

Public Function FormUserOnCurrent(frm As Form)

    SetFocusOnForm frm, "UserName"
    RefreshObjects frm
    
End Function


