Attribute VB_Name = "ChangeAssetDirectory Mod"
Option Compare Database
Option Explicit

Public Function ChangeAssetDirectoryCreate(frm As Object, FormTypeID)
    
    If FormTypeID = 4 Then
        
        frm.recordSource = ""
        
        With frm("AssetDirectory")
            .OnClick = ""
            .Locked = True
            .ControlSource = ""
        End With
        
        frm("cmdAssetDirectory").OnClick = "=SelectDirectory([Form]," & EscapeString("AssetDirectory") & ")"
        frm("cmdSaveClose").OnClick = "=UpdateApplicationSetting([Form]," & EscapeString("Asset Directory") & "," & _
                                      "[AssetDirectory])"
        
    End If
    
End Function

Public Function UpdateApplicationSetting(frm As Object, ApplicationSettingName, vApplicationSettingValue)
    
    If areDataValid2(frm, "ChangeAssetDirectory") Then
        RunSQL "UPDATE tblApplicationSettings SET ApplicationSettingValue = " & EscapeString(vApplicationSettingValue) & _
               " WHERE ApplicationSettingName = " & EscapeString(ApplicationSettingName)
        DoCmd.Close acForm, frm.Name, acSaveNo
    End If
    
End Function

