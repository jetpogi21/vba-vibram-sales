Attribute VB_Name = "GlobalSetting Mod"
Option Compare Database
Option Explicit

Public Function GetApplicationSetting(ApplicationSettingName) As String
    
    GetApplicationSetting = ELookup("tblApplicationSettings", "ApplicationSettingName = " & EscapeString(ApplicationSettingName), "ApplicationSettingValue")
    
End Function
