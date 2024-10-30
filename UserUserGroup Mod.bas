Attribute VB_Name = "UserUserGroup Mod"
Option Compare Database
Option Explicit

Public Function UserUserGroupValidation(frm As Form) As Boolean
    ''If the UserGroupID Col 1 is Administrator Then
    ''Other UserGroup should not be allowed
    ''It will also depend if the current record is a new one or an existing one
    Dim UserUserGroupID, UserGroupID, UserID, fltrStr
    
    UserUserGroupID = frm("UserUserGroupID")
    UserGroupID = frm("UserGroupID")
    UserID = frm("UserID")
    
    If UserGroupID = 1 Then
        ''Lookup if there's already an exisiting access account
        fltrStr = "UserID = " & UserID
        If Not IsNull(UserUserGroupID) Then fltrStr = fltrStr & " And UserUserGroupID <> " & UserUserGroupID
        
        If isPresent("tblUserUserGroups", fltrStr) Then
            ShowError "This user already has an access group." & vbCrLf & _
                "Please delete any existing group first to continue using the administrator group.."
            UserUserGroupValidation = False
            Exit Function
        End If
    Else
        ''Lookup if there's already an admin account
        fltrStr = "UserID = " & UserID & " And UserGroupID = 1"
        If Not IsNull(UserUserGroupID) Then fltrStr = fltrStr & " And UserUserGroupID <> " & UserUserGroupID
        
        If isPresent("tblUserUserGroups", fltrStr) Then
            ShowError "This user is already an administrator so there's no need to add more access.."
            UserUserGroupValidation = False
            Exit Function
        End If
        
    End If
    
    UserUserGroupValidation = True
            
End Function
