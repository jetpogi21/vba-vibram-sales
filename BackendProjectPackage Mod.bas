Attribute VB_Name = "BackendProjectPackage Mod"
Option Compare Database
Option Explicit

Public Function BackendProjectPackageCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            frm("Packages").Top = frm("Packages").Left
            frm("Packages").HorizontalAnchor = acHorizontalAnchorBoth
            frm("Packages").VerticalAnchor = acVerticalAnchorBoth
            frm("cmdExtractBackendPackages").Top = GetBottom(frm("Packages")) + InchToTwip(0.1)
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function ExtractBackendPackages(frm As Form)
    
    Dim Packages: Packages = frm("Packages")
    Dim BackendProjectID: BackendProjectID = frm("BackendProjectID")
    
    Dim packageJSON As Object: Set packageJSON = JsonConverter.ParseJson(Packages)
    
    ' Replace the name and type property
    Dim key, item
    Dim dependencies As Object: Set dependencies = packageJSON("dependencies")
    
    Dim fields As New clsArray: fields.arr = "NPMPackage"
    Dim fieldValues As New clsArray
    
    Dim NPMPackageID
    For Each key In dependencies.Keys
        item = dependencies(key)
        ''Check if the key exists from the NPMPackages
        NPMPackageID = ELookup("tblNPMPackages", "NPMPackage = " & Esc(key), "NPMPackageID")
        If Not isFalse(NPMPackageID) Then
            ''Insert to the tblBackendNPMPackages table
            RunSQL "INSERT INTO tblBackendNPMPackages (BackendProjectID,NPMPackageID) values (" & BackendProjectID & "," & NPMPackageID & ")"
        Else
            Set fieldValues = New clsArray
            fieldValues.Add key
            UpsertRecord "tblNPMPackages", fields, fieldValues
            ''MsgBox Esc(key) & " is missing as a dependency package."
        End If
    Next key
    
    Dim devDependencies As Object: Set devDependencies = packageJSON("devDependencies")
    For Each key In devDependencies.Keys
        item = devDependencies(key)
        NPMPackageID = ELookup("tblNPMPackages", "NPMPackage = " & Esc(key), "NPMPackageID")
        If Not isFalse(NPMPackageID) Then
            ''Insert to the tblBackendNPMPackages table
            RunSQL "INSERT INTO tblBackendNPMPackages (BackendProjectID,NPMPackageID) values (" & BackendProjectID & "," & NPMPackageID & ")"
        Else
            Set fieldValues = New clsArray
            fieldValues.Add key
            UpsertRecord "tblNPMPackages", fields, fieldValues
            ''MsgBox Esc(key) & " is missing as a dev dependency package."
        End If
    Next key
    
    If IsFormOpen("frmBackendProjects") Then
        Forms("frmBackendProjects")("subBackendNPMPackages").Form.Requery
    End If
    
End Function
