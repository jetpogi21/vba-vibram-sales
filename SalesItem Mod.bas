Attribute VB_Name = "SalesItem Mod"
Option Compare Database
Option Explicit

Public Function SalesItemCreate(frm As Object, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
        Case 8: ''Cont Form
        Case 9: ''Selector Form
            Dim contFrm As Form: Set contFrm = frm("subform").Form
    End Select

End Function

Public Function dshtSalesItems_ItemCode_AfterUpdate(frm As Form)
    
    Dim ItemCode: ItemCode = frm("ItemCode")
    Dim CustomerCode: CustomerCode = frm("CustomerCode")
    Dim SalesDate: SalesDate = frm("SalesDate")
    
    Dim Price, DiscAllowed
    
    ''Verified Price for this customer's item
    Dim sqlStr: sqlStr = "Select Price, DiscAllowed FROM qrySalesItems WHERE ItemCode = " & Esc(ItemCode) & " AND " & _
        "CustomerCode = " & Esc(CustomerCode) & " AND " & _
        "SalesDate < #" & (SalesDate) & "# AND IsPriceVerified"
        
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    If Not rs.EOF Then
        Price = rs.fields("Price")
        DiscAllowed = rs.fields("DiscAllowed")
        
        frm("Price") = Price
        frm("DiscAllowed") = DiscAllowed
        Exit Function
    End If
    
    ''Verified Price but even if not for this customer
    sqlStr = "Select Price, DiscAllowed FROM qrySalesItems WHERE ItemCode = " & Esc(ItemCode) & " AND " & _
        "SalesDate < #" & (SalesDate) & "# AND IsPriceVerified"
        
    Set rs = ReturnRecordset(sqlStr)
    If Not rs.EOF Then
        Price = rs.fields("Price")
        DiscAllowed = rs.fields("DiscAllowed")
        
        frm("Price") = Price
        frm("DiscAllowed") = DiscAllowed
        Exit Function
    End If
    
    ''Any Price
    sqlStr = "Select Price, DiscAllowed FROM qrySalesItems WHERE ItemCode = " & Esc(ItemCode) & " AND " & _
        "SalesDate < #" & (SalesDate) & "#"
        
    Set rs = ReturnRecordset(sqlStr)
    If Not rs.EOF Then
        Price = rs.fields("Price")
        DiscAllowed = rs.fields("DiscAllowed")
        
        frm("Price") = Price
        frm("DiscAllowed") = DiscAllowed
        Exit Function
    End If
    
    
End Function

