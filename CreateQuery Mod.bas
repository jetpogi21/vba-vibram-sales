Attribute VB_Name = "CreateQuery Mod"
Option Compare Database
Option Explicit

Public Function CreateStockCardQuery()
    
    ''Beginning Balance
    ''Transaction Type, 1 Beg Balance, 2 Production, 3 Sales, 4 Adjustment
    
    Dim sqlBEG, sqlPROD, sqlSOLD, sqlADJ, sqlArr As New clsArray
    sqlBEG = "SELECT InventoryID, #1/1/2016# As TransactionDate, BeginningBalance AS Debit, 0 AS Credit, " & _
             EscapeString("Beginning Balance") & " AS Description, Null As Reference, 1 AS TransactionType " & _
             "FROM tblInventories WHERE OnStockCard = -1"
             
    sqlPROD = "SELECT InventoryID, ProductionDate As TransactionDate, ProductionQTY As Debit, 0 AS Credit, " & _
              EscapeString("Baril Production") & " AS Description, Null AS Reference, 2 AS TransactionType " & _
              "FROM tblInventoryProductions WHERE ProductionDate >= #1/1/2016#"
     
    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblSalesInventories"
        .AddFilter "OnStockCard = -1 And SalesDate >= #1/1/2016#"
        .fields = "tblSalesInventories.InventoryID, SalesDate AS TransactionDate, 0 AS Debit, SalesQTY AS Credit, " & _
                  EscapeString("Sales") & " AS Description, DRNumber AS Reference, 3 AS TransactionType"
        .joins.Add GenerateJoinObj("tblSales", "SalesID")
        .joins.Add GenerateJoinObj("tblInventories", "InventoryID")
        sqlSOLD = .sql
    End With
    
    
    sqlADJ = "SELECT InventoryID, AdjustmentDate As TransactionDate, iif(AdjustmentQTY >=0,AdjustmentQTY,0) As Debit, iif(AdjustmentQTY < 0,AdjustmentQTY,0) AS Credit, " & _
             "AdjustmentNote AS Description, Null AS Reference, 4 AS TransactionType FROM tblInventoryAdjusments"
             
    sqlArr.Add sqlBEG: sqlArr.Add sqlPROD: sqlArr.Add sqlSOLD: sqlArr.Add sqlADJ
    Dim sqlStr
    sqlStr = sqlArr.JoinArr(" UNION ALL ")
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .fields = "temp.*,InventoryCode,InventoryDescription,StandardPrice,InventoryClass,InventoryUnit, Null As RunningBalance"
        .joins.Add GenerateJoinObj("tblInventories", "InventoryID")
        .OrderBy = "TransactionDate, TransactionType"
        .SourceAlias = "temp"
        sqlStr = .sql
    End With
    
    Dim qDef As QueryDef
    Dim db As Database
    Set db = CurrentDb
    Set qDef = db.QueryDefs("qryStockCards")
    qDef.sql = sqlStr
    
    MsgBox "qryStockCards created successfully..."
    
End Function
