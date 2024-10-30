Attribute VB_Name = "Table Fields"
Option Compare Database
Option Explicit

Public Function LogTable(tblName As String)

    Dim fields() As String
    Dim fieldValues() As String
    Dim ArrayLength
    ArrayLength = 0
    ReDim fields(ArrayLength): ReDim fieldValues(ArrayLength)
    
    fields(0) = "TableName": fieldValues(0) = """" & tblName & """"
    
    Dim tableID
    tableID = InsertAndLog("tblTables", fields, fieldValues)

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(tblName)
    
    ArrayLength = 3
    ReDim fields(ArrayLength): ReDim fieldValues(ArrayLength)
        
    Dim fld As field
    For Each fld In rs.fields
        fields(0) = "FieldName": fieldValues(0) = """" & fld.Name & """"
        fields(1) = "FieldCaption": fieldValues(1) = """" & AddSpaces(fld.Name) & """"
        fields(2) = "FieldTypeID": fieldValues(2) = fld.Type
        fields(3) = "[TableID]": fieldValues(3) = tableID
        InsertAndLog "tblFormFields", fields, fieldValues
    Next fld
    
    
End Function


