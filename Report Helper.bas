Attribute VB_Name = "Report Helper"
Option Compare Database
Option Explicit

Public Function RunReport(rptCaption)

    Select Case rptCaption
        Case "Recepients":
            DoCmd.OpenForm "frmChildRecipientReport", , , , acFormAdd
        Case "Donation Requests":
            DoCmd.OpenForm "frmDonationRequestReport", , , , acFormAdd
        Case "Books And Petty Fees":
            DoCmd.OpenForm "frmPettyFeeReport", , , , acFormAdd
        Case "Textbook Stationeries":
            DoCmd.OpenForm "frmTextbookStationeryReport", , , , acFormAdd
        Case "Voucher Allocation Amounts":
            DoCmd.OpenForm "frmVoucherAllocationAmountReports", , , , acFormAdd
        Case "Uniform List":
            DoCmd.OpenForm "frmChildUniformReport", , , , acFormAdd
    End Select
    
End Function

Public Function GetOfficeInfo(OfficeID) As String
 
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblOffices WHERE OfficeID = " & OfficeID)
    
    Dim officeArr As New clsArray
    officeArr.Add rs.fields("OfficeAddress")
    officeArr.Add rs.fields("ContactNumber")
    officeArr.Add rs.fields("EmailAddress")
    
    GetOfficeInfo = officeArr.JoinArr(" | ")
    
End Function
