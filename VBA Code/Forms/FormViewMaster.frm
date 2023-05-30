VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormViewMaster 
   Caption         =   "View / Edit Customer Master Data"
   ClientHeight    =   5184
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   15744
   OleObjectBlob   =   "FormViewMaster.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormViewMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'This form provides users with the ability to view the Customer Master Data
'and make changes by removing or editing specific customer information within the Customer Master Data.


Private Sub btEdit_Click()
'Edit specific customer's information in the FormInputMaster

    Dim CustomerTable As ListObject
    Dim answer As VbMsgBoxResult
    
    'Set the reference to CustomerTable
    Set CustomerTable = shMaster.ListObjects("CustomerTable")
    
    'Check if an item is selected in the lbMaster ListBox
    If Me.lbMaster.ListIndex > -1 Then
        i = Me.lbMaster.ListIndex + 1
        
        'Trigger FromInputMaster to appear for editing existing customer's information
        Load FormInputMaster
        'Change ADD button's caption to UPDATE
        FormInputMaster.btADD.Caption = "UPDATE"
        'Change form title
        FormInputMaster.lbCustomerFormTitle.Caption = "Edit Customer Information"
        'Change form caption
        FormInputMaster.Caption = "Update Customer Data"
        'Populate FormInputMaster with selected customer's data
        With CustomerTable
            FormInputMaster.tbCustomer.Value = .ListColumns("Customer").DataBodyRange(i).Value
            FormInputMaster.tbCompany.Value = .ListColumns("Company").DataBodyRange(i).Value
            FormInputMaster.tbAL1.Value = .ListColumns("Address Line 1").DataBodyRange(i).Value
            FormInputMaster.tbAL2.Value = .ListColumns("Address Line 2").DataBodyRange(i).Value
            FormInputMaster.tbAL3.Value = .ListColumns("Address Line 3").DataBodyRange(i).Value
            FormInputMaster.tbUID.Value = .ListColumns("UID").DataBodyRange(i).Value
            FormInputMaster.tbEmail.Value = .ListColumns("Email").DataBodyRange(i).Value
            If .ListColumns("VAT").DataBodyRange(i).Value <> "" Then
                FormInputMaster.tbVAT.Value = .ListColumns("VAT").DataBodyRange(i).Value * 100
            Else
                FormInputMaster.tbVAT.Value = ""
            End If
        End With
        'Show FormInputMaster
        FormInputMaster.Show
        'Close FormViewMaster
        Unload Me
    Else
        MsgBox "Please select a record to edit", vbExclamation, "No Record Selected"
    End If
End Sub



Private Sub btDelete_Click()
'Delete specific customer's information from the Customer Master Data

    Dim CustomerTable As ListObject
    Dim ToDelete As VbMsgBoxResult
    Dim SelectedCustomer As String
    Dim SelectedCompany As String
    
    'Set the reference to CustomerTable
    Set CustomerTable = shMaster.ListObjects("CustomerTable")
    
    'Check if an item is selected in the lbMaster ListBox
    If Me.lbMaster.ListIndex > -1 Then  'If nothing is selected, me.lbMaster.ListIndex = -1
        ToDelete = MsgBox("Are you sure you want to delete this record from the customer database?", _
                          vbQuestion + vbYesNo + vbDefaultButton2, "Delete Record?")
        If ToDelete = vbYes Then
            i = Me.lbMaster.ListIndex + 1
            SelectedCustomer = CustomerTable.ListColumns("Customer").DataBodyRange(i).Value
            SelectedCompany = CustomerTable.ListColumns("Company").DataBodyRange(i).Value
            'Delete the selected record from CustomerTable
            CustomerTable.ListRows(i).Delete
            MsgBox "The customer " & SelectedCustomer & " from " & SelectedCompany _
                   & " has been removed from the Customer Database!", vbInformation, "Record Deleted"
        End If
    Else
        MsgBox "Please select a record to delete", vbExclamation, "No Record Selected"
    End If
    
End Sub



Private Sub btCancel_Click()
'Exit and close the form
    Unload Me
End Sub
