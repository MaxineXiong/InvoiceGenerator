Attribute VB_Name = "MainInvoiceGenerator"
Option Explicit

'Public variables that can be used across all modules and forms
Public i As Long   'Index for selected customer record in CustomerTable


Sub add_data_to_master()
'Open the form FormInputMaster
    FormInputMaster.Show
End Sub


Sub edit_view_master_data()
'Open the form FormViewMaster
    FormViewMaster.Show
End Sub


Sub create_invoice()
'Open the form FormOpenInvoice
    FormOpenInvoice.Show
End Sub


Sub view_invoices()
'Navigate to the "Invoice" sheet to see the list of invoices

    shInvoice.Select
    Range("A1").Select
End Sub


Sub return_dashboard()
'Navigate back to Dashboard

    shDash.Select
    Range("A1").Select
End Sub



