VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormInputMaster 
   Caption         =   "Input Customer Data"
   ClientHeight    =   5004
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   6432
   OleObjectBlob   =   "FormInputMaster.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormInputMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'This form enables users to add a new record or update an existing one in the Customer Master Data.


Private Sub btADD_Click()
'Add data to Masterdata if btADD.caption is ADD
'Update original data record if btAdd.caption is UPDATE

    Dim ctrl As Control
    Dim CustomerTable As ListObject
    Dim NewRowId As Long

    'Set the reference to CustomerTable
    Set CustomerTable = shMaster.ListObjects("CustomerTable")

    'Check if all required fields are not empty
    For Each ctrl In Me.Controls
        If ctrl.Tag = "required" Then
            If ctrl.Value = "" Then
                MsgBox "Please fill out all required fields that are marked with ""*""", vbCritical, "Input Incomplete"
                ctrl.SetFocus
                Exit Sub
            End If
        End If
    Next ctrl
    
    'Check if VAT is in valid format
    If (Me.tbVAT.Value <> "") And (IsNumeric(Me.tbVAT.Value) = False) Then
        MsgBox "Please input a valid number value for VAT", vbCritical, "Invalid VAT"
        Exit Sub
    End If
        
    If Me.btADD.Caption = "ADD" Then
        'Append a new row to CustomerTable
        CustomerTable.ListRows.Add
        'Set the id for the new row
        i = CustomerTable.ListRows.Count
    End If
    
    'Populate the target row with input data
    With CustomerTable
        .ListColumns("Customer").DataBodyRange(i).Value = Me.tbCustomer.Value
        .ListColumns("Company").DataBodyRange(i).Value = Me.tbCompany.Value
        .ListColumns("Address Line 1").DataBodyRange(i).Value = Me.tbAL1.Value
        .ListColumns("Address Line 2").DataBodyRange(i).Value = Me.tbAL2.Value
        .ListColumns("Address Line 3").DataBodyRange(i).Value = Me.tbAL3.Value
        .ListColumns("UID").DataBodyRange(i).Value = Me.tbUID.Value
        .ListColumns("Email").DataBodyRange(i).Value = Me.tbEmail.Value
        If Me.tbVAT.Value <> "" Then
            .ListColumns("VAT").DataBodyRange(i).Value = Me.tbVAT.Value / 100
        Else
            .ListColumns("VAT").DataBodyRange(i).Value = ""
        End If
    End With
    
    If Me.btADD.Caption = "ADD" Then
        MsgBox "A new record has been added to the Customer Database!", vbInformation, "New Customer Added"
    Else
        MsgBox "The record for " & Me.tbCustomer.Value & " from " & Me.tbCompany.Value _
        & " has been updated in the Customer Database!", vbInformation, "Customer Record Updated"
    End If

End Sub


Private Sub btClear_Click()
'Clear all input fields

    Dim ctrl As Control

    For Each ctrl In Me.Controls
        'Check if the control's name starts with "tb" (text box)
        If Left(ctrl.Name, 2) = "tb" Then
            'Clear the value of the control
            ctrl.Value = ""
        End If
    Next ctrl
End Sub



Private Sub btCancel_Click()
'Exit and close the form
    Unload Me
End Sub




