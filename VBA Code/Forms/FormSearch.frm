VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSearch 
   Caption         =   "Find Customer / Company"
   ClientHeight    =   5532
   ClientLeft      =   60
   ClientTop       =   276
   ClientWidth     =   5184
   OleObjectBlob   =   "FormSearch.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'This form allows users to quickly add relevant customer attributes to the invoice table
'by searching for their customer name or company name in the Customer Master Data.


Private Sub tbCustomer_Change()
'Show matched results in the ListBox as soon as a value is entered into the text box tbCustomer

    Dim r As Long
    Dim CustomerTable As ListObject
    Dim r_matches As Long
    Dim ResultsFound As Boolean
    Dim varCustomer As String
    Dim varCompany As String
    
    'Set the reference to CustomerTable
    Set CustomerTable = shMaster.ListObjects("CustomerTable")
    
    'Clear the listbox
    Me.lbResult.Clear
    'Set initial values
    r_matches = 0
    ResultsFound = False
    
    'Loop through the CustomerTable to find similar Customer and Company names
    With CustomerTable
        For r = 1 To .ListRows.Count
            varCustomer = .ListColumns("Customer").DataBodyRange(r).Value
            varCompany = .ListColumns("Company").DataBodyRange(r).Value
            
            'Check if the entered customer value matches any Customer name from CustomerTable
            If Me.tbCustomer.Value <> "" And Me.tbCompany.Value = "" Then
                If InStr(1, varCustomer, Me.tbCustomer.Value, 1) > 0 Then
                    GoTo ListResults
                End If
            
            'Check if the entered company value matches any Company name from CustomerTable
            ElseIf Me.tbCustomer.Value = "" And Me.tbCompany.Value <> "" Then
                
                If InStr(1, varCompany, Me.tbCompany.Value, 1) > 0 Then
                    GoTo ListResults
                End If
            
            'Check if there's any record from CustomerTable that matches both entered customer and company values
            ElseIf Me.tbCustomer <> "" And Me.tbCompany <> "" Then
                If InStr(1, varCustomer, Me.tbCustomer.Value, 1) > 0 And _
                   InStr(1, varCompany, Me.tbCompany.Value, 1) > 0 Then
                   
ListResults:                    ResultsFound = True
                    'Enable Result list and Select button
                    Me.lbResult.Enabled = True
                    Me.btSelect.Enabled = True
                    'Display matched customer and company name in the listbox
                    Me.lbResult.AddItem
                    Me.lbResult.List(r_matches, 0) = varCustomer
                    Me.lbResult.List(r_matches, 1) = varCompany
                    r_matches = r_matches + 1
                End If
        
            'Check if both tbCustomer and tbCompany fields are empty
            Else
                'Enable Result list and Select button
                Me.lbResult.Enabled = True
                Me.btSelect.Enabled = True
                'If no values are entered into tbCustomer and tbCompany, show the full list of Customer Master Data
                Me.lbResult.AddItem
                Me.lbResult.List(r - 1, 0) = varCustomer
                Me.lbResult.List(r - 1, 1) = varCompany
            End If
        Next r
        
        'If no matching results found
        If (ResultsFound = False) And ((Me.tbCustomer.Value <> "") Or (Me.tbCompany.Value <> "")) Then
            Me.lbResult.AddItem "No matches found"
            'Disable Result list and Select button
            Me.lbResult.Enabled = False
            Me.btSelect.Enabled = False
        End If
        
    End With
    
End Sub




Private Sub tbCompany_Change()
'Show matched results in the ListBox as soon as a value is entered into the text box tbCompany
    
    Call tbCustomer_Change
    
End Sub




Private Sub btSelect_Click()
'Insert the selected customer into the invoice table

    Dim r As Long
    Dim CustomerTable As ListObject
    Dim InvoiceTable As ListObject
    Dim NewRowId As Long
    
    'Set the references to CustomerTable and InvoiceTable
    Set CustomerTable = shMaster.ListObjects("CustomerTable")
    Set InvoiceTable = shInvoice.ListObjects("InvoiceTable")
    
    If Me.lbResult.ListIndex > -1 Then
        'Loop through CustomerTable to find the selected customer's full record
        With CustomerTable
            For r = 1 To .ListRows.Count
                'If a row matches the selected customer
                If .ListColumns("Customer").DataBodyRange(r).Value = Me.lbResult.List(Me.lbResult.ListIndex, 0) And _
                   .ListColumns("Company").DataBodyRange(r).Value = Me.lbResult.List(Me.lbResult.ListIndex, 1) Then
                    
                    'Insert a new row in the second last position of Invoice Table
                    NewRowId = InvoiceTable.ListRows.Count
                    InvoiceTable.ListRows.Add NewRowId
                    
                    'Populate the new row with the relevant customer data from MasterData sheet
                    InvoiceTable.ListColumns("Customer").DataBodyRange(NewRowId).Value = .ListColumns("Customer").DataBodyRange(r).Value
                    InvoiceTable.ListColumns("Company").DataBodyRange(NewRowId).Value = .ListColumns("Company").DataBodyRange(r).Value
                    InvoiceTable.ListColumns("Address Line 1").DataBodyRange(NewRowId).Value = .ListColumns("Address Line 1").DataBodyRange(r).Value
                    InvoiceTable.ListColumns("Address Line 2").DataBodyRange(NewRowId).Value = .ListColumns("Address Line 2").DataBodyRange(r).Value
                    InvoiceTable.ListColumns("Address Line 3").DataBodyRange(NewRowId).Value = .ListColumns("Address Line 3").DataBodyRange(r).Value
                    InvoiceTable.ListColumns("UID").DataBodyRange(NewRowId).Value = .ListColumns("UID").DataBodyRange(r).Value
                    InvoiceTable.ListColumns("VAT").DataBodyRange(NewRowId).Value = .ListColumns("VAT").DataBodyRange(r).Value
                End If
            Next r
        End With
    Else
        'If no record is selected
        MsgBox "Please select a valid record from the list", vbExclamation
    End If
End Sub



Private Sub btClear_Click()
'Clear search entries

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















