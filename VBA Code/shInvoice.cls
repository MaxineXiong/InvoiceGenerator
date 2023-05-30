VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Private Sub axCollapse_Click()
'Hide/unhide columns that show low-level details

    If Me.axCollapse.Value Then
        'If the checkbox is checked, hide columns that show low-level details
        Me.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    Else
        'If the checkbox is unchecked, unhide columns that show low-level details
        Me.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    End If
End Sub



Private Sub axHide_Click()
'Hide/unhide the rows where invoices are already paid

    Dim InvoiceTable As ListObject
    Dim r As Long
    
    'Disable screen updating for smoother execution
    Application.ScreenUpdating = False
    
    'Set the reference to InvoiceTable
    Set InvoiceTable = Me.ListObjects("InvoiceTable")

    'Check the value of the axHide checkbox
    If Me.axHide.Value Then
        'If the checkbox is checked, iterate through each row in the InvoiceTable
        For r = 1 To InvoiceTable.DataBodyRange.Rows.Count
            'Check if the Status of the current invoice is "Paid"
            If InvoiceTable.ListColumns("Status").DataBodyRange(r).Value = "Paid" Then
                'If the invoice is paid, hide the entire row
                InvoiceTable.ListRows(r).Range.EntireRow.Hidden = True
            End If
        Next r
    Else
        'If the checkbox is unchecked, unhide all rows in the worksheet
        Me.Rows.Hidden = False
    End If
    
    'Enable screen updating after completing the operation
    Application.ScreenUpdating = True
    
End Sub



Private Sub Worksheet_Change(ByVal Target As Range)
' Upon entering a value into the designated orange cell, a search box will be triggered to appear,
' presenting a listbox that showcases all relevant results.

    Dim intersection As Range
    Dim InvoiceTable As ListObject

    'Set the reference to InvoiceTable
    Set InvoiceTable = Me.ListObjects("InvoiceTable")

    'Check if a value has been entered into the designated orange cell
    Set intersection = Excel.Application.intersect(Target, InvoiceTable.DataBodyRange(InvoiceTable.DataBodyRange.Rows.Count, 1))
    
    On Error GoTo ErrorHandler
    If Not (intersection Is Nothing) Then
        'Check if the value entered in the orange cell is not the default message
        If Target.Value <> "Type a customer's name and press Enter to search for the relevant customer data" Then
            'Load and show the search form
            Load FormSearch
            FormSearch.tbCustomer.Value = Target.Value
            FormSearch.Show
            
            'Restore the orange cell back to its original state
            With InvoiceTable.DataBodyRange(InvoiceTable.DataBodyRange.Rows.Count, 1)
                .Value = "Type a customer's name and press Enter to search for the relevant customer data"
                .Interior.Color = RGB(255, 192, 0)
            End With
        End If
    End If

Exit Sub
ErrorHandler:
    'Ignore errors caused by deleting rows in the current sheet
    If Err.Number = -2147352571 Or Err.Number = 13 Then
        Resume Next
    Else
        MsgBox "An error occurred!" & Err.Number, vbExclamation, "Error Found"
    End If

End Sub


