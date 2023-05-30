VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOpenInvoice 
   Caption         =   "Create Invoice"
   ClientHeight    =   4536
   ClientLeft      =   12
   ClientTop       =   84
   ClientWidth     =   6804
   OleObjectBlob   =   "FormOpenInvoice.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOpenInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'This form provides users with the capability to create invoices in both PDF and Excel formats
'for selected open invoices. Additionally, users have the option to generate an email to the customer
'with the PDF invoice attached.


'Module-level variables that can be used across all events within this form
Dim OpenInvoices() As Long  'Array of row indices for open invoices in InvoiceTable
Dim InvRowId As Long     'Row index of the selected open invoice in InvoiceTable
Dim PDFFolderPath As String  'Folder path to PDF version of invoices
Dim InvFileName As String    'Invoice file name



Private Sub UserForm_Initialize()
'Populate the list box in FormOpenInvoice with a list of open invoices before displaying the form

    Dim InvoiceTable As ListObject
    Dim r As Long
    Dim c As Long

    'Always jump to page 1 whenever the Userform is open
    Me.MPInvoice.Value = 0
    
    'Set the reference to InvoiceTable
    Set InvoiceTable = shInvoice.ListObjects("InvoiceTable")
    'Clear listbox
    Me.lbOpenInvoices.Clear
    'Set initial value
    c = 0
    
    'Loop through InvoiceTable to check if Status of each record is "" to identify open invoice
    For r = 1 To (InvoiceTable.ListRows.Count - 1)
        If InvoiceTable.ListColumns("Status").DataBodyRange(r).Value = "" Then
            'Display a list of open invoices in the listbox
            Me.lbOpenInvoices.AddItem InvoiceTable.ListColumns("Company").DataBodyRange(r).Value _
                                      & " - " & Format(InvoiceTable.ListColumns("Final Invoiced Amount").DataBodyRange(r).Value, _
                                                       "$#,##0.00")
            ReDim Preserve OpenInvoices(0 To c)
            OpenInvoices(c) = r
            c = c + 1
        End If
    Next r
    
    'If no open invoice has been found
    If Me.lbOpenInvoices.ListCount = 0 Then
        Me.lbOpenInvoices.AddItem "No open invoice so far"
        'Disable CREATE INVOICE button
        Me.btCreateInvoice.Enabled = False
        'Disable the listbox
        Me.lbOpenInvoices.Enabled = False
    End If
    
End Sub




Private Sub btCreateInvoice_Click()
'Create invoices in both PDF and Excel format

    Dim InvoiceTable As ListObject
    Dim PDFFolder As String
    Dim ExcelFolder As String
    Dim ExcelFolderPath As String
    Dim InvNum As Long
    Dim cmt As Comment
    Dim col As ListColumn
    Dim newWB As Workbook
    
    'Set the reference to InvoiceTable
    Set InvoiceTable = shInvoice.ListObjects("InvoiceTable")
    
    'Check if an open invoice is selected in the list box
    If Me.lbOpenInvoices.ListIndex = -1 Then
        MsgBox "Please select an open invoice", vbExclamation, "Open Invoice Not Selected"
        Exit Sub
    End If
    
    'Check if the PDF and Excel folders exist
    PDFFolderPath = ThisWorkbook.Path & "\" & shDash.Range("PDFfolder").Value
    ExcelFolderPath = ThisWorkbook.Path & "\" & shDash.Range("Excelfolder").Value
    PDFFolder = Dir(PDFFolderPath, vbDirectory)
    ExcelFolder = Dir(ExcelFolderPath, vbDirectory)
    
    'Create the Excel folder if it doesn't exist
    If ExcelFolder = "" Then
        MkDir ExcelFolderPath
    End If
    
    'Create the PDF folder if it doesn't exist
    If PDFFolder = "" Then
        MkDir PDFFolderPath
    End If
    
    'Disable screen updating for smoother execution
    Application.ScreenUpdating = False
    
    'Get row index of the selected open invoice in InvoiceTable
    InvRowId = OpenInvoices(Me.lbOpenInvoices.ListIndex)
    
    With InvoiceTable
        'Do not generate invoice file if invoice amount(s) are not entered
        If .ListColumns("Agreed Amount Total").DataBodyRange(InvRowId).Value = 0 Then
            MsgBox "Unable to generate invoice due to incomplete information. Please ensure that the invoice amount has been entered", vbExclamation
            Exit Sub
        End If
        
        'Add Invoice number to the InvoiceTable for the selected open invoice
        InvNum = Excel.WorksheetFunction.Large(.ListColumns("Invoice Number").DataBodyRange, 1) + 1
        .ListColumns("Invoice Number").DataBodyRange(InvRowId).Value = InvNum
        
        'Add Invoice Date to the InvoiceTable for the selected open invoice
        .ListColumns("Invoice Date").DataBodyRange(InvRowId).Value = shDash.Range("TODAY").Value
    End With
    
    'Populate the Template tab with the information from the InvoiceTable
    'based on the comments associated with each cell.
    For Each cmt In shTemp.Comments
        For Each col In InvoiceTable.ListColumns
            If col.Name = cmt.Text Then
                cmt.Parent.Value = col.DataBodyRange(InvRowId).Value
            End If
        Next col
    Next cmt
    
    'Export the invoice template
    InvFileName = InvNum & "_" & shTemp.Range("B12").Value
    Set newWB = Workbooks.Add
    shTemp.Copy before:=newWB.Sheets(1)
    With newWB.Sheets(1)
        'Rename the sheet to InvoiceNum
        .Name = InvNum
        'Delete all comments
        For Each cmt In .Comments
            cmt.Delete
        Next cmt
        'Export the invoice template as PDF
        .ExportAsFixedFormat xlTypePDF, PDFFolderPath & "\" & InvFileName & ".pdf"
    End With
    'Save the invoice template as Excel
    newWB.SaveAs FileName:=ExcelFolderPath & "\" & InvFileName & ".xlsx", FileFormat:=xlOpenXMLStrictWorkbook
    'Close the Excel invoice
    newWB.Close
    
    'Update the status of the selected open invoice to 'Created'
    InvoiceTable.ListColumns("Status").DataBodyRange(InvRowId).Value = "Created"
    
    'Enable screen updating after completing the operation
    Application.ScreenUpdating = True
    
    'Display a success message with the invoice details and folder paths
    MsgBox "The new invoice #" & InvNum & " has successfully been created!" & vbNewLine & _
        vbNewLine & "The Excel invoice is now available in '" & shDash.Range("Excelfolder").Value & "' folder" & vbNewLine & vbNewLine & _
        "The PDF invoice is now available in '" & shDash.Range("PDFfolder").Value & "' folder", vbInformation, "Invoice Created"
        
    'Move on the second page of Multipage control
    Me.MPInvoice.Value = 1
    
End Sub




Private Sub btCancel_Click()
'Close and exit the program
    Unload Me
End Sub




Private Sub btCreateEmail_Click()
'Create an email to customer with PDF invoice attached

    Dim CustomerTable As ListObject
    Dim InvoiceTable As ListObject
    Dim r As Long
    Dim ToEmail As String
    Dim OutlookApp As outlook.Application
    Dim InvEmail As outlook.MailItem
    
    'Set the references to CustomerTable and InvoiceTable
    Set CustomerTable = shMaster.ListObjects("CustomerTable")
    Set InvoiceTable = shInvoice.ListObjects("InvoiceTable")
    ToEmail = ""
    
    'Loop through CustomerTable to get email address for the invoiced customer
    With CustomerTable
        For r = 1 To .ListRows.Count
            If .ListColumns("Customer").DataBodyRange(r).Value = InvoiceTable.ListColumns("Customer").DataBodyRange(InvRowId).Value And _
               .ListColumns("Company").DataBodyRange(r).Value = InvoiceTable.ListColumns("Company").DataBodyRange(InvRowId).Value Then
                ToEmail = Replace(.ListColumns("Email").DataBodyRange(r).Value, " ", "")
            End If
        Next r
    End With
    
    'If email address is not found for the invoice
    If ToEmail = "" Then
        MsgBox "No email address has been found for this customer. Please update it in the Customer Master Data.", vbExclamation, "No Email Address Found"
        Unload Me
        Exit Sub
    End If
    
    'Create and save an email to customer with PDF invoice attached
    Set OutlookApp = New outlook.Application
    Set InvEmail = outlook.CreateItem(olMailItem)
    
    With InvEmail
        .To = ToEmail
        .Subject = "Your Invoice From Maxine"
        .HTMLBody = "Hi " & InvoiceTable.ListColumns("Customer").DataBodyRange(InvRowId).Value & ",<br><br>" _
                    & "Please find the attached invoice.<br>Thanks for choosing Maxine's service.<br><br>Best Regards<br>Maxine Xiong"
        .Attachments.Add PDFFolderPath & "\" & InvFileName & ".pdf"
        .Save
        'Uncomment the following lines to display and send the email automatically
        '.Display
        '.Send
    End With
    
    'Exit the outlook application
    OutlookApp.Quit
    
    'Release the object references
    Set OutlookApp = Nothing
    Set InvEmail = Nothing
    
    MsgBox "Your email has been saved in your Outlook Draft folder, with the PDF invoice attached and ready to be sent.", vbInformation, "Email Created and Saved"
    Unload Me
    
End Sub




Private Sub btNotNow_Click()
'Close and exit the program
    Unload Me
End Sub







