Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz, Committ As String
Dim RST As DAO.Recordset
Dim IDX As Long
Private Sub BuyerRefCBO_Click()
Call RemoteBuyerRef
End Sub
Private Sub InvoiceList_Click()
Call RemoteInvoice
End Sub
Public Sub RemoteBuyerRef()
    Me.Form.RecordSource = ""
    Me.Detail.Visible = True
    SQLz = "SELECT InvoiceRequestID As Req, nz(InvoiceNumber,'Request') as Invoice, InvoiceRequestDate As RequestDate, BuyerID " & _
           "From RmktInvoiceRequestTbl WHERE BuyerId =" & Me.BuyerRefCBO & " AND Not isnull(Accepted)  " & _
           "ORDER BY InvoiceRequestID DESC; "

    Me.InvoiceList.RowSource = SQLz
    Me.UnitCount.Visible = False
    Me.UnitCountLabel.Visible = False
    Me.Cover1.Visible = True

End Sub
Public Sub RemoteInvoice()
If Me.Detail.Visible = False Then Me.Detail.Visible = True
SQLz = "SELECT * From RMKT_Invoices " & _
       "WHERE BuyerID=" & Me.BuyerRefCBO & " AND  InvoiceRequestID=" & Me.InvoiceList.Column(0) & ";"
   Me.Form.RecordSource = SQLz
    If Me.UnitCount.Visible = False Then Me.UnitCount.Visible = True
    If Me.UnitCountLabel.Visible = False Then Me.UnitCountLabel.Visible = True
  Me.UnitCountLabel.Visible = True
  Me.UnitCount.Visible = True
  Me.Cover1.Visible = False

End Sub

Private Sub BillLog_Click()
    DoCmd.close acReport, "RemarketingBillLog"
    DoCmd.OpenReport "RemarketingBillLog", acViewReport
End Sub

Private Sub ExportLog2Excel_Click()
    FormatExcel "RmktBilllogq", "C:\Faas\" & FAUserName & "\Temp\RmktBilllogq.xlsx", "BillLog"
End Sub

Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    Me.UnitCount.Visible = False
    Me.UnitCountLabel.Visible = False
    Me.Detail.Visible = False
End Sub
Private Sub Form_Current()
Me.Form.RecordSource = ""
End Sub
Public Sub SetSentStatus(StatusX As String)
DoCmd.SetWarnings False
If StatusX = "S" Then
SQLz = "UPDATE RMKT " & _
               "SET InvoicePaidDate = #" & Me.InvoicePaidDate & "# " & _
               "WHERE InvoiceRequestNum = " & Me.InvoiceRequestID & " AND IsNull(InvoicePaidDate) ;"
        DoCmd.RunSQL (SQLz)
End If
SQLz = "UPDATE RMKT " & _
               "SET InvoiceSentDate = #" & Me.InvoiceSentDate & "# " & _
               "WHERE InvoiceRequestNum = " & Me.InvoiceRequestID & " AND IsNull(InvoicePaidDate) ;"
        DoCmd.RunSQL (SQLz)
SQLz = "UPDATE Units INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef " & _
               "SET Units.UnitStatus = '" & StatusX & "' " & _
               "WHERE RMKT.InvoiceRequestNum = " & Me.InvoiceRequestID & " AND Units.UnitStatus <> 'S' ;"
    DoCmd.RunSQL (SQLz)
    Me.ReMarketInvoiceSub.Form.Requery
    DoCmd.SetWarnings True
End Sub
Private Sub InvoiceReprint_Click()
Me.Form.Refresh
    DoCmd.OpenReport "RmktInvoiceRePrint", acViewPreview
End Sub

Private Sub InvoiceSentDate_AfterUpdate()
If IsDate(Me.InvoiceSentDate) And Not IsDate(Me.InvoicePaidDate) Then Call SetSentStatus("PS")
If IsDate(Me.InvoiceSentDate) And IsDate(Me.InvoicePaidDate) Then Call SetSentStatus("S")
End Sub

Private Sub Makelink_Click()
'First, set a Reference to the Microsoft Office XX.X Object Library
    Dim strButtonCaption As String, strDialogTitle As String
    Dim strHyperlinkFile As String, strSelectedFile As String

    'Define your own Captions if necessary
    strButtonCaption = "Save Hyperlink"
    strDialogTitle = "Select File to Create Hyperlink to"

    With Application.FileDialog(msoFileDialogFilePicker)
        With .Filters
            .Clear
            .Add "All Files", "*.*"     'Allow ALL File types
        End With
        'The Show Method returns True if 1 or more files are selected
        .AllowMultiSelect = False       'Critical Line
        .FilterIndex = 1    'Database files
        .ButtonName = strButtonCaption
        .InitialFileName = vbNullString
        .InitialView = msoFileDialogViewDetails     'Detailed View
        .Title = strDialogTitle
        If .Show Then
            For Each varitem In .SelectedItems    'There will only be 1
                'Display Text + # + Address of Hyperlink (Display Text#Hyperlink Address)
                strSelectedFile = varitem
                strHyperlinkFile = fGetBaseFileName(strSelectedFile) & "#" & strSelectedFile
                Me.BillLink = strHyperlinkFile
            Next varitem
        End If
    End With

End Sub

Private Sub Print_Invoice_Click()
Me.Form.Refresh
    If IsNull(Me.InvoiceSentDate) Then
        MsgBox "A Invoice Request Date is required", vbCritical, "No Date"
        Exit Sub
    End If
    If Me.ReMarketInvoiceSub!unitstatus = "S" Or Me.ReMarketInvoiceSub!unitstatus = "KS" Then
        MsgBox "These units have a status of = " & Me.ReMarketInvoiceSub!unitstatus & "  (SOLD)" & vbNewLine & "I will open a 'reprint'", vbInformation, "Sold Status"
        DoCmd.OpenReport "RmktInvoiceReprint", acViewPreview
        Exit Sub
    End If
    IDX = Me.InvoiceRequestID
    Committ = MsgBox("The Invoice 'Paid Date' will be updated to '" & Nz(Me.InvoicePaidDate, Date) & "' date, Do you agree?", vbQuestion + vbYesNo + vbCritical, "Update Date?")
    If Committ = vbNo Then Exit Sub
    If IsNull(Me.InvoicePaidDate) Then Me.InvoicePaidDate = Date
        Me.ReMarketInvoiceSub.Requery
    Call Update_RMKT(0, IDX) ' set to zero for testing else 1
    Me.Form.Refresh
    DoCmd.OpenReport "RmktInvoice", acViewPreview
End Sub

Public Function Update_RMKT(Optional lvl As Integer, Optional lIDX As Long)
        If lvl = 0 Then Exit Function ' for testing
        If lIDX = 0 Then lIDX = Me.InvoiceRequestID
        SQLz = "UPDATE RMKT Set "
        SQLz = SQLz & " InvoiceRequestNum = " & IDX & " "
        SQLz = SQLz & ",InvoiceSentDate = #" & Me.InvoiceSentDate & "# "
        SQLz = SQLz & ",InvoiceRequestDate = #" & Me.InvoiceRequestDate & "# "
        SQLz = SQLz & ",BuyerREF= " & Me.BuyerRef & " "
        SQLz = SQLz & ",RmkSm='" & Me.SP & "' "
        If lvl = 1 Or lvl = 3 Then SQLz = SQLz & ",InvoicePaidDate = #" & Me.InvoicePaidDate & "# "
        If lvl = 2 Or lvl = 3 Then SQLz = SQLz & ",RmktBOSSent = #" & Date & "# "
        SQLz = SQLz & "WHERE UnitREf in (SELECT UnitID FROM RMKT_BuyersOrders WHERE BuyersOrderID = " & lIDX & ");"

        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)

        SQLz = "UPDATE Rmkt INNER JOIN RMKT_BuyersOrders ON Rmkt.UnitRef = RMKT_BuyersOrders.UnitID "
        SQLz = SQLz & "SET Rmkt.SoldPrice = RMKT_BuyersOrders.salePrice "
        SQLz = SQLz & "WHERE RMKT_BuyersOrders.BuyersOrderID=" & lIDX & ";"
        DoCmd.RunSQL (SQLz)

        SQLz = "UPDATE Units SET UnitStatus = 'S' " & _
               "WHERE UnitID in (SELECT UnitID FROM vw_BuyersOrder WHERE (Status='R' Or Status='P' Or Status='PS') and InvoiceRequestID = " & lIDX & ");"
        DoCmd.RunSQL (SQLz)
        SQLz = "UPDATE Units SET UnitStatus = 'KS' " & _
               "WHERE UnitID in (SELECT UnitID FROM vw_BuyersOrder WHERE (Status='K') and InvoiceRequestID = " & lIDX & ");"
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        If lvl = 3 Then
           Call KillFiles("\\fa-poseidon\Atlaas\Temp", "B*_*.pdf")
           DoCmd.OutputTo acOutputReport, "RemarketingBillOfSale", acFormatPDF, "\\fa-poseidon\Atlaas\Temp\BOS_" & lIDX & ".PDF"
           Call SendBOEmail("BuyersOrderPaid", lIDX, "\\fa-poseidon\Atlaas\Temp\BOS_" & InvoiceRequestID & ".PDF")
        End If
End Function
Private Sub MarkUnitsPaid_Click()
    If IsNull(Me.InvoicePaidDate) Or InvoicePaidDate = "" Then
        MsgBox "A Invoice Paid Date is required", vbCritical, "No Paid Date"
        Exit Sub
    End If
    IDX = Me.InvoiceRequestID
    Committ = MsgBox("The Invoice Paid Date (" & Me.InvoicePaidDate & ") " & vbNewLine & "will be posted to the units within this Invoice" & _
    "" & vbNewLine & vbNewLine & "Do you agree?", vbQuestion + vbYesNo, "We Got Paid")
    If Committ <> vbYes Then Exit Sub
        Call Update_RMKT(1, IDX) 'set to zero for testing else 1
        Me.Form.Requery
        'Me.InvoiceList = IDx
         Call RemoteInvoice
End Sub

Private Sub Print_BillOfSale_Click()
        If Me.Form.Dirty Then Me.Form.Refresh
        IDX = Me.InvoiceRequestID
    If IsNull(Me.InvoicePaidDate) Or InvoicePaidDate = "" Then
        MsgBox "You must have a Invoice Paid Date", , "No Paid Date"
        Exit Sub
    End If
        If DMin("Status", "vw_BuyersOrder", "InvoiceRequestID = " & IDX) <> "D" Then
        MsgBox "Unit Status will be updated to 'Sold' Status, accordingly", vbInformation, "Unit Status Update"
        Call Update_RMKT(3, IDX) 'set to 0 for testing else 3
        End If
        DoCmd.OpenReport "RemarketingBillOfSale", acViewPreview

  End Sub

Private Sub ShowBillOfSale_Click()
DoCmd.close acReport, "RemarketingBillOfSale"
DoCmd.OpenReport "RemarketingBillOfSale", acViewPreview

End Sub
