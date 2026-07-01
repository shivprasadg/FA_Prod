Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz    As String
Private Sub AddBillBtn_Click()
If Nz(Me.SchID, Me.Parent.SchID) & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "NewBill"
    DoCmd.OpenForm "NewBill", , , , acFormAdd
    Forms!NewBill.Form.BillSchRef = Me.Parent.SchID
    Forms!NewBill.Form.GroupID = Me.Parent.GroupID
    On Error Resume Next
    Forms!NewBill.Form.MLNo = Me.Parent.SelLease
End Sub

Private Sub BillType_Click()
   Me.BillTypeID = Me.BillType.Column(1)
   SQLz = "UPDATE B SET B.BillTypeID = T.BillTypeID"
   SQLz = SQLz & " FROM Bills AS B"
   SQLz = SQLz & " INNER JOIN Lookup_BillType AS T on B.BillType = T.BillType"
   SQLz = SQLz & " WHERE B.BillTypeID Is Null"
   DoCmd.SetWarnings False
 '  DoCmd.RunSQL (SQLz)
   DoCmd.SetWarnings True
End Sub

Private Sub DeleteMe_Click()
    Dim Response As String
    Response = MsgBox("Are you sure you want to Delete Bill No: " & Me.BillNo & "?", vbQuestion + vbYesNo + vbCritical, "Oh No Mr. Bill!")
    If Response = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunCommand acCmdDeleteRecord
        DoCmd.SetWarnings True
        Me.Form.Requery
    End If
End Sub



Private Sub Form_Current()

    Me.Parent.BillID = Me.BillID
    Me.Parent.GroupID = Me.GroupID
    On Error GoTo FixMe
    Me.txtCurrent = Me.BillID
FixMe:
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
            .Add "All Files", "*.*"                 'Allow ALL File types
        End With
                                                    'The Show Method returns True
                                                    'if 1 or more files are selected
        .AllowMultiSelect = False                   'Critical Line
        .FilterIndex = 1                            'Database files
        .ButtonName = strButtonCaption
        .InitialFileName = vbNullString
        .InitialView = msoFileDialogViewDetails     'Detailed View
        .Title = strDialogTitle
        If .Show Then
            For Each varitem In .SelectedItems      'There will only be 1
                                                    'Display Text + # + Address of Hyperlink
                                                    '(Display Text#Hyperlink Address)
                strSelectedFile = varitem
                strHyperlinkFile = fGetBaseFileName(strSelectedFile) & "#" & strSelectedFile
                Me.BillLink = strHyperlinkFile
            Next varitem
        End If
    End With

End Sub

Private Sub ReviewSchedule_Click()
    DoCmd.OpenForm "SignOffDashboardSingle", acNormal, , "SchId=" & Me.SchID & ""
End Sub
Private Sub InvoiceToExcel_Click()
    Call MakeInvoice(1)
End Sub
Private Sub Runbill_Click()
    Call MakeInvoice(0)
End Sub
Public Function MakeInvoice(ToExcel As Boolean)

    Dim FundStatus As String
    If Me.Dirty Then
        Me.Dirty = False
        Me.Form.Requery
    End If
    Dim HeaderTypex As String
    FundStatus = DMax("Assigned_Indicator", "SchGrp", "sGrpSchID=" & Me.SchID)
    Me.Parent.BillID = Me.BillID
    Me.Parent.GroupID = Me.GroupID
    Dim ReportX, ReportX2, OpenArgsX As String
    Dim RecX As Integer

    If Me.BillType = "PerDiem" Then ReportX = "PerDiemInvoice"
    If Me.BillType = "PerDiem" And Me.GroupBill Like "*Group*" Then ReportX = "PerDiemInvoiceGroup"

    If Me.BillType = "BankBill" Then ReportX = "BankBillReg"
    If Me.BillType = "BankFA Remit" Then ReportX = "BankBill_FARemit"
    If Me.BillType = "BankRentOnly" Then ReportX = "BankBillRentOnly"
    If Me.BillType = "BankBill" And Me.ClientID = 4636 Then ReportX = "FABankBill_ADM"
    If Me.BillType = "BankBill" And Me.ClientID = 4029 And FunStatus = "F" Then ReportX = "FABankBill_Toyota"
    If Me.BillType = "BankBill" And Me.GroupID & "test" <> "test" Then ReportX = "FABankBillGroup"

    If Me.BillType = "Rent" Then ReportX = "FARentBillNew"
    If Me.BillType = "Rent" And Me.ClientID = 4029 And FunStatus = "F" Then ReportX = "FABankBill_Toyota"
    If Me.BillType = "Rent Group" Then ReportX = "FARentBillNew_Group"

    If Me.BillType = "PDSurrender" Then ReportX = "PerDiemSurrenderInvoice"
    If Me.BillType = "PDSurrLOC" Then ReportX = "PerDiemSurrenderInvoiceLocation"

    If Me.BillType = "AnnRent" Then ReportX = "ToyotaAncillaryRent"
    If Me.BillType = "ServFeeOnly" Then ReportX = "BillServFeeOnly"
    If Me.BillType = "Warranty" Then ReportX = "BillWarranty"

    'Run The Bill -----------------------------------------
    DoCmd.close acReport, ReportX
    On Error GoTo FixMe
    If ToExcel = True Then
        HeaderTypex = BillType
        If BillType = "PerDiem" Then HeaderTypex = "Interim Rent"
        Dim XLS As String

   '----Default SQL
        SQLz = "Select"
        SQLz = SQLz & " BillNo As [Invoice Number]"
        SQLz = SQLz & ",BankContract as [Assignee Contract #]"
        SQLz = SQLz & ",BankName as [Remit To]"
        SQLz = SQLz & ",Replace(Replace(Replace(RemitAddr,'<div>',''),'</div>',''),'&amp;','&') as [Remit Address]"
        SQLz = SQLz & ",MLNO as [Lease:]"
        SQLz = SQLz & ",SchNo & "" | "" & SGRPNo as [Sch | Grp]"
        SQLz = SQLz & ",Billdt as [Inv Date]"
        SQLz = SQLz & ",BillDueDt as [Due Date]"
        SQLz = SQLz & ",[Billfrom] & "" Through "" & [Billto] AS [Rent Bill for Period:]"
        SQLz = SQLz & ",LocationName as Domicle"
        SQLz = SQLz & ",LState as State"
        SQLz = SQLz & ",UnitUnitNum as [Unit Num]"
        SQLz = SQLz & ",UnitVIN as VIN"
        SQLz = SQLz & ",[Desc] as [Description]"
        SQLz = SQLz & ",UnitStatus as Status"
        SQLz = SQLz & ",CurrRent as Rent"
        SQLz = SQLz & ",iif(UnitTaxExempt=True,""N"",""Y"") as [Taxed?]"
        SQLz = SQLz & ",LTaxRate as [Tax Rate]"
        SQLz = SQLz & ",TaxAmt as Tax"
        SQLz = SQLz & ",FSLFee as [Serv Fee]"
        SQLz = SQLz & ",LineTot as [Total Payment] "
        SQLz = SQLz & " From vw_BankBill "
        SQLz = SQLz & " WHERE BillID=" & Me.BillID & " "
        If Me.BillType = "Rent Group" Then SQLz = SQLz & "AND SGrpID = " & [Forms]![Billing]![GroupID] & " "
   '----Default XLS
        XLS = "|K,TC|O,BC|P,$E|R,P3|Q,BC|R,TC|S,$E|T,$E|U,$E|T,W09|"

        If Me.BillType Like "PDSurr*" Then
            SQLz = " SELECT MLNo"                       'A
            SQLz = SQLz & ",SchNo"
            SQLz = SQLz & ",SGrpNO"
            SQLz = SQLz & ",'FA' as Bank"               'D
            SQLz = SQLz & ",BLCD"
            SQLz = SQLz & ",UnitId As AssetID"
            SQLz = SQLz & ",UnitUnitNum as UnitNum"
            SQLz = SQLz & ",UnitVIN"                    'H
            SQLz = SQLz & ",UnitStatus as Sts"
            SQLz = SQLz & ",UnitDescShortVer as Description"  'J
            SQLz = SQLz & ",LocationName as Domicle"
            SQLz = SQLz & ",UnitPDStartDate as [PD Start]"
            SQLz = SQLz & ",PDBegin As BillPDStart"     'M
            SQLz = SQLz & ",PDEnd As BillPDEnd"
            SQLz = SQLz & ",UnitTaxExempt as Exempt"
            SQLz = SQLz & ",LTaxRate as TaxRate"
            SQLz = SQLz & ",PDRent" ' P
            SQLz = SQLz & ",PDDays"
            SQLz = SQLz & ",BasePerDiem   as BillPDRent"
            SQLz = SQLz & ",ServiceFee  as ServFee"
            SQLz = SQLz & ",Taxes"
            SQLz = SQLz & ",LineTotal"
            SQLz = SQLz & ",OffLeaseActual"
            SQLz = SQLz & ",InspectionPassed "
            SQLz = SQLz & "FROM vw_PerdiemSurrenderInvoice WHERE BillID = " & Me.BillID & ";"
            XLS = "|I,BC|O,BC|P,P3|Q,$E|R,PC|S,$E|T,$E|U,$E|V,$E|"
        End If
        If Me.BillType = "PerDiem" And Me.GroupBill Like "*Group*" Then
            SQLz = "SELECT vw_PerdiemInvoice.SGrpID, * "
            SQLz = SQLz & "FROM vw_PerdiemInvoice "
            SQLz = SQLz & "WHERE (((vw_PerdiemInvoice.SGrpID)= " & [Forms]![Billing]![GroupID] & ") AND ((vw_PerdiemInvoice.BillID)= " & [Forms]![Billing]![BillID] & ")) ;"
        End If
        If Me.BillType = "PerDiem" Then
            SQLz = "Select BillNo AS [Invoice Number]"              '-Col A
            SQLz = SQLz & ",'Fleet Advantage, LLC' AS [Remit To]"   '-Col B
            SQLz = SQLz & ",'401 E. Las Olas Blvd Sutie 1720 Fort Lauderdale, Fl. 33301' AS [Remit Address]"
            SQLz = SQLz & ",MLNo"                                   '-Col D
            SQLz = SQLz & ",SchNo AS Sch"                           '-Col E
            SQLz = SQLz & ",SGrpNO aS Grp"                          ' Col F
            SQLz = SQLz & ",BLCD"                                   '-Col G
            SQLz = SQLz & ",UnitUnitNum AS [Unit Num]"              '-Col H
            SQLz = SQLz & ",UnitVIN AS [Unit VIN]"                  '-Col I
            SQLz = SQLz & ",UnitStatus AS [Unit Status]"            '-Col J
            SQLz = SQLz & ",UnitDescShortVer as [Description]"      '-Col K
            SQLz = SQLz & ",LocationName As Domicle"                '-Col L
            SQLz = SQLz & ",LCity"                                  '-Col M
            SQLz = SQLz & ",LState"                                 '-Col N
            SQLz = SQLz & ",Billdt AS [Invoice Date]"               '-Col O
            SQLz = SQLz & ",BillDueDt As [Due Date]"                '-Col P
            SQLz = SQLz & ",Billfrom AS [Start Date]"               '-Col Q
            SQLz = SQLz & ",Billto AS [End Date]"                   '-Col R
            SQLz = SQLz & ",PDBegin AS [Interim Begin]"             '-Col S
            SQLz = SQLz & ",BilledPDEndDate AS [Interim End]"  '     -Col T
            SQLz = SQLz & ",PDDays AS [Interim Days]"               '-Col U
            SQLz = SQLz & ",BasePerDiem AS [Base Interim Rent]"     '-Col V
            SQLz = SQLz & ",PDTax AS [Total Tax]"                   '-Col W
            SQLz = SQLz & ",LineTotal AS [Total Interim Rent]"      '-Col X
            SQLz = SQLz & " FROM vw_PerdiemInvoice"
            SQLz = SQLz & " WHERE BillID=" & [Forms]![Billing]![BillID] & ";"
            XLS = "|A,DD3|A,Wrap|F,TC|I,BC|J,TC|N,TC|U,BC|V,$E|W,$E|X,$E"
            XLS = XLS & "|A,W12|B,W19|F,W09|J,W08|O,W11|P,W11|Q,W11|R,W11|S,W11|T,W12|U,W9|V,W11|W,W09|X,W12|"
        End If

        Call MessageUser("Open", "Getting " & Me.BillType & " Report", "One Moment...")
        SendReport2Excel SQLz, Me.BillType, "Invoice: " & Me.BillNo & " | " & Me.Billdt & "| " & HeaderTypex, True, XLS, 300
        Call MessageUser("Close")
    Else
        DoCmd.OpenReport ReportX, acViewPreview, , , , OpenArgsX
    End If
    Exit Function
FixMe:
    MsgBox Err.Description, vbCritical, "(Error!)" & AtlaasVersion
End Function
