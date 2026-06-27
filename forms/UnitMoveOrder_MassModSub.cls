Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz, SQLset As String
Public Committ As String
Public BillIdx As String
Dim RST As DAO.Recordset

Public Sub BillBankIDSubmit_Click()
    Call ConfrimIT("BillBankCBO", "Bank")
    SQLz = "UPDATE BillID Set BillBankID=" & Me.BillBankIDCBO & " Where BillID In(BillIDx);"
    If Committ = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Public Sub BilldtSubmit_Click()
    Call ConfrimIT("BilldtCBO", "Bill Date")
    SQLz = "UPDATE Bills Set Billdt=#" & Me.BilldtCBO & "# Where BillID In(BillIDx);"
    If Committ = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Public Sub BillDueDtSubmit_Click()
    Call ConfrimIT("BillDueDtCBO", "Bill Due Date")
    SQLz = "UPDATE Bills Set BillDueDt=#" & Me.BillDueDtCbo & "# Where BillID In(BillIDx);"
    If Committ = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Public Sub BillFromSubmit_Click()
    Call ConfrimIT("BillFromCBO", "Bill From")
    SQLz = "UPDATE Bills Set BillFrom=#" & Me.BillFromCbo & "# Where BillID In(BillIDx);"
    If Committ = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Public Sub BillIDentSubmit_Click()
    Call ConfrimIT("BillIDentCBO", "Bill Sent Date")
    SQLz = "UPDATE Bills Set BillIDent=#" & Me.BillIDentCBO & "# Where BillID In(BillIDx);"
    If Committ = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub BillIDent_AfterUpdate()
    Call CheckSoldStatus
End Sub

Private Sub BillID_DblClick(Cancel As Integer)
  '  Call OpenBillsForm(Me.BillID)
    'DoCmd.OpenForm "BillsForm", , , "BillID=" & Me.BillID & ""
 '   Forms!Billsform.BillsTabControl.Value = 6
  '  Me.UnitNum.SetFocus
End Sub

Public Sub BilltoSubmit_Click()
    Call ConfrimIT("BilltoCBO", "Bill To Date")
    SQLz = "UPDATE Bills Set Billto=#" & Me.BilltoCBO & "# Where BillID In(BillIDx);"
    If Committ = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub BillTypeID_AfterUpdate()
Me.BillType = Me.BillTypeID.Column(1)
End Sub

Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Public Sub ConfrimIT(ByVal FieldBoX As String, ByVal TragetX As String, Optional FieldValue As String)
    Dim Value2Display
    If FieldValue <> "" Then Value2Display = FieldValue Else Value2Display = Me(FieldBoX).Value
    If Value2Display & "TEST" = "TEST" Then
        Committ = MsgBox("You're about set the " & TragetX & " to Null, for ALL Bills listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Update-Change confrim " & TragetX & "?")
        'If Committ = vbNo Then Exit Sub
    Else
        Committ = MsgBox("You're about set the " & TragetX & " to " & Value2Display & ", for ALL Bills listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & TragetX & "?")
        'If Committ = vbNo Then Exit Sub
    End If
End Sub

Private Sub DeleteRecord_Click()
    On Error GoTo FixMe
    Dim Schex, Typex As String
    Schex = DLookup("Schedule", "vw_FiveKeys", "SchID=" & DLookup("BillSchRef", "Bills", "BillId=" & CLng(Me.txtCurrent)) & "")
    Typex = DLookup("BillType", "Bills", "BillId=" & CLng(Me.txtCurrent) & "")
    Committ = MsgBox("You're about Delete a " & Typex & " from Sch: " & Schex & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change ")
    If Committ = vbNo Then Exit Sub
    DoCmd.RunSQL ("Delete From Bills WHERE BillID = " & Me.BillID & ";")
    Me.Form.Requery
FixMe:
End Sub

Private Sub DueonRecSubmit_Click()
    Call ConfrimIT("DueonRecCBO", "Due On Receipt")
    SQLz = "UPDATE Bills Set DueonRec=" & Me.DueonRecCbo & " Where BillID in(" & SplitString(SplitString(CStr(Me.Form.Filter), "(", 1), ")", 0) & ");"
    If Committ = vbNo Then Exit Sub
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    Me.Form.Refresh
    Exit Sub
End Sub

Private Sub Form_Current()
    On Error Resume Next
    Me.txtCurrent = Me.BillID
    On Error GoTo 0
End Sub

Private Sub InvoicePaidDate_AfterUpdate()
    Call CheckSoldStatus
End Sub

Private Sub InvoiceSentDate_AfterUpdate()
    Call CheckSoldStatus
End Sub

Private Sub Form_Load()
    SQLz = "SELECT BillID, BillBankID, BillSchRef, GroupId, BillType, BillTypeID, BillNo, Billdt, Billfrom"
    SQLz = SQLz & ", Billto, BillDueDt, BillSent, BillLink, DueonRec, SchNo, SGrpNO, AssigneeBankID"
    SQLz = SQLz & ", iif(nz(SchBnkClNO,'')='', '',SchBnkClNO & '-') & SgrpBnkNo as BContract "
    SQLz = SQLz & " FROM (((Bills"
    SQLz = SQLz & " INNER JOIN Schedule ON Bills.BillSchRef = Schedule.SchID)"
    SQLz = SQLz & " INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID)"
    SQLz = SQLz & " LEFT JOIN SchGrp ON Bills.GroupId = SchGrp.SGrpID)"
    SQLz = SQLz & " WHERE BillId =0"
    SQLz = SQLz & " ORDER BY Billdt DESC"
    Me.Form.RecordSource = SQLz
    Me.txtCurrent = Null
End Sub

Private Sub InvoiceToExcel_Click()
    Call MakeInvoice(1)
End Sub
Private Sub Runbill_Click()
    Call MakeInvoice(0)
End Sub
Public Function MakeInvoice(ToExcel As Boolean)
    If Me.Dirty Then
        Me.Dirty = False
        Me.Form.Requery
    End If
    If Me.txtCurrent & "test" = "test" Then Exit Function
    Set RST = CurrentDb.OpenRecordset("SELECT * FROM Bills WHERE BillId=" & CLng(Me.txtCurrent) & ";", dbOpenSnapshot)
    Dim FundStatus As String

    FundStatus = DMax("Assigned_Indicator", "SchGrp", "sGrpSchID=" & RST!BillSchRef)
    Dim ReportX, ReportX2, OpenArgsX As String
    Dim RecX As Integer
    If RST!BillType = "BankBill" Then ReportX = "FABankBillGroup_Batch"
    If RST!BillType = "BankFA Remit" Then ReportX = "FABankBillGroup_Batch"
    If RST!BillType = "PDSurrender" Then ReportX = "PerDiemSurrender_Batch"
    If RST!BillType = "PDSurrLOC" Then ReportX = "PerDiemSurrender_Batch"
    If RST!BillType = "PerDiem" Then ReportX = "PerDiemInvoice_Batch"
    If RST!BillType = "Rent" Then ReportX = "FABankBillGroup_Batch"
    If RST!BillType = "Rent Group" Then ReportX = "FABankBillGroup_Batch"
    If RST!BillType = "BankBill" And Me.SGrpNO & "test" <> "test" Then ReportX = "FABankBillGroup_Batch"

    'Run The Bill -----------------------------------------
    DoCmd.close acReport, ReportX
    On Error GoTo FixMe
    If ToExcel = True Then
        Dim XLS As String
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
       SQLz = SQLz & " WHERE BillID=" & RST!BillID & ";"

       XLS = "|K,TC|O,BC|P,$E|R,P3|Q,BC|R,TC|S,$E|T,$E|U,$E|V,$E"

        If RST!BillType Like "PDSurr*" Then
            SQLz = " SELECT MLNo"                       'A
            SQLz = SQLz & ",SchNo"
            SQLz = SQLz & ",SGrpNO"
            SQLz = SQLz & ",'FA' as Bank"               'D
            SQLz = SQLz & ",BLCD"
            SQLz = SQLz & ",UnitId As AssetID"
            SQLz = SQLz & ",UnitUnitNum as UnitNum"
            SQLz = SQLz & ",UnitVIN"                    'H
            SQLz = SQLz & ",UnitStatus as Sts"
            SQLz = SQLz & ",UnitDescShortVer as Descp"  'J
            SQLz = SQLz & ",LocationName as Location"
            SQLz = SQLz & ",UnitPDStartDate as UnitPDStart"
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
        If RST!BillType = "PerDiem" And RST!groupID & "test" <> "test" Then
        SQLz = "SELECT vw_PerdiemInvoice.SGrpID, * "
        SQLz = SQLz & "FROM vw_PerdiemInvoice "
        SQLz = SQLz & "WHERE (((vw_PerdiemInvoice.SGrpID)= " & RST!groupID & ") AND ((vw_PerdiemInvoice.BillID)= " & [Forms]![Billing]![BillID] & ")) ;"
        End If
        If RST!BillType = "PerDiem" Then
           SQLz = "SELECT * FROM vw_PerdiemInvoice WHERE BillID=" & RST!BillID & ";"
           XLS = "|A,W20|C,$E|D,$E|E,$E|F,$E|G,$E|H,$E|I,$E|V,BC|R,HHY,HH|BB,BC|"
        End If
        If RST!BillType = "Rent Group" Then SQLz = "Select * From FABankBillGroup"
        'If RST!BillType = "Rent" Then SQLz = "Select * From FABankBillReg"
        SendReport2Excel SQLz, RST!BillType, "Invoice: " & RST!BillNo & " | " & RST!Billdt & "| " & BillType, True, XLS, 0
    Else
    'make Bill
        Call MAKE_GroupBillQuery("SELECT * FROM Bills WHERE BillID =" & RST!BillID & ";", False)
        DoCmd.OpenReport ReportX, acViewPreview
    'make perdiem
     '   DoCmd.OpenReport ReportX, acViewPreview, , , , OpenArgsX
    End If
   ' RST.close
    Set RST = Nothing
    Exit Function
FixMe:
    MsgBox Err.Description, vbCritical, "(Error!)" & AtlaasVersion

End Function
Public Function MAKE_GroupBillQuery(Optional SQLx As String, Optional NoQuery As Boolean, Optional BillType As String)
    Dim qdf As DAO.QueryDef
    Dim i As Integer
    On Error Resume Next
    If BillType = "" Then BillType = "BankBill"
    If SQLx = "" Then
        Set RST = CurrentDb.OpenRecordset("SELECT * FROM Bills_Batch_WorkTable")
    Else
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("Delete From Bills_Batch_WorkTable")
        DoCmd.SetWarnings True
        Set RST = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
    End If
    RST.MoveLast
    RST.MoveFirst
    BillIdx = ""
    For i = 1 To RST.RecordCount
        BillIdx = BillIdx & RST!BillID
        BillFromx = RST!Billfrom
        BillToX = RST!Billto
        If SQLx <> "" Then DoCmd.RunSQL ("Insert into Bills_Batch_WorkTable (BillID) Select" & RST!BillID & ";")
        If RST.RecordCount > 1 And i < RST.RecordCount Then BillIdx = BillIdx & ","
        RST.MoveNext
    Next i
    RST.close
    Set RST = Nothing
    On Error GoTo 0
    If NoQuery = True Then Exit Function
    If BillType = "BankBill" Then
        SQLz = "SELECT vw_BankBill_Batch.*"
        SQLz = SQLz & ", CurrRent AS RentNoTax"
        SQLz = SQLz & ", IIf([UnitTaxExempt]=False,CCur([FSLFee]*[LTaxRate]),0) AS FSLTax"
        SQLz = SQLz & ", [CurrRent]+CCur(Nz([FSLFee],0))+CCur(Nz([TaxAmt],0)) AS UnitTotalRent"
        SQLz = SQLz & ", BillToGroup AS BillToAddress, BillNo AS InvoiceNumber"
        SQLz = SQLz & " From vw_BankBill_Batch"
        SQLz = SQLz & " WHERE BillID In (" & BillIdx & ");"
        On Error Resume Next
        DoCmd.DeleteObject acQuery, "FABankBillGroupBatch"
        On Error GoTo 0
        Set qdf = CurrentDb.CreateQueryDef("FABankBillGroupBatch", SQLz)
    End If
    If BillType = "PerDiem" Then
        SQLz = "SELECT vw_PerDiem_Batch.*"
      '  Sqlz = Sqlz & ", CurrRent AS RentNoTax"
      '  Sqlz = Sqlz & ", IIf([UnitTaxExempt]=False,CCur([FSLFee]*[LTaxRate]),0) AS FSLTax"
      '  Sqlz = Sqlz & ", [CurrRent]+CCur(Nz([FSLFee],0))+CCur(Nz([TaxAmt],0)) AS UnitTotalRent"
      '  Sqlz = Sqlz & ", BillToGroup AS BillToAddress, BillNo AS InvoiceNumber"
      '  Sqlz = Sqlz & " From vw_BankBill_Batch"
        SQLz = SQLz & " WHERE BillID In (" & BillIdx & ");"
        On Error Resume Next
        DoCmd.DeleteObject acQuery, "FAPerDiemBatch"
        On Error GoTo 0
        Set qdf = CurrentDb.CreateQueryDef("FAPerDiemBatch", SQLz)
    End If
End Function
