Attribute VB_Name = "BillingReports"
Option Compare Database
Dim SQLz As String

Public Function ExportBilling(StartDatex As String, EndDatex As String, BillTypeX As String, Optional ClientGroupIDx As Long)
    'On Error GoTo Fixme
    Dim Headerx As String
    If BillTypeX = "Bank" Then
        Dim XLS As String
        SQLz = "Select Client "
        SQLz = SQLz & ",[Invoice Number]"
        SQLz = SQLz & ",[Assignee Contract #]"
        SQLz = SQLz & ",[Remit To]"
        SQLz = SQLz & ",[Remit Address]"
        SQLz = SQLz & ",[Lease]"
        SQLz = SQLz & ",[Schedule No]"
        SQLz = SQLz & ",[Invoice Date]"
        SQLz = SQLz & ",[Due Date]"
        SQLz = SQLz & ",[Billing Period]"
        SQLz = SQLz & ",Location"
        SQLz = SQLz & ",State"
        SQLz = SQLz & ",[Unit Num]"
        SQLz = SQLz & ",VIN"
        SQLz = SQLz & ",Description"
        SQLz = SQLz & ",[Status]"
        SQLz = SQLz & ",Rent"
        SQLz = SQLz & ",iif(UnitTaxExempt=True,""N"",""Y"") as [Taxed?]"
        SQLz = SQLz & ",[Tax Rate]"
        SQLz = SQLz & ",Tax"
        SQLz = SQLz & ",[Serv Fee]"
        SQLz = SQLz & ",[Total Payment] "
        SQLz = SQLz & ",BillType"
        SQLz = SQLz & ",CurrRentType "
        SQLz = SQLz & " From vw_BankBillReport "
        SQLz = SQLz & " WHERE Billfrom >= #" & StartDatex & "# AND BillTo <= #" & EndDatex & "# "
        If ClientGroupIDx > 0 Then
        SQLz = SQLz & " AND ClientGroupId= " & ClientGroupIDx & " "
        End If
        SQLz = SQLz & " Order By VIN, [Invoice Date] Desc, [Invoice Number], Client"
        XLS = "|L,TC|P,BC|Q,$E|R,BC|S,P3|T,$E|U,$E|V,$E|W,$E"
    End If
    If BillTypeX = "PDSurr*" Then
        SQLz = " SELECT MLNo"
        SQLz = SQLz & ",SchNo"
        SQLz = SQLz & ",SGrpNO"
        SQLz = SQLz & ",'FA' as Bank"
        SQLz = SQLz & ",BLCD"
        SQLz = SQLz & ",UnitId As AssetID"
        SQLz = SQLz & ",UnitUnitNum as UnitNum"
        SQLz = SQLz & ",UnitVIN"
        SQLz = SQLz & ",UnitStatus as Sts"
        SQLz = SQLz & ",UnitDescShortVer as Descp"
        SQLz = SQLz & ",UnitPDStartDate as UnitPDStart"
        SQLz = SQLz & ",PDBegin As BillPDStart"
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
        SQLz = SQLz & "FROM vw_PerdiemSurrenderInvoice "
        SQLz = SQLz & " WHERE [Billfrom] <=#" & StartDatex & "# AND [Billfrom] >= #" & EndDatex & "# ;"
        XLS = "|I,BC|N,BC|O,P3|P,$E|Q,PC|R,$E|S,$E|T,$E|U,$E|"
    End If
   Headerx = "Billing Invoices From: " & StartDatex & " ~ " & EndDatex
    If ClientGroupIDx > 0 Then Headerx = Headerx & " by Client"
    Call MessageUser("Open", "Making Billing Report", "Excel will Open")
    SendReport2Excel SQLz, "Billing", Headerx, True, XLS, 0
    Call MessageUser("Close")
    Exit Function
FixMe:
    MsgBox Err.Description, vbCritical, "(Error!)" & AtlaasVersion
End Function
