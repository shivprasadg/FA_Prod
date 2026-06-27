dbMemo "SQL" ="SELECT vw_BankBill_Batch.*, CurrRent AS RentNoTax, IIf([UnitTaxExempt]=False,CCu"
    "r([FSLFee]*[LTaxRate]),0) AS FSLTax, [CurrRent]+CCur(Nz([FSLFee],0))+CCur(Nz([Ta"
    "xAmt],0)) AS UnitTotalRent, BillToGroup AS BillToAddress, BillNo AS InvoiceNumbe"
    "r\015\012FROM vw_BankBill_Batch\015\012WHERE BillID In (36718);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
