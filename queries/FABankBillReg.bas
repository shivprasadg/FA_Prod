Operation =1
Option =0
Where ="(((FABankBill.BillID)=Forms!Billing!BillID) And ((FABankBill.LocationID) Is Not "
    "Null) And ((FABankBill.SGrpID) Is Not Null))"
Begin InputTables
    Name ="vw_BankBill"
    Alias ="FABankBill"
End
Begin OutputColumns
    Expression ="FABankBill.*"
    Alias ="RentNoTax"
    Expression ="FABankBill.CurrRent"
    Alias ="FSLTax"
    Expression ="0"
    Alias ="UnitTotalRent"
    Expression ="[RentNoTax]+CCur(Nz([FSLFee],0))+CCur(Nz([TaxAmt],0))"
    Alias ="BillToAddress"
    Expression ="FABankBill.BillToSchedule"
    Alias ="InvoiceNumber"
    Expression ="FABankBill.BillNo"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="RentNoTax"
        dbInteger "ColumnOrder" ="66"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalRent"
        dbInteger "ColumnWidth" ="2205"
        dbInteger "ColumnOrder" ="67"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLTax"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillToAddress"
        dbInteger "ColumnWidth" ="13350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.FSLFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SchBnkClNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BLocStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BillNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BillBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Billfrom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Billto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BillToGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BillToSchedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BankName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UnitVin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UNITSRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.TaxAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BLoST3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SGrpFABill"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.EndOrActualDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SgrpBnkNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.ShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.RemitAssignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BankContract"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.DateFeesEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UnitsBillRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LineTot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.FAVendorNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.FAPOnumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UNITPDRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.Vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.CurrRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.CurrRentType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.UCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.RemitAddr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.BillAdj"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-1050
    Top =319
    Right =998
    Bottom =1151
    Left =-1
    Top =-1
    Right =2024
    Bottom =342
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =36
        Top =23
        Right =221
        Bottom =408
        Top =0
        Name ="FABankBill"
        Name =""
    End
End
