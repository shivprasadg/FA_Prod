Operation =1
Option =0
Where ="(((FABankBill.SGrpID)=[Forms]![Billing]![GroupId]))"
Begin InputTables
    Name ="FABankBill"
End
Begin OutputColumns
    Expression ="FABankBill.*"
    Alias ="RentNoTax"
    Expression ="FABankBill.CurrRent"
    Alias ="FSLTax"
    Expression ="IIf([UnitTaxExempt]=False,CCur([FSLFee]*[LTaxRate]),0)"
    Alias ="UnitTotalRent"
    Expression ="[CurrRent]+CCur(Nz([FSLFee],0))+CCur(Nz([TaxAmt],0))"
    Alias ="BillToAddress"
    Expression ="FABankBill.BillToGroup"
    Alias ="InvoiceNumber"
    Expression ="FABankBill.BillNo"
    Expression ="FABankBill.SGrpID"
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
        dbInteger "ColumnWidth" ="4500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UNITSRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UNITPDRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BankContract"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.EndOrActualDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UnitsBillRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BillToSchedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BLoST3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BillNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BLocStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BillToGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.FAPOnumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.DateFeesEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.Billfrom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.Billto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BankName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.SgrpBnkNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LineTot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.FSLFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.ShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.TaxAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.RemitAssignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.SGrpFABill"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.SchBnkClNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.FAVendorNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.Vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.CurrRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.CurrRentType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.RemitAddr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.UCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BillAdj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BillMonths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.ExtRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BillBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.vw_BankBill.BillType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBill.SGrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =41
    Top =105
    Right =2089
    Bottom =940
    Left =-1
    Top =-1
    Right =2024
    Bottom =416
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =89
        Top =43
        Right =336
        Bottom =504
        Top =0
        Name ="FABankBill"
        Name =""
    End
End
