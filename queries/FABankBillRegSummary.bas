Operation =1
Option =0
Begin InputTables
    Name ="FABankBillReg"
End
Begin OutputColumns
    Expression ="FABankBillReg.LocationName"
    Expression ="FABankBillReg.SchNo"
    Alias ="GroupNo"
    Expression ="FABankBillReg.SGrpNO"
    Expression ="FABankBillReg.LState"
    Expression ="FABankBillReg.UnitTaxExempt"
    Expression ="FABankBillReg.StTaxType"
    Expression ="FABankBillReg.LTaxRate"
    Alias ="UnitCount"
    Expression ="Count(FABankBillReg.AssetID)"
    Alias ="LineTot"
    Expression ="Sum(FABankBillReg.LineTot)"
    Alias ="FSLFee"
    Expression ="Sum(FABankBillReg.FSLFee)"
    Alias ="RentNoTax"
    Expression ="Sum(FABankBillReg.RentNoTax)"
    Alias ="TaxAmt"
    Expression ="Sum(FABankBillReg.TaxAmt)"
    Alias ="UnitTotalRent"
    Expression ="Sum(FABankBillReg.UnitTotalRent)"
    Alias ="FSLTax"
    Expression ="Sum(0)"
    Expression ="FABankBillReg.BillMonths"
End
Begin Groups
    Expression ="FABankBillReg.LocationName"
    GroupLevel =0
    Expression ="FABankBillReg.SchNo"
    GroupLevel =0
    Expression ="FABankBillReg.SGrpNO"
    GroupLevel =0
    Expression ="FABankBillReg.LState"
    GroupLevel =0
    Expression ="FABankBillReg.UnitTaxExempt"
    GroupLevel =0
    Expression ="FABankBillReg.StTaxType"
    GroupLevel =0
    Expression ="FABankBillReg.LTaxRate"
    GroupLevel =0
    Expression ="FABankBillReg.BillMonths"
    GroupLevel =0
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
        dbText "Name" ="UnitCount"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="LineTot"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLFee"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLTax"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalRent"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentNoTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupNo"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillReg.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillReg.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillReg.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillReg.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillReg.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillReg.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfAssetID"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillReg.BillMonths"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-21
    Top =31
    Right =1548
    Bottom =846
    Left =-1
    Top =-1
    Right =1545
    Bottom =345
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="FABankBillReg"
        Name =""
    End
End
