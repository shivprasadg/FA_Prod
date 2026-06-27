Operation =1
Option =0
Begin InputTables
    Name ="FABankBillGroup"
End
Begin OutputColumns
    Expression ="FABankBillGroup.LocationName"
    Expression ="FABankBillGroup.SchNo"
    Alias ="GroupNo"
    Expression ="FABankBillGroup.SGrpNO"
    Expression ="FABankBillGroup.LState"
    Expression ="FABankBillGroup.UnitTaxExempt"
    Expression ="FABankBillGroup.StTaxType"
    Expression ="FABankBillGroup.LTaxRate"
    Alias ="UnitCount"
    Expression ="Sum(FABankBillGroup.UnitCount)"
    Alias ="LineTot"
    Expression ="Sum(FABankBillGroup.LineTot)"
    Alias ="FSLFee"
    Expression ="Sum(FABankBillGroup.FSLFee)"
    Alias ="RentNoTax"
    Expression ="Sum(FABankBillGroup.RentNoTax)"
    Alias ="TaxAmt"
    Expression ="Sum(FABankBillGroup.TaxAmt)"
    Alias ="UnitTotalRent"
    Expression ="Sum(FABankBillGroup.UnitTotalRent)"
    Alias ="FSLTax"
    Expression ="Sum(0)"
    Expression ="FABankBillGroup.CurrRent"
End
Begin Groups
    Expression ="FABankBillGroup.LocationName"
    GroupLevel =0
    Expression ="FABankBillGroup.SchNo"
    GroupLevel =0
    Expression ="FABankBillGroup.SGrpNO"
    GroupLevel =0
    Expression ="FABankBillGroup.LState"
    GroupLevel =0
    Expression ="FABankBillGroup.UnitTaxExempt"
    GroupLevel =0
    Expression ="FABankBillGroup.StTaxType"
    GroupLevel =0
    Expression ="FABankBillGroup.LTaxRate"
    GroupLevel =0
    Expression ="FABankBillGroup.CurrRent"
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
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="FABankBillGroup.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup.CurrRent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =80
    Top =57
    Right =1649
    Bottom =872
    Left =-1
    Top =-1
    Right =1545
    Bottom =328
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =194
        Bottom =292
        Top =0
        Name ="FABankBillGroup"
        Name =""
    End
End
