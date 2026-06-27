Operation =1
Option =0
Begin InputTables
    Name ="FABankBillGroupBatch"
End
Begin OutputColumns
    Expression ="FABankBillGroupBatch.LocationName"
    Expression ="FABankBillGroupBatch.LState"
    Alias ="UnitCount"
    Expression ="Sum(FABankBillGroupBatch.UnitCount)"
    Alias ="LineTot"
    Expression ="Sum(FABankBillGroupBatch.LineTot)"
    Alias ="FSLFee"
    Expression ="Sum(FABankBillGroupBatch.FSLFee)"
    Alias ="RentNoTax"
    Expression ="Sum(FABankBillGroupBatch.RentNoTax)"
    Alias ="TaxAmt"
    Expression ="Sum(FABankBillGroupBatch.TaxAmt)"
    Alias ="UnitTotalRent"
    Expression ="Sum(FABankBillGroupBatch.UnitTotalRent)"
    Alias ="FSLTax"
    Expression ="0"
End
Begin Groups
    Expression ="FABankBillGroupBatch.LocationName"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.LState"
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
        dbText "Name" ="FABankBillGroupBatch.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.CurrRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfCurrRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrRent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-1
    Top =470
    Right =1568
    Bottom =1285
    Left =-1
    Top =-1
    Right =1545
    Bottom =260
    Left =240
    Top =0
    ColumnsShown =543
    Begin
        Left =-164
        Top =37
        Right =-20
        Bottom =181
        Top =0
        Name ="FABankBillGroupBatch"
        Name =""
    End
End
