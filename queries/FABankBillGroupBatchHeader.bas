Operation =1
Option =0
Begin InputTables
    Name ="FABankBillGroupBatch"
End
Begin OutputColumns
    Expression ="FABankBillGroupBatch.BillID"
    Expression ="FABankBillGroupBatch.BillNo"
    Expression ="FABankBillGroupBatch.BankName"
    Expression ="FABankBillGroupBatch.BillGroupID"
    Expression ="FABankBillGroupBatch.BankContract"
    Expression ="FABankBillGroupBatch.Billfrom"
    Expression ="FABankBillGroupBatch.Billto"
    Expression ="FABankBillGroupBatch.SchNo"
    Alias ="GroupNo"
    Expression ="FABankBillGroupBatch.SGrpNO"
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
    Expression ="Sum(0)"
    Alias ="SumOfCurrRent"
    Expression ="Sum(FABankBillGroupBatch.CurrRent)"
    Expression ="FABankBillGroupBatch.BillToSchedule"
    Expression ="FABankBillGroupBatch.RemitAssignee"
    Expression ="FABankBillGroupBatch.aSort"
End
Begin Groups
    Expression ="FABankBillGroupBatch.BillID"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.BillNo"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.BankName"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.BillGroupID"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.BankContract"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.Billfrom"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.Billto"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.SchNo"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.SGrpNO"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.BillToSchedule"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.RemitAssignee"
    GroupLevel =0
    Expression ="FABankBillGroupBatch.aSort"
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
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLFee"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLTax"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalRent"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentNoTax"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="TaxAmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="GroupNo"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.BillID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.BillNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfCurrRent"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.BankContract"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.BankName"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.BillToSchedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.Billto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.RemitAssignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.BillGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.Billfrom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroupBatch.aSort"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =199
    Top =165
    Right =1768
    Bottom =980
    Left =-1
    Top =-1
    Right =1545
    Bottom =324
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =76
        Top =37
        Right =220
        Bottom =226
        Top =0
        Name ="FABankBillGroupBatch"
        Name =""
    End
End
