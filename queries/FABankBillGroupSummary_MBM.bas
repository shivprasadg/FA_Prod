Operation =1
Option =0
Begin InputTables
    Name ="FABankBillGroup_MBM"
End
Begin OutputColumns
    Expression ="FABankBillGroup_MBM.LocationName"
    Expression ="FABankBillGroup_MBM.SchNo"
    Alias ="GroupNo"
    Expression ="FABankBillGroup_MBM.SGrpNO"
    Expression ="FABankBillGroup_MBM.LState"
    Expression ="FABankBillGroup_MBM.UnitTaxExempt"
    Expression ="FABankBillGroup_MBM.StTaxType"
    Expression ="FABankBillGroup_MBM.LTaxRate"
    Alias ="UnitCount"
    Expression ="Sum(FABankBillGroup_MBM.UnitCount)"
    Alias ="LineTot"
    Expression ="Sum(FABankBillGroup_MBM.LineTot)"
    Alias ="FSLFee"
    Expression ="Sum(FABankBillGroup_MBM.FSLFee)"
    Alias ="RentNoTax"
    Expression ="Sum(FABankBillGroup_MBM.RentNoTax)"
    Alias ="TaxAmt"
    Expression ="Sum(FABankBillGroup_MBM.TaxAmt)"
    Alias ="UnitTotalRent"
    Expression ="Sum(FABankBillGroup_MBM.UnitTotalRent)"
    Alias ="FSLTax"
    Expression ="Sum(0)"
End
Begin Groups
    Expression ="FABankBillGroup_MBM.LocationName"
    GroupLevel =0
    Expression ="FABankBillGroup_MBM.SchNo"
    GroupLevel =0
    Expression ="FABankBillGroup_MBM.SGrpNO"
    GroupLevel =0
    Expression ="FABankBillGroup_MBM.LState"
    GroupLevel =0
    Expression ="FABankBillGroup_MBM.UnitTaxExempt"
    GroupLevel =0
    Expression ="FABankBillGroup_MBM.StTaxType"
    GroupLevel =0
    Expression ="FABankBillGroup_MBM.LTaxRate"
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
        dbText "Name" ="FABankBillGroup_MBM.LocationName"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_MBM.SchNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_MBM.stateRef"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_MBM.UnitTaxExempt"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_MBM.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_MBM.LocTaxRt"
        dbInteger "ColumnOrder" ="6"
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
    Right =1537
    Bottom =379
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =243
        Top =92
        Right =387
        Bottom =236
        Top =0
        Name ="FABankBillGroup_MBM"
        Name =""
    End
End
