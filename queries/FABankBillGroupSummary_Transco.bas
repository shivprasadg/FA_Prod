Operation =1
Option =0
Begin InputTables
    Name ="FABankBillGroup_Transco"
End
Begin OutputColumns
    Expression ="FABankBillGroup_Transco.LocationName"
    Expression ="FABankBillGroup_Transco.SchNo"
    Alias ="GroupNo"
    Expression ="FABankBillGroup_Transco.SGrpNO"
    Expression ="FABankBillGroup_Transco.LState"
    Expression ="FABankBillGroup_Transco.UnitTaxExempt"
    Expression ="FABankBillGroup_Transco.StTaxType"
    Expression ="FABankBillGroup_Transco.LTaxRate"
    Alias ="UnitCount"
    Expression ="Sum(FABankBillGroup_Transco.UnitCount)"
    Alias ="LineTot"
    Expression ="Sum(FABankBillGroup_Transco.LineTot)"
    Alias ="FSLFee"
    Expression ="Sum(FABankBillGroup_Transco.FSLFee)"
    Alias ="RentNoTax"
    Expression ="Sum(FABankBillGroup_Transco.RentNoTax)"
    Alias ="TaxAmt"
    Expression ="Sum(FABankBillGroup_Transco.TaxAmt)"
    Alias ="UnitTotalRent"
    Expression ="Sum(FABankBillGroup_Transco.UnitTotalRent)"
    Alias ="FSLTax"
    Expression ="Sum(0)"
End
Begin Groups
    Expression ="FABankBillGroup_Transco.LocationName"
    GroupLevel =0
    Expression ="FABankBillGroup_Transco.SchNo"
    GroupLevel =0
    Expression ="FABankBillGroup_Transco.SGrpNO"
    GroupLevel =0
    Expression ="FABankBillGroup_Transco.LState"
    GroupLevel =0
    Expression ="FABankBillGroup_Transco.UnitTaxExempt"
    GroupLevel =0
    Expression ="FABankBillGroup_Transco.StTaxType"
    GroupLevel =0
    Expression ="FABankBillGroup_Transco.LTaxRate"
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
        dbText "Name" ="FABankBillGroup_Transco.LocationName"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_Transco.SchNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_Transco.stateRef"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_Transco.UnitTaxExempt"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_Transco.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FABankBillGroup_Transco.LocTaxRt"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =269
    Top =123
    Right =1838
    Bottom =938
    Left =-1
    Top =-1
    Right =1537
    Bottom =396
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =58
        Top =55
        Right =296
        Bottom =190
        Top =0
        Name ="FABankBillGroup_Transco"
        Name =""
    End
End
