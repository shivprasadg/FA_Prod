Operation =1
Option =0
Where ="(((FARentBillAssignedQ.GroupID)=[Forms]![Billing]![GroupId]))"
Begin InputTables
    Name ="FARentBillAssignedQ"
End
Begin OutputColumns
    Expression ="FARentBillAssignedQ.LocationName"
    Expression ="FARentBillAssignedQ.SchNo"
    Alias ="GroupNo"
    Expression ="FARentBillAssignedQ.SGrpNO"
    Expression ="FARentBillAssignedQ.LState"
    Expression ="FARentBillAssignedQ.UnitTaxExempt"
    Expression ="FARentBillAssignedQ.StTaxType"
    Expression ="FARentBillAssignedQ.LTaxRate"
    Alias ="UnitCount"
    Expression ="Count(FARentBillAssignedQ.Unitvin)"
    Alias ="LineTot"
    Expression ="Sum(FARentBillAssignedQ.UnitTotalRent)"
    Alias ="FSLFee"
    Expression ="Sum(FARentBillAssignedQ.FSLFee)"
    Alias ="RentNoTax"
    Expression ="Sum(FARentBillAssignedQ.RentNoTax)"
    Alias ="TaxAmt"
    Expression ="Sum(FARentBillAssignedQ.TaxAmt)"
    Alias ="UnitTotalRent"
    Expression ="Sum(FARentBillAssignedQ.UnitTotalRent)"
    Alias ="FSLTax"
    Expression ="Sum(FARentBillAssignedQ.FSLTax)"
End
Begin Groups
    Expression ="FARentBillAssignedQ.LocationName"
    GroupLevel =0
    Expression ="FARentBillAssignedQ.SchNo"
    GroupLevel =0
    Expression ="FARentBillAssignedQ.SGrpNO"
    GroupLevel =0
    Expression ="FARentBillAssignedQ.LState"
    GroupLevel =0
    Expression ="FARentBillAssignedQ.UnitTaxExempt"
    GroupLevel =0
    Expression ="FARentBillAssignedQ.StTaxType"
    GroupLevel =0
    Expression ="FARentBillAssignedQ.LTaxRate"
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
        dbText "Name" ="FARentBillAssignedQ.LocationName"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.SchNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.UnitTaxExempt"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FARentBillAssignedQ.StTaxType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =46
    Right =1657
    Bottom =861
    Left =-1
    Top =-1
    Right =1537
    Bottom =362
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =47
        Top =40
        Right =242
        Bottom =394
        Top =0
        Name ="FARentBillAssignedQ"
        Name =""
    End
End
