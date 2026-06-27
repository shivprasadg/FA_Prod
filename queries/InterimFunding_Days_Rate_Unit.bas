Operation =1
Option =0
Begin InputTables
    Name ="InterimFunding_Days_Rate"
End
Begin OutputColumns
    Expression ="InterimFunding_Days_Rate.UnitID"
    Expression ="InterimFunding_Days_Rate.GroupREF"
    Expression ="InterimFunding_Days_Rate.InterimAmount"
    Expression ="InterimFunding_Days_Rate.InterimIntRate"
    Expression ="InterimFunding_Days_Rate.InterimLoanDays"
    Alias ="UnitInterimExp"
    Expression ="IIf(Nz([InterimLoanDays],0)<1,0,[InterimAmount]+FV(([InterimIntRate]/360),[Inter"
        "imLoanDays],0,[InterimAmount],1))"
    Expression ="InterimFunding_Days_Rate.UnitCost"
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
        dbText "Name" ="InterimFunding_Days_Rate.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimFunding_Days_Rate.InterimAmount"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimFunding_Days_Rate.InterimIntRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimFunding_Days_Rate.InterimLoanDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitInterimExp"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="InterimFunding_Days_Rate.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimFunding_Days_Rate.UnitCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-286
    Top =61
    Right =955
    Bottom =1059
    Left =-1
    Top =-1
    Right =1223
    Bottom =646
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =134
        Top =63
        Right =582
        Bottom =363
        Top =0
        Name ="InterimFunding_Days_Rate"
        Name =""
    End
End
