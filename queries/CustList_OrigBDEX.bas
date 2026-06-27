Operation =1
Option =0
Having ="(((Deals.dealOutcome)=\"Awarded\"))"
Begin InputTables
    Name ="Deals"
End
Begin OutputColumns
    Alias ="MinOfdealOutcomeDate"
    Expression ="Min(Deals.dealOutcomeDate)"
    Expression ="Deals.clientRef"
    Alias ="OrigBDEX"
    Expression ="First(Deals.DealLOEx)"
    Expression ="Deals.dealOutcome"
End
Begin OrderBy
    Expression ="Deals.clientRef"
    Flag =0
End
Begin Groups
    Expression ="Deals.clientRef"
    GroupLevel =0
    Expression ="Deals.dealOutcome"
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
        dbText "Name" ="Deals.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfdealOutcomeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrigBDEX"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =110
    Top =-6
    Right =1568
    Bottom =854
    Left =-1
    Top =-1
    Right =1426
    Bottom =697
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =271
        Top =65
        Right =632
        Bottom =663
        Top =0
        Name ="Deals"
        Name =""
    End
End
