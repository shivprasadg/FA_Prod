Operation =1
Option =0
Where ="(((Deals.dealOutcome)=\"Awarded\"))"
Having ="(((Max(Deals.dealOutcomeDate)) Is Not Null))"
Begin InputTables
    Name ="Clients"
    Name ="Deals"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Deals.FAID"
    Expression ="Deals.DealLOEx"
    Expression ="Clients.clientBecameDate"
    Alias ="dealOutcomeDate"
    Expression ="Max(Deals.dealOutcomeDate)"
    Alias ="ClientAgeAtDeal"
    Expression ="DateDiff(\"m\",[clientBecameDate],[dealOutcomeDate])"
    Alias ="DealCommission"
    Expression ="IIf(DateDiff(\"m\",[clientBecameDate],[dealOutcomeDate])>12,CDbl(0.08),CDbl(0.12"
        "))"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.clientID"
    Flag =0
    Expression ="Max(Deals.dealOutcomeDate)"
    Flag =0
End
Begin Groups
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="Deals.FAID"
    GroupLevel =0
    Expression ="Deals.DealLOEx"
    GroupLevel =0
    Expression ="Clients.clientBecameDate"
    GroupLevel =0
    Expression ="IIf(DateDiff(\"m\",[clientBecameDate],[dealOutcomeDate])>12,CDbl(0.08),CDbl(0.12"
        "))"
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
        dbText "Name" ="ClientAgeAtDeal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealCommission"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientBecameDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.DealLOEx"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealOutcomeDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-285
    Top =135
    Right =991
    Bottom =958
    Left =-1
    Top =-1
    Right =1258
    Bottom =478
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =40
        Top =65
        Right =184
        Bottom =472
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =249
        Top =87
        Right =488
        Bottom =505
        Top =0
        Name ="Deals"
        Name =""
    End
End
