Operation =1
Option =0
Where ="(((Deals.dealOutcome)=\"Awarded\") AND ((Clients.ClientGroupId)<>18))"
Begin InputTables
    Name ="Clients"
    Name ="Deals"
End
Begin OutputColumns
    Expression ="Deals.FAID"
    Alias ="Client"
    Expression ="Clients.clientcompanyname"
    Alias ="SentYear"
    Expression ="DatePart(\"yyyy\",[DealLPAFSent])"
    Alias ="DealYear"
    Expression ="DatePart(\"yyyy\",[dealoutcomedate])"
    Alias ="DealUnitQty"
    Expression ="Sum(Deals.dealUnitQty)"
    Alias ="DealGrandTotal"
    Expression ="Sum(Deals.dealGrandTotal)"
    Alias ="OutcomeDate"
    Expression ="Max(Deals.dealOutcomeDate)"
    Expression ="Deals.transTypeRef"
    Alias ="DealBEDEX"
    Expression ="Deals.DealLOEx"
    Expression ="Clients.ClientGroupId"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.clientcompanyname"
    Flag =0
    Expression ="Deals.FAID"
    Flag =0
End
Begin Groups
    Expression ="Deals.FAID"
    GroupLevel =0
    Expression ="Clients.clientcompanyname"
    GroupLevel =0
    Expression ="DatePart(\"yyyy\",[DealLPAFSent])"
    GroupLevel =0
    Expression ="DatePart(\"yyyy\",[dealoutcomedate])"
    GroupLevel =0
    Expression ="Deals.transTypeRef"
    GroupLevel =0
    Expression ="Deals.DealLOEx"
    GroupLevel =0
    Expression ="Clients.ClientGroupId"
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
        dbText "Name" ="Deals.faid"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnWidth" ="2805"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealGrandTotal"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealUnitQty"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SentYear"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealYear"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OutcomeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.transTypeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.DealLOEx"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealBEDEX"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-537
    Top =55
    Right =721
    Bottom =701
    Left =-1
    Top =-1
    Right =1234
    Bottom =152
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =17
        Top =0
        Right =204
        Bottom =166
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =221
        Top =3
        Right =383
        Bottom =216
        Top =0
        Name ="Deals"
        Name =""
    End
End
