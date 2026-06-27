Operation =1
Option =0
Where ="(((Deals.dealOutcomeDate) Is Not Null) AND ((Deals.dealOutcome)=\"Awarded\"))"
Begin InputTables
    Name ="Clients"
    Name ="Deals"
End
Begin OutputColumns
    Expression ="Deals.FAID"
    Alias ="Client"
    Expression ="Clients.clientcompanyname"
    Alias ="LPAFQtr"
    Expression ="DatePart(\"q\",[DealLPAFSent])"
    Alias ="LPAFYear"
    Expression ="DatePart(\"yyyy\",[DealLPAFSent])"
    Alias ="DealQtr"
    Expression ="DatePart(\"q\",[dealoutcomedate])"
    Alias ="DealYear"
    Expression ="DatePart(\"yyyy\",[dealoutcomedate])"
    Alias ="DealUnitCount"
    Expression ="Sum(Deals.dealUnitQty)"
    Alias ="DealGrandTotal"
    Expression ="Sum(Deals.dealGrandTotal)"
    Expression ="Clients.clientID"
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
    Expression ="DatePart(\"q\",[DealLPAFSent])"
    GroupLevel =0
    Expression ="DatePart(\"yyyy\",[DealLPAFSent])"
    GroupLevel =0
    Expression ="DatePart(\"q\",[dealoutcomedate])"
    GroupLevel =0
    Expression ="DatePart(\"yyyy\",[dealoutcomedate])"
    GroupLevel =0
    Expression ="Clients.clientID"
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
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealGrandTotal"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealUnitCount"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LPAFYear"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealYear"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LPAFQTR"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealQtr"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =84
    Top =185
    Right =1536
    Bottom =831
    Left =-1
    Top =-1
    Right =1434
    Bottom =81
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
        Left =286
        Top =15
        Right =448
        Bottom =228
        Top =0
        Name ="Deals"
        Name =""
    End
End
