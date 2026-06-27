Operation =1
Option =0
Having ="(((Clients.clientBecameDate) Is Not Null And (Clients.clientBecameDate)<#1/1/201"
    "0#) AND ((Deals.dealOutcome)=\"Awarded\"))"
Begin InputTables
    Name ="Clients"
    Name ="Deals"
    Name ="Documentation"
End
Begin OutputColumns
    Alias ="Client Yr"
    Expression ="DatePart(\"yyyy\",[dealoutcomedate])"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientBecameDate"
    Expression ="Deals.dealOutcome"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
    Flag =1
    LeftTable ="Deals"
    RightTable ="Documentation"
    Expression ="Deals.dealID = Documentation.dealRef"
    Flag =1
End
Begin OrderBy
    Expression ="DatePart(\"yyyy\",[dealoutcomedate])"
    Flag =0
    Expression ="Clients.clientCompanyName"
    Flag =0
End
Begin Groups
    Expression ="DatePart(\"yyyy\",[dealoutcomedate])"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Clients.clientBecameDate"
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
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="3960"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientBecameDate"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="mm/dd/yy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client Yr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =978
    Bottom =583
    Left =-1
    Top =-1
    Right =946
    Bottom =283
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =38
        Right =267
        Bottom =180
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =315
        Top =12
        Right =538
        Bottom =156
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =586
        Top =12
        Right =757
        Bottom =156
        Top =0
        Name ="Documentation"
        Name =""
    End
End
