Operation =1
Option =0
Where ="(((Clients.creditFAScore)>=0))"
Begin InputTables
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.creditFAScore"
    Expression ="Clients.ClientGroupID"
    Expression ="Clients.ClientShNm"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.creditFAProfile"
End
Begin OrderBy
    Expression ="Clients.creditFAScore"
    Flag =0
    Expression ="Clients.ClientGroupID"
    Flag =0
    Expression ="Clients.ClientShNm"
    Flag =0
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
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="3165"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAProfile"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAScore"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[CLIENT GROUP]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =364
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =408
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
End
