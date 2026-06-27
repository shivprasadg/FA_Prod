Operation =1
Option =0
Where ="(((Clients.clientType)<>\"VENDOR\" And (Clients.clientType)<>\"BANK\" And (Clien"
    "ts.clientType)<>\"dead\" And (Clients.clientType)<>\"Remarketing\"))"
Begin InputTables
    Name ="Clients"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientID"
    Expression ="Clients.ClientShNm"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1158
    Bottom =969
    Left =-1
    Top =-1
    Right =1083
    Bottom =561
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
End
