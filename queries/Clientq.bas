Operation =1
Option =0
Where ="(((Clients.clientCompanyName)<>\"TEST\") AND ((Clients.clientType)=\"CLIENT\"))"
Begin InputTables
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientDBA"
    Expression ="Clients.clientCity"
    Expression ="Clients.clientState"
    Expression ="Clients.clientBecameDate"
    Expression ="Clients.clientID"
End
Begin OrderBy
    Expression ="Clients.clientCompanyName"
    Flag =0
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
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientBecameDate"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="mm/dd/yy"
    End
    Begin
        dbText "Name" ="Clients.clientID"
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
    Bottom =222
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
