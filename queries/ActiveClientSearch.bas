Operation =1
Option =0
Where ="(((Clients.clientType)<>\"Bank\" And (Clients.clientType)<>\"Vendor\" And (Clien"
    "ts.clientType)<>\"Remarketing\" And (Clients.clientType)<>\"FA\" And (Clients.cl"
    "ientType)<>\"Dead\"))"
Begin InputTables
    Name ="Clients"
    Name ="Salesperson"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.ClientShNm"
    Expression ="Clients.clientType"
    Expression ="Clients.clientState"
    Expression ="Clients.clientCity"
    Alias ="Edit"
    Expression ="\"Edit\""
    Expression ="Salesperson.SPInitials"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Salesperson"
    Expression ="Clients.LOinitial = Salesperson.SPInitials"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.clientCompanyName"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Edit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salesperson.SPInitials"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =149
    Top =0
    Right =1494
    Bottom =878
    Left =-1
    Top =-1
    Right =1327
    Bottom =544
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =206
        Bottom =485
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Salesperson"
        Name =""
    End
End
