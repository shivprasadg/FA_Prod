Operation =1
Option =0
Where ="(((Units.ClientID)=[AskClientID]) AND ((Locations_1.ClientREF)=[AskClientID]))"
Begin InputTables
    Name ="Units"
    Name ="Locations"
    Name ="Clients"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Clients"
    Alias ="Clients_1"
End
Begin OutputColumns
    Expression ="Units.ClientID"
    Expression ="Units.UnitID"
    Expression ="Units.UnitVIN"
    Expression ="Locations.LocationName"
    Alias ="REALClientLoc"
    Expression ="Locations_1.LocationName"
    Expression ="Units.UnitGarageLocationRef"
    Alias ="RealLocID"
    Expression ="Locations_1.LocationID"
    Alias ="ListedClient"
    Expression ="Clients.ClientShNm"
    Alias ="RealClient"
    Expression ="Clients_1.ClientShNm"
    Alias ="ListedClientID"
    Expression ="Clients.clientID"
    Alias ="RealClientID"
    Expression ="Locations_1.ClientREF"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="Clients"
    Expression ="Locations.ClientREF = Clients.clientID"
    Flag =1
    LeftTable ="Locations_1"
    RightTable ="Clients_1"
    Expression ="Locations_1.ClientREF = Clients_1.clientID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="Locations_1"
    Expression ="Locations.LocationName = Locations_1.LocationName"
    Flag =1
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
dbMemo "Filter" ="([LocationCompare].[ListedClientID]=5216)"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ListedClientID"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RealClientID"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RealLocID"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ListedClient"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RealClient"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitGarageLocationRef"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="REALClientLoc"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="REAL"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GClientID"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ListedClientLoc"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.ClientREF"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1417
    Bottom =984
    Left =-1
    Top =-1
    Right =1418
    Bottom =607
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =92
        Top =52
        Right =455
        Bottom =375
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =494
        Top =107
        Right =638
        Bottom =251
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =700
        Top =199
        Right =844
        Bottom =343
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =708
        Top =355
        Right =852
        Bottom =548
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =901
        Top =405
        Right =1085
        Bottom =607
        Top =0
        Name ="Clients_1"
        Name =""
    End
End
