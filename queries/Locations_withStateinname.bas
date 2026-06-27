Operation =1
Option =0
Where ="(((Locations.LCity) Like '*' & [LState] & '*') AND ((Locations.LState)<>\"\" And"
    " (Locations.LState)<>\"OK\" And (Locations.LState)<>\"HI\")) OR (((Locations.Loc"
    "ationName) Like '*' & [LState] & '*') AND ((Locations.LState)<>\"\" And (Locatio"
    "ns.LState)<>\"OK\" And (Locations.LState)<>\"HI\"))"
Begin InputTables
    Name ="Clients"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="Locations.LocationName"
    Expression ="Locations.LCity"
    Expression ="Locations.Locationsn"
    Expression ="Locations.LState"
    Expression ="Clients.clientID"
    Expression ="Locations.LocationID"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Locations"
    Expression ="Clients.clientID = Locations.ClientREF"
    Flag =2
    LeftTable ="Clients"
    RightTable ="Locations"
    Expression ="Clients.clientID = Locations.ClientREF"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="Locations.LocationName"
    Flag =0
    Expression ="Locations.LState"
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
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =32
    Top =11
    Right =1277
    Bottom =664
    Left =-1
    Top =-1
    Right =1213
    Bottom =588
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =289
        Top =107
        Right =694
        Bottom =516
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =39
        Top =84
        Right =183
        Bottom =228
        Top =0
        Name ="Clients"
        Name =""
    End
End
