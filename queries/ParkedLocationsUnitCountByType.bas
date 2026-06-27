Operation =1
Option =0
Begin InputTables
    Name ="Locations"
    Name ="Clients"
    Name ="vw_SixKeys"
    Name ="ClientGroups"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Clients.ClientShNm"
    Alias ="Location"
    Expression ="Locations.LocationName"
    Alias ="City"
    Expression ="Locations.LCity"
    Alias ="State"
    Expression ="Locations.LState"
    Alias ="Units"
    Expression ="IIf(Count([UnitID])>0,Count([UnitID]),0)"
    Alias ="FAUnits"
    Expression ="Sum(IIf([Status]=\"A\" Or [Status]=\"TR\" And [MLorig]=\"FA\",1,0))"
    Alias ="FANewUnits"
    Expression ="Sum(IIf([Status]=\"B\",1,0))"
    Alias ="DayCabs"
    Expression ="Sum(IIf([SubType]=\"Daycab\" And [Status]=\"A\" And [MLorig]=\"FA\" Or [SubType]"
        "=\"Daycab\" And [Status]=\"TR\" And [MLorig]=\"FA\",1,0))"
    Alias ="Sleepers"
    Expression ="Sum(IIf([SubType]=\"Sleeper\" And [Status]=\"A\" And [MLorig]=\"FA\" Or [SubType"
        "]=\"Sleeper\" And [Status]=\"TR\" And [MLorig]=\"FA\",1,0))"
    Alias ="Straight"
    Expression ="Sum(IIf([Type]=\"Straight Truck\" And [Status]=\"A\" And [MLorig]=\"FA\" Or [Typ"
        "e]=\"Straight Truck\" And [Status]=\"TR\" And [MLorig]=\"FA\",1,0))"
    Alias ="Van"
    Expression ="Sum(IIf([Type]=\"Van\" And [Status]=\"A\" And [MLorig]=\"FA\" Or [Type]=\"Van\" "
        "And [Status]=\"TR\" And [MLorig]=\"FA\",1,0))"
    Alias ="Yard"
    Expression ="Sum(IIf([Type]=\"Yard Tractor\" And [Status]=\"A\" And [MLorig]=\"FA\" Or [Type]"
        "=\"Yard Trator\" And [Status]=\"TR\" And [MLorig]=\"FA\",1,0))"
    Alias ="Trailers"
    Expression ="Sum(IIf([Type]=\"Trailer\" And [Status]=\"A\" And [MLorig]=\"FA\" Or [Type]=\"Tr"
        "ailer\" And [Status]=\"TR\" And [MLorig]=\"FA\",1,0))"
    Alias ="FirstDelvery"
    Expression ="Min(vw_SixKeys.ActualDeliveryDate)"
    Alias ="LastDelvery"
    Expression ="Max(vw_SixKeys.ActualDeliveryDate)"
    Alias ="LastEstDelDate"
    Expression ="Max(vw_SixKeys.EsitmatedDeliveryDate)"
    Alias ="RMKTUnits"
    Expression ="Sum(IIf([Status]=\"PS\" Or [Status]=\"R\",1,0))"
    Alias ="TrackingUnits"
    Expression ="Sum(IIf([Status]=\"T\" And [mlorig]<>\"FA\",1,0))"
    Alias ="T_DayCabs"
    Expression ="Sum(IIf([SubType]=\"Daycab\" And [Status]=\"T\" And [MLorig]<>\"FA\",1,0))"
    Alias ="T_Sleepers"
    Expression ="Sum(IIf([SubType]=\"Sleeper\" And [Status]=\"T\" And [MLorig]<>\"FA\",1,0))"
    Alias ="T_Straight"
    Expression ="Sum(IIf([Type]=\"Straight Truck\" And [Status]=\"T\" And [MLorig]<>\"FA\",1,0))"
    Alias ="T_Van"
    Expression ="Sum(IIf([Type]=\"Van\" And [Status]=\"T\" And [MLorig]<>\"FA\",1,0))"
    Alias ="T_Yard"
    Expression ="Sum(IIf([Type]=\"Yard Tractor\" And [Status]=\"T\" And [MLorig]<>\"FA\",1,0))"
    Alias ="T_Trailers"
    Expression ="Sum(IIf([Type]=\"Trailer\" And [Status]=\"T\" And [MLorig]<>\"FA\",1,0))"
    Expression ="ClientGroups.ClientGroupName"
    Expression ="Locations.LocationID"
    Expression ="Clients.clientID"
    Expression ="Locations.ClientGroupID"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="vw_SixKeys"
    Expression ="Clients.clientID = vw_SixKeys.clientID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="vw_SixKeys"
    Expression ="Locations.LocationID = vw_SixKeys.LocationID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="ClientGroups"
    Expression ="Locations.ClientGroupID = ClientGroups.ClientGroupId"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="Locations.LocationName"
    Flag =0
    Expression ="Locations.LCity"
    Flag =0
    Expression ="Locations.LState"
    Flag =0
End
Begin Groups
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="Locations.LocationName"
    GroupLevel =0
    Expression ="Locations.LCity"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="ClientGroups.ClientGroupName"
    GroupLevel =0
    Expression ="Locations.LocationID"
    GroupLevel =0
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="Locations.ClientGroupID"
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
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FirstDelvery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrackingUnits"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastDelvery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTUnits"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trailers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Straight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sleepers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Sleepers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Straight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FANewUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Trailers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DayCabs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DayCabs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientGroups.ClientGroupName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yard"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Van"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Van"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Yard"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =105
    Top =74
    Right =1367
    Bottom =1008
    Left =-1
    Top =-1
    Right =1238
    Bottom =338
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =102
        Top =14
        Right =360
        Bottom =433
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =451
        Top =23
        Right =595
        Bottom =167
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =643
        Top =12
        Right =787
        Bottom =287
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =573
        Top =214
        Right =717
        Bottom =358
        Top =0
        Name ="ClientGroups"
        Name =""
    End
End
