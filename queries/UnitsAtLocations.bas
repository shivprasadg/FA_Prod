Operation =1
Option =0
Where ="(((vw_ClientFleetList.Status)<>\"O\"<>\"S\"<>\"C\"<>\"KS\"))"
Begin InputTables
    Name ="vw_ClientFleetList"
    Name ="Locations"
    Name ="Units"
End
Begin OutputColumns
    Expression ="vw_ClientFleetList.Client"
    Expression ="vw_ClientFleetList.DC_Garage"
    Expression ="Locations.Locationsn"
    Alias ="City"
    Expression ="Locations.LCity"
    Alias ="State"
    Expression ="Locations.LState"
    Alias ="AllUnitsAt"
    Expression ="Count(1)"
    Alias ="NewUnits"
    Expression ="IIf([vw_ClientFleetList].[Status]=\"B\",1,0)"
    Alias ="FirstEstDelvery"
    Expression ="Min(Units.UnitEstDelDate)"
End
Begin Joins
    LeftTable ="vw_ClientFleetList"
    RightTable ="Locations"
    Expression ="vw_ClientFleetList.LocationID = Locations.LocationID"
    Flag =1
    LeftTable ="vw_ClientFleetList"
    RightTable ="Units"
    Expression ="vw_ClientFleetList.AssetID = Units.UnitID"
    Flag =1
End
Begin Groups
    Expression ="vw_ClientFleetList.Client"
    GroupLevel =0
    Expression ="vw_ClientFleetList.DC_Garage"
    GroupLevel =0
    Expression ="Locations.Locationsn"
    GroupLevel =0
    Expression ="Locations.LCity"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="IIf([vw_ClientFleetList].[Status]=\"B\",1,0)"
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
        dbText "Name" ="AllUnitsAt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstEstDelvery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewUnits"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =39
    Top =28
    Right =1516
    Bottom =1026
    Left =-1
    Top =-1
    Right =1445
    Bottom =714
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =73
        Top =28
        Right =524
        Bottom =343
        Top =0
        Name ="Client_FleetList"
        Name =""
    End
    Begin
        Left =611
        Top =207
        Right =755
        Bottom =563
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =659
        Top =42
        Right =945
        Bottom =539
        Top =0
        Name ="Units"
        Name =""
    End
End
