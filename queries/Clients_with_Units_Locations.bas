Operation =1
Option =0
Begin InputTables
    Name ="Clients_with_Units_Address"
    Name ="Locations"
    Name ="ClientDivisions"
    Name ="Location_Divisions_UnitCount"
End
Begin OutputColumns
    Expression ="Clients_with_Units_Address.ClientGroupId"
    Expression ="Clients_with_Units_Address.clientID"
    Alias ="DivisionID"
    Expression ="Nz([ClientDivisions].[DivisionID],0)"
    Expression ="Locations.LocationID"
    Expression ="Clients_with_Units_Address.clientCompanyName"
    Alias ="DivisionName"
    Expression ="Nz([ClientDivisions].[DivisionName],\"All\")"
    Expression ="Locations.LocationType"
    Expression ="Locations.LocationName"
    Expression ="Location_Divisions_UnitCount.Locationsn"
    Alias ="Units@Location"
    Expression ="Location_Divisions_UnitCount.Units"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LAddrLine2"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
End
Begin Joins
    LeftTable ="Clients_with_Units_Address"
    RightTable ="Locations"
    Expression ="Clients_with_Units_Address.clientID = Locations.ClientREF"
    Flag =1
    LeftTable ="Locations"
    RightTable ="ClientDivisions"
    Expression ="Locations.DivisionID = ClientDivisions.DivisionID"
    Flag =2
    LeftTable ="Locations"
    RightTable ="Location_Divisions_UnitCount"
    Expression ="Locations.LocationID = Location_Divisions_UnitCount.GarageLocationID"
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
Begin
    Begin
        dbText "Name" ="Clients_with_Units_Address.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_with_Units_Address.clientCompanyName"
        dbInteger "ColumnWidth" ="3120"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_with_Units_Address.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DivisionName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units@Location"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DivisionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location_Divisions_UnitCount.Locationsn"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-50
    Top =37
    Right =1459
    Bottom =888
    Left =-1
    Top =-1
    Right =1477
    Bottom =359
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =8
        Top =11
        Right =176
        Bottom =377
        Top =0
        Name ="Clients_with_Units_Address"
        Name =""
    End
    Begin
        Left =219
        Top =90
        Right =363
        Bottom =435
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =423
        Top =255
        Right =567
        Bottom =399
        Top =0
        Name ="ClientDivisions"
        Name =""
    End
    Begin
        Left =424
        Top =71
        Right =568
        Bottom =249
        Top =0
        Name ="Location_Divisions_UnitCount"
        Name =""
    End
End
