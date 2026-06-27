Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Locations"
End
Begin OutputColumns
    Alias ="ClientID"
    Expression ="Locations.ClientREF"
    Alias ="DivisionID"
    Expression ="Nz([Locations].[DivisionID],0)"
    Alias ="GarageLocationID"
    Expression ="Units.UnitGarageLocationRef"
    Expression ="Locations.LocationName"
    Expression ="Locations.Locationsn"
    Alias ="Units"
    Expression ="Count(Units.UnitID)"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
End
Begin Groups
    Expression ="Locations.ClientREF"
    GroupLevel =0
    Expression ="Nz([Locations].[DivisionID],0)"
    GroupLevel =0
    Expression ="Units.UnitGarageLocationRef"
    GroupLevel =0
    Expression ="Locations.LocationName"
    GroupLevel =0
    Expression ="Locations.Locationsn"
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
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DivisionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GarageLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1241
    Bottom =1021
    Left =-1
    Top =-1
    Right =1209
    Bottom =681
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =391
        Top =154
        Right =598
        Bottom =298
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =165
        Top =129
        Right =309
        Bottom =321
        Top =0
        Name ="Locations"
        Name =""
    End
End
