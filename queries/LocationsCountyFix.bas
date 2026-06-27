Operation =4
Option =0
Begin InputTables
    Name ="Locations"
    Name ="LocationsCountyTemp"
End
Begin OutputColumns
    Name ="Locations.LCounty"
    Expression ="[LocationsCountyTemp].[LookCounty]"
    Name ="Locations.LCountyID"
    Expression ="[LocationsCountyTemp].[LookCountyID]"
End
Begin Joins
    LeftTable ="Locations"
    RightTable ="LocationsCountyTemp"
    Expression ="Locations.LocationID = LocationsCountyTemp.LocationID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="[Query11].[LCountyID]"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCounty"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LStateID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCountyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCityID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationsCountyTemp.LookCountyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationsCountyTemp.LookCounty"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =552
    Top =61
    Right =2030
    Bottom =778
    Left =-1
    Top =-1
    Right =1460
    Bottom =423
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =267
        Top =4
        Right =411
        Bottom =433
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =6
        Top =12
        Right =150
        Bottom =156
        Top =0
        Name ="LocationsCountyTemp"
        Name =""
    End
End
