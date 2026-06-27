Operation =1
Option =0
Begin InputTables
    Name ="Locations"
    Name ="vw_sixkeys"
End
Begin OutputColumns
    Expression ="vw_sixkeys.CompanyName"
    Expression ="Locations.LocationID"
    Expression ="Locations.Locationsn"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LAddrLine1"
    Alias ="Zip"
    Expression ="Locations.LZip"
    Alias ="Client"
    Expression ="vw_sixkeys.ClientShortName"
    Alias ="Units"
    Expression ="Count(vw_sixkeys.UnitID)"
    Expression ="Locations.Longitude"
    Expression ="Locations.Latitude"
    Expression ="Locations.ClientGroupId"
    Expression ="vw_sixkeys.clientID"
End
Begin Joins
    LeftTable ="Locations"
    RightTable ="vw_sixkeys"
    Expression ="Locations.LocationID = vw_sixkeys.LocationID"
    Flag =2
End
Begin Groups
    Expression ="vw_sixkeys.CompanyName"
    GroupLevel =0
    Expression ="Locations.LocationID"
    GroupLevel =0
    Expression ="Locations.Locationsn"
    GroupLevel =0
    Expression ="Locations.LocationStreet"
    GroupLevel =0
    Expression ="Locations.LAddrLine1"
    GroupLevel =0
    Expression ="Locations.LZip"
    GroupLevel =0
    Expression ="vw_sixkeys.ClientShortName"
    GroupLevel =0
    Expression ="Locations.Longitude"
    GroupLevel =0
    Expression ="Locations.Latitude"
    GroupLevel =0
    Expression ="Locations.ClientGroupId"
    GroupLevel =0
    Expression ="vw_sixkeys.clientID"
    GroupLevel =0
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
dbMemo "Filter" ="([GeoCodeLocations].[Longitude] Is Null)"
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
        dbText "Name" ="Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_sixkeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine1"
        dbInteger "ColumnWidth" ="3180"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="3885"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Longitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Latitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.ClientREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_sixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_sixkeys.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =12
    Top =-26
    Right =1624
    Bottom =925
    Left =-1
    Top =-1
    Right =1594
    Bottom =616
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =397
        Top =0
        Name ="vw_sixkeys"
        Name =""
    End
End
