Operation =1
Option =0
Begin InputTables
    Name ="Locations"
    Name ="UnitsParkedLocations"
    Name ="ContactsLinkedToLocationQuery"
End
Begin OutputColumns
    Expression ="Locations.ClientREF"
    Expression ="Locations.LocationID"
    Expression ="Locations.DivisionID"
    Expression ="Locations.LocationName"
    Expression ="Locations.LCity"
    Alias ="LState"
    Expression ="Locations.LState"
    Alias ="Units"
    Expression ="IIf(Sum([infleet])>0,\"Y\",\"\")"
    Alias ="Contacts"
    Expression ="IIf(Count([ContactID])>0,Count([ContactID]))"
    Expression ="Locations.LocationType"
    Expression ="Locations.LocObsolete"
End
Begin Joins
    LeftTable ="Locations"
    RightTable ="UnitsParkedLocations"
    Expression ="Locations.LocationID = UnitsParkedLocations.GarageLocationID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="ContactsLinkedToLocationQuery"
    Expression ="Locations.LocationID = ContactsLinkedToLocationQuery.LocationID"
    Flag =2
End
Begin OrderBy
    Expression ="Locations.LocationName"
    Flag =0
End
Begin Groups
    Expression ="Locations.ClientREF"
    GroupLevel =0
    Expression ="Locations.LocationID"
    GroupLevel =0
    Expression ="Locations.DivisionID"
    GroupLevel =0
    Expression ="Locations.LocationName"
    GroupLevel =0
    Expression ="Locations.LCity"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="Locations.LocationType"
    GroupLevel =0
    Expression ="Locations.LocObsolete"
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
        dbText "Name" ="Locations.LocObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.ClientREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.DivisionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =240
    Top =204
    Right =1502
    Bottom =1138
    Left =-1
    Top =-1
    Right =1230
    Bottom =446
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =306
        Bottom =431
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =477
        Top =73
        Right =857
        Bottom =446
        Top =0
        Name ="UnitsParkedLocations"
        Name =""
    End
    Begin
        Left =922
        Top =125
        Right =1066
        Bottom =269
        Top =0
        Name ="ContactsLinkedToLocationQuery"
        Name =""
    End
End
