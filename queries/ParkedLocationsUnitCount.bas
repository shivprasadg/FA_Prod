Operation =1
Option =0
Begin InputTables
    Name ="Locations"
    Name ="Units"
    Name ="ContactsLinkedToLocation"
End
Begin OutputColumns
    Expression ="Locations.ClientREF"
    Expression ="Locations.ClientGroupID"
    Expression ="Locations.LocationID"
    Expression ="Locations.DivisionID"
    Expression ="Locations.LocationName"
    Alias ="LState"
    Expression ="Locations.LState"
    Alias ="Units"
    Expression ="IIf(Count([UnitID])>0,\"Y\",\"\")"
    Alias ="Contacts"
    Expression ="IIf(Count([ContactID])>0,\"Y\")"
    Alias ="Location"
    Expression ="[LocationName] & IIf([LocationName] Like '*' & [LState] & '*',\"\",\" \" & [LSta"
        "te])"
    Expression ="Locations.LocationType"
    Expression ="Locations.LocObsolete"
    Alias ="TaxRate"
    Expression ="Format([LTaxRate],\"Percent\")"
    Alias ="UnitCount"
    Expression ="Count(Units.unitId)"
End
Begin Joins
    LeftTable ="Locations"
    RightTable ="Units"
    Expression ="Locations.LocationID = Units.UnitGarageLocationRef"
    Flag =2
    LeftTable ="Locations"
    RightTable ="ContactsLinkedToLocation"
    Expression ="Locations.LocationID = ContactsLinkedToLocation.LocationID"
    Flag =2
End
Begin OrderBy
    Expression ="[LocationName] & IIf([LocationName] Like '*' & [LState] & '*',\"\",\" \" & [LSta"
        "te])"
    Flag =0
End
Begin Groups
    Expression ="Locations.ClientREF"
    GroupLevel =0
    Expression ="Locations.ClientGroupID"
    GroupLevel =0
    Expression ="Locations.LocationID"
    GroupLevel =0
    Expression ="Locations.DivisionID"
    GroupLevel =0
    Expression ="Locations.LocationName"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="Locations.LocationType"
    GroupLevel =0
    Expression ="Locations.LocObsolete"
    GroupLevel =0
    Expression ="Format([LTaxRate],\"Percent\")"
    GroupLevel =0
    Expression ="Locations.LCity"
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
        dbText "Name" ="Units"
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
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxRate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-209
    Top =52
    Right =1053
    Bottom =986
    Left =-1
    Top =-1
    Right =1244
    Bottom =491
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
        Left =376
        Top =15
        Right =683
        Bottom =542
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =748
        Top =44
        Right =964
        Bottom =188
        Top =0
        Name ="ContactsLinkedToLocation"
        Name =""
    End
End
