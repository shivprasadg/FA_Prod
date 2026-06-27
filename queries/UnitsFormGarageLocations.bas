Operation =1
Option =0
Where ="(((Locations.LocationType)=\"Garage\") AND ((Locations.ClientREF) Is Not Null) A"
    "ND ((Locations.LocObsolete)=No))"
Begin InputTables
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Locations.LocationID"
    Alias ="Loc Nm"
    Expression ="Locations.LocationName"
    Alias ="Str Add"
    Expression ="Locations.LAddrLine1"
    Alias ="City"
    Expression ="Locations.LCity"
    Alias ="St"
    Expression ="Locations.LState"
    Alias ="Zip"
    Expression ="Locations.LZip"
End
Begin OrderBy
    Expression ="Locations.LState"
    Flag =0
    Expression ="Locations.LCity"
    Flag =0
    Expression ="Locations.LocationName"
    Flag =0
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
Begin
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Loc Nm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Str Add"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="St"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zip"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-488
    Top =22
    Right =1003
    Bottom =1003
    Left =-1
    Top =-1
    Right =1459
    Bottom =664
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
End
