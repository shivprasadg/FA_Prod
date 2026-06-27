Operation =1
Option =0
Where ="(((Locations.ClientREF)=Forms!UnitsForm!clientID) And ((Locations.LocationType)="
    "\"garage\") And ((Locations.LocObsolete)=No))"
Begin InputTables
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Locations.LocationID"
    Expression ="Locations.Locationsn"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
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
End
Begin
    State =0
    Left =38
    Top =86
    Right =1546
    Bottom =882
    Left =-1
    Top =-1
    Right =1476
    Bottom =445
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
