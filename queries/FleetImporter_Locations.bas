Operation =1
Option =2
Where ="(((Locations.ClientREF)=Forms!FleetStudy_Importer!ClientID) And ((Locations.Loca"
    "tionType)=\"garage\"))"
Begin InputTables
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Locations.LocationID"
    Expression ="Locations.LocationType"
    Expression ="Locations.LocationName"
    Expression ="Locations.Locationsn"
    Expression ="Locations.ClientREF"
    Expression ="Locations.LocationType"
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
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.ClientREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1350
    Bottom =1070
    Left =-1
    Top =-1
    Right =1332
    Bottom =729
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =351
        Bottom =279
        Top =0
        Name ="Locations"
        Name =""
    End
End
