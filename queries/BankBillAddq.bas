Operation =1
Option =0
Where ="(((Locations.LocObsolete)=No) AND ((Clients.clienttype)=\"bank\") AND ((Location"
    "s.LocationType)=\"billing\"))"
Begin InputTables
    Name ="Clients"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Locations.LocObsolete"
    Expression ="Locations.*"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Locations"
    Expression ="Clients.clientID = Locations.ClientREF"
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
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.ClientREF"
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
        dbText "Name" ="Locations.District"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSubLessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt2"
        dbInteger "ColumnWidth" ="3240"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCounty"
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
        dbText "Name" ="Locations.LocTaxRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.locationPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.locationFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocDCs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocDCMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSlprs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.locSplrMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTrucks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTrkMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTrl"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTrlMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocReeferTrlrs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocRefrigUnitHPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsdate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsCmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.oldlocid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.RLocId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTaxOpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocCostCtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocABA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocACH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocBKACCT"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =346
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
End
