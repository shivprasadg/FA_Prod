Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Locations"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="LastTitleLogEntry"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="Units.UnitID"
    Alias ="UnitNo"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.Unitvin"
    Expression ="Units.UnitStatus"
    Expression ="Units.GroupREF"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & IIf(Nz([Axle],\"TBD\")=\"T"
        "BD\",Null,\" \" & [Axle]) & IIf(Nz([UnitSubType],\"\")=\"\",\"\",\" \" & [UnitSu"
        "bType]) & \" \" & [UnitType]"
    Alias ="Location"
    Expression ="Locations.Locationsn"
    Alias ="StreetAddress"
    Expression ="Locations.LAddrLine1"
    Alias ="City"
    Expression ="Locations.LCity"
    Alias ="State"
    Expression ="Locations.LState"
    Alias ="Zip"
    Expression ="Locations.LZip"
    Alias ="Title Num"
    Expression ="Nz([Units].[UnitTitleNumber],[LastTitleLogEntry].[TitleNumber])"
    Alias ="TitledState"
    Expression ="Locations_1.LState"
    Alias ="Status"
    Expression ="Units.unitstatus"
    Alias ="InService"
    Expression ="Units.UnitInServ"
    Alias ="UnitCost"
    Expression ="CCur([UnitFinalCost])"
    Expression ="Units.UnitInvoiceNum"
    Alias ="UnitInvoicePaid"
    Expression ="Units.UnitVendorPytDate"
    Alias ="InterimBank"
    Expression ="Units.UnitIFBank"
    Alias ="InterimDate"
    Expression ="Units.UnitIFDate"
    Alias ="InterimAmt"
    Expression ="Units.UnitIFAmt"
    Expression ="Units.DrawDownId"
    Expression ="Units.VendorREF"
    Alias ="SYNDRent"
    Expression ="Units.UNITSRENT"
    Alias ="PerDiemRent"
    Expression ="Units.UNITPDRENT"
    Alias ="LORent"
    Expression ="Units.UNITRENT"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_1"
    Expression ="Units.UnitTitleLocationRef = Locations_1.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="LastTitleLogEntry"
    Expression ="Units.UnitID = LastTitleLogEntry.AssetId"
    Flag =2
End
Begin OrderBy
    Expression ="Units.Unitvin"
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
        dbText "Name" ="Units.Unitvin"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="5715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="InService"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="StreetAddress"
        dbInteger "ColumnWidth" ="3390"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbInteger "ColumnWidth" ="1560"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="State"
        dbInteger "ColumnOrder" ="12"
        dbText "Format" =">"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zip"
        dbInteger "ColumnWidth" ="1485"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbInteger "ColumnWidth" ="1410"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Title Num"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="990"
        dbInteger "ColumnOrder" ="16"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitInvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitledState"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNo"
        dbInteger "ColumnWidth" ="1125"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LORent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerdiemRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SYNDRent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1489
    Bottom =1046
    Left =-1
    Top =-1
    Right =1457
    Bottom =335
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =271
        Top =18
        Right =452
        Bottom =343
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =494
        Top =237
        Right =638
        Bottom =381
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =725
        Top =167
        Right =869
        Bottom =389
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =497
        Top =19
        Right =641
        Bottom =213
        Top =0
        Name ="LastTitleLogEntry"
        Name =""
    End
End
