Operation =1
Option =0
Where ="(((vw_SixKeys.clientID)=[Forms]![FleetStudy_Importer]![clientID]))"
Begin InputTables
    Name ="FleetStudyImport"
    Name ="vw_SixKeys"
    Name ="UMP_MilesFuelMPG"
    Name ="UnitVinDetails"
End
Begin OutputColumns
    Expression ="vw_SixKeys.clientID"
    Alias ="FAClient"
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="FAUnitStatus"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.ObcUnitNumber"
    Alias ="ImportDesc"
    Expression ="[UnitYear] & \" \" & [FleetStudyImport].[UnitMake] & \" \" & [FleetStudyImport]."
        "[UnitModel] & \" \" & [FleetStudyImport].[Type] & \" \" & [FleetStudyImport].[Su"
        "bType]"
    Alias ="FAUnitDesc"
    Expression ="vw_SixKeys.UnitDescShortVer"
    Alias ="ImportVin"
    Expression ="FleetStudyImport.VIN"
    Alias ="FA_Vin"
    Expression ="vw_SixKeys.VIN"
    Alias ="VinDiff"
    Expression ="[vw_SixKeys].[VIN]<>[FleetStudyImport].[VIN]"
    Alias ="FA_VinCheck"
    Expression ="IIf(Nz([UnitVinDetails].[model] & [UnitVinDetails].[make] & [UnitVinDetails].[ma"
        "nufacturer],\"\")<>\"\",\"Yes\",\"No\")"
    Alias ="FA_LastReport"
    Expression ="UMP_MilesFuelMPG.LastReport"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientGroupId"
End
Begin Joins
    LeftTable ="FleetStudyImport"
    RightTable ="vw_SixKeys"
    Expression ="FleetStudyImport.UnitNumber = vw_SixKeys.Unitnum"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="UMP_MilesFuelMPG"
    Expression ="vw_SixKeys.AssetID = UMP_MilesFuelMPG.UnitRef"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="UnitVinDetails"
    Expression ="vw_SixKeys.UnitID = UnitVinDetails.unitid"
    Flag =2
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="vw_SixKeys.Status"
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
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="ImportVin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="FAUnitDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="FAUnitStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="FAClient"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="ImportDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4305"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="FA_Vin"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FA_VinCheck"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="fA_LastReport"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VinDiff"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ObcUnitNumber"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =119
    Top =134
    Right =2290
    Bottom =941
    Left =-1
    Top =-1
    Right =2139
    Bottom =527
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =41
        Top =34
        Right =318
        Bottom =625
        Top =0
        Name ="FleetStudyImport"
        Name =""
    End
    Begin
        Left =375
        Top =38
        Right =656
        Bottom =439
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =756
        Top =63
        Right =900
        Bottom =264
        Top =0
        Name ="UMP_MilesFuelMPG"
        Name =""
    End
    Begin
        Left =758
        Top =282
        Right =996
        Bottom =624
        Top =0
        Name ="UnitVinDetails"
        Name =""
    End
End
