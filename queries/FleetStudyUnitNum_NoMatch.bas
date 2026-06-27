Operation =1
Option =0
Where ="(((vw_SixKeys.Unitnum)<>[FleetStudyImportBase].[UnitNumber]))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="UMP_MilesFuelMPG"
    Name ="UnitVinDetails"
    Name ="FleetStudyImportBase"
End
Begin OutputColumns
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.VIN"
    Alias ="FAClient"
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="ImportUnitNum"
    Expression ="FleetStudyImportBase.UnitNumber"
    Alias ="FA_UnitNum"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.ObcUnitNumber"
    Alias ="FA_LastReport"
    Expression ="UMP_MilesFuelMPG.LastReport"
    Alias ="ImportDesc"
    Expression ="[FleetStudyImportBase].[UnitYear] & \" \" & [FleetStudyImportBase].[UnitMake] & "
        "\" \" & [FleetStudyImportBase].[UnitModel] & \" \" & [FleetStudyImportBase].[Typ"
        "e] & \" \" & [FleetStudyImportBase].[SubType]"
    Alias ="FADesc"
    Expression ="vw_SixKeys.UnitDescShortVer"
    Alias ="YearCk"
    Expression ="IIf(CStr([UnitYear])<>CStr([MYear]),\"Diff\",\"Match\")"
    Alias ="MakeCk"
    Expression ="IIf([vw_SixKeys].[Make]<>[UnitMake],\"Diff\",\"Match\")"
    Alias ="ModelCk"
    Expression ="IIf([UnitModel]<>[vw_SixKeys].[Model],\"Diff\",\"Match\")"
    Alias ="TypeCk"
    Expression ="IIf(Trim([FleetStudyImportBase].[SubType])<>Trim([vw_SixKeys].[SubType]),\"Diff\""
        ",\"Match\")"
    Alias ="FA_VinCheck"
    Expression ="IIf(Nz([UnitVinDetails].[model] & [UnitVinDetails].[make] & [UnitVinDetails].[ma"
        "nufacturer],\"\")<>\"\",\"Yes\",\"No\")"
    Expression ="UMP_MilesFuelMPG.LastReport"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientGroupId"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="UMP_MilesFuelMPG"
    Expression ="vw_SixKeys.AssetID = UMP_MilesFuelMPG.UnitRef"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="UnitVinDetails"
    Expression ="vw_SixKeys.UnitID = UnitVinDetails.unitid"
    Flag =2
    LeftTable ="FleetStudyImportBase"
    RightTable ="vw_SixKeys"
    Expression ="FleetStudyImportBase.VIN = vw_SixKeys.VIN"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="[FleetStudyImportBase].[UnitYear] & \" \" & [FleetStudyImportBase].[UnitMake] & "
        "\" \" & [FleetStudyImportBase].[UnitModel] & \" \" & [FleetStudyImportBase].[Typ"
        "e] & \" \" & [FleetStudyImportBase].[SubType]"
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
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImportUnitNum"
        dbInteger "ColumnWidth" ="1890"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FaDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImportDesc"
        dbInteger "ColumnWidth" ="4305"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FA_VinCheck"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FA_UnitNum"
        dbInteger "ColumnWidth" ="1605"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MakeCk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAClient"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="fA_LastReport"
        dbInteger "ColumnWidth" ="1755"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ObcUnitNumber"
        dbInteger "ColumnWidth" ="1935"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModelCk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YearCk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TypeCk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UMP_MilesFuelMPG.LastReport"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =135
    Top =18
    Right =2441
    Bottom =941
    Left =-1
    Top =-1
    Right =2274
    Bottom =582
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =772
        Top =95
        Right =916
        Bottom =296
        Top =0
        Name ="UMP_MilesFuelMPG"
        Name =""
    End
    Begin
        Left =773
        Top =309
        Right =1011
        Bottom =651
        Top =0
        Name ="UnitVinDetails"
        Name =""
    End
    Begin
        Left =9
        Top =-3
        Right =153
        Bottom =205
        Top =0
        Name ="FleetStudyImportBase"
        Name =""
    End
End
