Operation =1
Option =0
Where ="(((vw_SixKeys.clientID)=[Forms]![FleetStudy_Importer]![clientID]))"
Begin InputTables
    Name ="FleetImporter_UnitsNotImported"
    Name ="vw_SixKeys"
    Name ="dbo_vw_LastMilesReportedAll"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="FleetImporter_UnitsNotImported.UnitNumber"
    Expression ="FleetImporter_UnitsNotImported.OBCUnitNumber"
    Expression ="FleetImporter_UnitsNotImported.VIN"
    Alias ="UnitYear"
    Expression ="CInt([FleetImporter_UnitsNotImported].[UnitYear])"
    Expression ="FleetImporter_UnitsNotImported.UnitMake"
    Expression ="FleetImporter_UnitsNotImported.UnitModel"
    Expression ="FleetImporter_UnitsNotImported.Type"
    Expression ="FleetImporter_UnitsNotImported.SubType"
    Expression ="vw_SixKeys.AssetID"
    Alias ="FA_OBCNum"
    Expression ="vw_SixKeys.ObcUnitNumber"
    Alias ="FA_Status"
    Expression ="vw_SixKeys.Status"
    Alias ="OBC_LastReport"
    Expression ="dbo_vw_LastMilesReportedAll.LastReport"
    Alias ="FA_ModelYear"
    Expression ="vw_SixKeys.MYear"
    Alias ="FA_Make"
    Expression ="vw_SixKeys.Make"
    Alias ="FA_Model"
    Expression ="vw_SixKeys.Model"
    Alias ="FA_SubType"
    Expression ="vw_SixKeys.SubType"
    Alias ="FA_Type"
    Expression ="vw_SixKeys.Type"
    Alias ="FA_VIN"
    Expression ="vw_SixKeys.VIN"
End
Begin Joins
    LeftTable ="FleetImporter_UnitsNotImported"
    RightTable ="vw_SixKeys"
    Expression ="FleetImporter_UnitsNotImported.UnitNumber = vw_SixKeys.Unitnum"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="dbo_vw_LastMilesReportedAll"
    Expression ="vw_SixKeys.AssetID = dbo_vw_LastMilesReportedAll.UnitID"
    Flag =2
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
        dbText "Name" ="FA_ModelYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="FA_Model"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="FleetImporter_UnitsNotImported.SubType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="FleetImporter_UnitsNotImported.UnitNumber"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="FleetImporter_UnitsNotImported.OBCUnitNumber"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="FleetImporter_UnitsNotImported.VIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="FA_OBCNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="FA_Type"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="OBC_LastReport"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="FleetImporter_UnitsNotImported.UnitMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="FleetImporter_UnitsNotImported.UnitModel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="FleetImporter_UnitsNotImported.Type"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="FA_VIN"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="FA_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="FA_Make"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="FA_SubType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="UnitYear"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-2
    Top =-9
    Right =1667
    Bottom =989
    Left =-1
    Top =-1
    Right =1637
    Bottom =714
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =98
        Top =37
        Right =272
        Bottom =258
        Top =0
        Name ="FleetImporter_UnitsNotImported"
        Name =""
    End
    Begin
        Left =381
        Top =12
        Right =709
        Bottom =482
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =846
        Top =154
        Right =1132
        Bottom =528
        Top =0
        Name ="dbo_vw_LastMilesReportedAll"
        Name =""
    End
End
