Operation =1
Option =0
Where ="(((Units.UnitVIN) Is Not Null))"
Begin InputTables
    Name ="FleetStudyImportBase"
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="FleetStudyImportBase.UnitNumber"
    Alias ="AtlaasUnitNum"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.ObcUnitNumber"
    Expression ="FleetStudyImportBase.UnitYear"
    Alias ="AtlaasYear"
    Expression ="Units.UnitYr"
    Expression ="FleetStudyImportBase.VIN"
    Expression ="FleetStudyImportBase.Type"
    Expression ="FleetStudyImportBase.UnitMake"
    Alias ="AtlaasMake"
    Expression ="Units.UnitMake"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="Units.GroupREF"
End
Begin Joins
    LeftTable ="FleetStudyImportBase"
    RightTable ="Units"
    Expression ="FleetStudyImportBase.VIN = Units.UnitVIN"
    Flag =1
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitID"
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
        dbText "Name" ="FleetStudyImportBase.UnitMake"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitYear"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitNumber"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.VIN"
        dbInteger "ColumnWidth" ="2145"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AtlaasUnitNum"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AtlaasMake"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ObcUnitNumber"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AtlaasYear"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbInteger "ColumnWidth" ="2640"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbInteger "ColumnWidth" ="1470"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Type"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =72
    Top =31
    Right =1628
    Bottom =914
    Left =-1
    Top =-1
    Right =1524
    Bottom =413
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =66
        Top =26
        Right =352
        Bottom =517
        Top =0
        Name ="FleetStudyImportBase"
        Name =""
    End
    Begin
        Left =480
        Top =31
        Right =624
        Bottom =445
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =973
        Top =81
        Right =1097
        Bottom =410
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
