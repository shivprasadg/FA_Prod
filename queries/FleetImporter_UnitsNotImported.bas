Operation =1
Option =0
Where ="(((vw_SixKeys.VIN) Is Null))"
Begin InputTables
    Name ="FleetStudyImportBase"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="FleetStudyImportBase.UnitNumber"
    Expression ="FleetStudyImportBase.OBCUnitNumber"
    Expression ="FleetStudyImportBase.VIN"
    Expression ="FleetStudyImportBase.UnitYear"
    Expression ="FleetStudyImportBase.UnitMake"
    Expression ="FleetStudyImportBase.UnitModel"
    Expression ="FleetStudyImportBase.Type"
    Expression ="FleetStudyImportBase.SubType"
    Expression ="FleetStudyImportBase.InServiceDate"
    Expression ="FleetStudyImportBase.Axle"
    Expression ="FleetStudyImportBase.Location"
End
Begin Joins
    LeftTable ="FleetStudyImportBase"
    RightTable ="vw_SixKeys"
    Expression ="FleetStudyImportBase.VIN = vw_SixKeys.VIN"
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
        dbText "Name" ="FleetStudyImportBase.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.SubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.OBCUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Type"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =71
    Top =22
    Right =1532
    Bottom =1003
    Left =-1
    Top =-1
    Right =1429
    Bottom =697
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =67
        Top =20
        Right =254
        Bottom =280
        Top =0
        Name ="FleetStudyImportBase"
        Name =""
    End
    Begin
        Left =435
        Top =-8
        Right =790
        Bottom =418
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
