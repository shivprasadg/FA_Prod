Operation =1
Option =0
Where ="(((FleetImporter_ExisitingUnitNums.UnitUnitNum) Is Not Null))"
Begin InputTables
    Name ="FleetStudyNewVinsBase"
    Name ="FleetImporter_ExisitingUnitNums"
    Name ="vw_SixKeys"
    Name ="UnitLastOBCReport"
    Name ="UnitVinDecoder"
End
Begin OutputColumns
    Expression ="FleetStudyNewVinsBase.VINLenght"
    Expression ="UnitLastOBCReport.LogActivityDate"
    Alias ="DecodeStatus"
    Expression ="UnitVinDecoder.status"
    Expression ="FleetStudyNewVinsBase.UnitNumber"
    Expression ="FleetStudyNewVinsBase.OBCUnitNumber"
    Expression ="FleetStudyNewVinsBase.VIN"
    Expression ="FleetStudyNewVinsBase.UnitYear"
    Expression ="FleetStudyNewVinsBase.UnitMake"
    Expression ="FleetStudyNewVinsBase.UnitModel"
    Expression ="FleetStudyNewVinsBase.Type"
    Expression ="FleetStudyNewVinsBase.SubType"
    Expression ="FleetStudyNewVinsBase.InServiceDate"
    Expression ="FleetStudyNewVinsBase.Axle"
    Expression ="FleetStudyNewVinsBase.UnitRent"
    Expression ="FleetStudyNewVinsBase.LeaseStartDate"
    Expression ="FleetStudyNewVinsBase.LeaseEndDate"
    Expression ="FleetStudyNewVinsBase.Unitcost"
    Expression ="FleetStudyNewVinsBase.Division"
    Expression ="FleetStudyNewVinsBase.[LeaseTerm ]"
    Expression ="FleetStudyNewVinsBase.[Odometer ]"
    Expression ="FleetStudyNewVinsBase.OdometerDate"
    Expression ="FleetStudyNewVinsBase.Location"
    Expression ="FleetStudyNewVinsBase.State"
    Expression ="FleetStudyNewVinsBase.Plate"
    Expression ="FleetStudyNewVinsBase.ReeferHours"
    Expression ="FleetImporter_ExisitingUnitNums.UnitUnitNum"
    Expression ="vw_SixKeys.UnitDescShortVer"
    Expression ="vw_SixKeys.VIN"
End
Begin Joins
    LeftTable ="FleetStudyNewVinsBase"
    RightTable ="FleetImporter_ExisitingUnitNums"
    Expression ="FleetStudyNewVinsBase.UnitNumber = FleetImporter_ExisitingUnitNums.UnitUnitNum"
    Flag =2
    LeftTable ="FleetImporter_ExisitingUnitNums"
    RightTable ="vw_SixKeys"
    Expression ="FleetImporter_ExisitingUnitNums.UnitID = vw_SixKeys.AssetID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="UnitLastOBCReport"
    Expression ="vw_SixKeys.AssetID = UnitLastOBCReport.UnitID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="UnitVinDecoder"
    Expression ="vw_SixKeys.AssetID = UnitVinDecoder.UnitID"
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
        dbText "Name" ="FleetStudyNewVinsBase.Location"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.UnitYear"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.ReeferHours"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.State"
        dbInteger "ColumnOrder" ="26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.Axle"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.UnitMake"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.VINLenght"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetImporter_ExisitingUnitNums.UnitUnitNum"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.UnitModel"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.UnitNumber"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.Type"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.OBCUnitNumber"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.VIN"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.OdometerDate"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.SubType"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.InServiceDate"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.Plate"
        dbInteger "ColumnOrder" ="27"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.UnitRent"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.[Odometer ]"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.LeaseStartDate"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.LeaseEndDate"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.Unitcost"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.Division"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbInteger "ColumnWidth" ="2145"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVinsBase.[LeaseTerm ]"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescShortVer"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLastOBCReport.LogActivityDate"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DecodeStatus"
        dbInteger "ColumnWidth" ="975"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =119
    Top =25
    Right =1350
    Bottom =994
    Left =-1
    Top =-1
    Right =1213
    Bottom =692
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =109
        Top =81
        Right =253
        Bottom =448
        Top =0
        Name ="FleetStudyNewVinsBase"
        Name =""
    End
    Begin
        Left =302
        Top =53
        Right =446
        Bottom =197
        Top =0
        Name ="FleetImporter_ExisitingUnitNums"
        Name =""
    End
    Begin
        Left =499
        Top =103
        Right =633
        Bottom =451
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =672
        Top =129
        Right =816
        Bottom =273
        Top =0
        Name ="UnitLastOBCReport"
        Name =""
    End
    Begin
        Left =725
        Top =306
        Right =869
        Bottom =450
        Top =0
        Name ="UnitVinDecoder"
        Name =""
    End
End
