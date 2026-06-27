Operation =1
Option =0
Where ="(((Units.UnitVIN) Is Null))"
Begin InputTables
    Name ="FleetStudyImportBase"
    Name ="Units"
End
Begin OutputColumns
    Alias ="VINLenght"
    Expression ="IIf(Len(Trim([VIN]))<17,\"Short\",IIf(Len(Trim([VIN]))>17,\"Long\",\"Good\"))"
    Expression ="FleetStudyImportBase.UnitNumber"
    Expression ="FleetStudyImportBase.OBCUnitNumber"
    Alias ="VIN"
    Expression ="Trim([FleetStudyImportBase].[VIN])"
    Expression ="FleetStudyImportBase.UnitYear"
    Expression ="FleetStudyImportBase.UnitMake"
    Expression ="FleetStudyImportBase.UnitModel"
    Expression ="FleetStudyImportBase.Type"
    Expression ="FleetStudyImportBase.SubType"
    Expression ="FleetStudyImportBase.InServiceDate"
    Expression ="FleetStudyImportBase.Axle"
    Expression ="FleetStudyImportBase.UnitRent"
    Expression ="FleetStudyImportBase.LeaseStartDate"
    Expression ="FleetStudyImportBase.LeaseEndDate"
    Expression ="FleetStudyImportBase.Unitcost"
    Expression ="FleetStudyImportBase.Division"
    Expression ="FleetStudyImportBase.[LeaseTerm ]"
    Expression ="FleetStudyImportBase.[Odometer ]"
    Expression ="FleetStudyImportBase.OdometerDate"
    Expression ="FleetStudyImportBase.Location"
    Expression ="FleetStudyImportBase.State"
    Expression ="FleetStudyImportBase.Plate"
    Expression ="FleetStudyImportBase.ReeferHours"
End
Begin Joins
    LeftTable ="FleetStudyImportBase"
    RightTable ="Units"
    Expression ="FleetStudyImportBase.VIN = Units.UnitVIN"
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
        dbText "Name" ="FleetStudyImportBase.UnitYear"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.LeaseStartDate"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitNumber"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitMake"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.OdometerDate"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Location"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Axle"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Odometer"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Plate"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.OBCUnitNumber"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitModel"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitRent"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.LeaseEndDate"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Unitcost"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.State"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.LeaseTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.SubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.ReeferHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VINLenght"
        dbInteger "ColumnWidth" ="1380"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbInteger "ColumnWidth" ="2340"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.[Odometer ]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Division"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.[LeaseTerm ]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =73
    Top =52
    Right =1629
    Bottom =935
    Left =-1
    Top =-1
    Right =1538
    Bottom =506
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =75
        Top =-5
        Right =361
        Bottom =486
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
End
