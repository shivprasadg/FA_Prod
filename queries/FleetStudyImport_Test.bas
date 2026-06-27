Operation =1
Option =0
Where ="(((FleetImporter_ExisitingUnitNums.UnitUnitNum) Is Null))"
Begin InputTables
    Name ="FleetStudyNewVins"
    Name ="FleetImporter_ExisitingUnitNums"
    Name ="FleetImporter_Locations"
End
Begin OutputColumns
    Expression ="FleetStudyNewVins.VINLenght"
    Expression ="FleetStudyNewVins.UnitNumber"
    Expression ="FleetStudyNewVins.OBCUnitNumber"
    Expression ="FleetStudyNewVins.VIN"
    Expression ="FleetStudyNewVins.UnitYear"
    Expression ="FleetStudyNewVins.UnitMake"
    Alias ="UnitModel"
    Expression ="FleetStudyNewVins.UnitModel"
    Expression ="FleetStudyNewVins.Type"
    Expression ="FleetStudyNewVins.SubType"
    Expression ="FleetStudyNewVins.InServiceDate"
    Alias ="Axle"
    Expression ="Nz([FleetStudyNewVins].[Axle],\"TBD\")"
    Expression ="FleetStudyNewVins.UnitRent"
    Expression ="FleetStudyNewVins.LeaseStartDate"
    Expression ="FleetStudyNewVins.LeaseEndDate"
    Expression ="FleetStudyNewVins.Unitcost"
    Expression ="FleetStudyNewVins.[LeaseTerm ]"
    Expression ="FleetStudyNewVins.Odometer"
    Expression ="FleetStudyNewVins.OdometerDate"
    Expression ="FleetStudyNewVins.Location"
    Expression ="FleetStudyNewVins.State"
    Expression ="FleetStudyNewVins.Plate"
    Expression ="FleetStudyNewVins.ReeferHours"
    Alias ="LocID"
    Expression ="FleetImporter_Locations.LocationID"
    Expression ="FleetImporter_Locations.ClientREF"
End
Begin Joins
    LeftTable ="FleetStudyNewVins"
    RightTable ="FleetImporter_ExisitingUnitNums"
    Expression ="FleetStudyNewVins.UnitNumber = FleetImporter_ExisitingUnitNums.UnitUnitNum"
    Flag =2
    LeftTable ="FleetStudyNewVins"
    RightTable ="FleetImporter_Locations"
    Expression ="FleetStudyNewVins.Location = FleetImporter_Locations.LocationName"
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
        dbText "Name" ="FleetStudyNewVins.VINLenght"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.UnitYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.UnitRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.SubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.ReeferHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.UnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.OBCUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.LeaseStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.OdometerDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.LeaseEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.Unitcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.Plate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetImporter_Locations.ClientREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.[LeaseTerm ]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyNewVins.Odometer"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-121
    Top =51
    Right =1383
    Bottom =1036
    Left =-1
    Top =-1
    Right =1486
    Bottom =561
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =215
        Bottom =398
        Top =0
        Name ="FleetStudyNewVins"
        Name =""
    End
    Begin
        Left =474
        Top =75
        Right =723
        Bottom =229
        Top =0
        Name ="FleetImporter_ExisitingUnitNums"
        Name =""
    End
    Begin
        Left =266
        Top =372
        Right =410
        Bottom =728
        Top =0
        Name ="FleetImporter_Locations"
        Name =""
    End
End
