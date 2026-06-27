Operation =1
Option =0
Where ="(((Trim(FS.VIN))<>\"\") And ((U.UnitVIN) Is Null))"
Begin InputTables
    Name ="FleetStudyImport_Work"
    Alias ="FS"
    Name ="SELECT Units.UnitID, Units.UnitVIN, Units.ClientID, Units.ClientGroupID FROM Uni"
        "ts WHERE Units.ClientGroupID=[Forms]![FleetStudy_Importer]![ClientGroupId]"
    Alias ="U"
End
Begin OutputColumns
    Alias ="VINLenght"
    Expression ="IIf(Len(Trim([VIN]))<17,\"Short\",IIf(Len(Trim([VIN]))>17,\"Long\",\"Good\"))"
    Expression ="FS.UnitNumber"
    Expression ="FS.OBCUnitNumber"
    Alias ="VIN"
    Expression ="Trim(FS.[VIN])"
    Expression ="FS.UnitYear"
    Expression ="FS.UnitMake"
    Expression ="FS.UnitModel"
    Expression ="FS.Type"
    Expression ="FS.SubType"
    Expression ="FS.InServiceDate"
    Expression ="FS.Axle"
    Expression ="FS.UnitRent"
    Expression ="FS.LeaseStartDate"
    Expression ="FS.LeaseEndDate"
    Expression ="FS.Unitcost"
    Expression ="FS.LeaseTerm"
    Alias ="Odometer"
    Expression ="Nz([FS].[Odometer],0)"
    Expression ="FS.OdometerDate"
    Expression ="FS.Ownership"
    Expression ="FS.Location"
    Expression ="FS.State"
    Expression ="FS.Plate"
    Expression ="FS.ReeferHours"
    Expression ="FS.Division"
    Expression ="FS.DivisionID"
    Expression ="FS.Engine"
    Expression ="FS.Trans"
End
Begin Joins
    LeftTable ="FS"
    RightTable ="U"
    Expression ="FS.VIN = U.UnitVIN"
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
        dbText "Name" ="Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.OdometerDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.Plate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.LeaseStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.UnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.OBCUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.UnitYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.Division"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.SubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.UnitRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.LeaseEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.Unitcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.ReeferHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.DivisionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.Engine"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.Trans"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.LeaseTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FS.Ownership"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =35
    Top =17
    Right =1591
    Bottom =900
    Left =-1
    Top =-1
    Right =1532
    Bottom =400
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =393
        Top =0
        Name ="FS"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="U"
        Name =""
    End
End
