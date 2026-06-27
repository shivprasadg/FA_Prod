Operation =4
Option =0
Begin InputTables
    Name ="Units"
    Name ="StagingVendorTracking_Lastest"
End
Begin OutputColumns
    Name ="Units.UnitVIN"
    Expression ="[StagingVendorTracking_Lastest]![VIN]"
    Name ="Units.UnitEstDelDate"
    Expression ="[StagingVendorTracking_Lastest]![UnitEstDelDate]"
    Name ="Units.UnitActualDelDate"
    Expression ="[StagingVendorTracking_Lastest]![UnitActualDelDate]"
    Name ="Units.UnitBuildDate"
    Expression ="[UnitActualDelDate]![UnitBuildDate]"
    Name ="Units.UnitPilot"
    Expression ="[UnitBuildDate]![IsPilot]"
    Name ="Units.UnitEngineSerial"
    Expression ="[UnitBuildDate]![EngineSer]"
    Name ="Units.UnitTransSerial"
    Expression ="[UnitBuildDate]![TransSer]"
    Name ="Units.UnitFASerial"
    Expression ="[UnitBuildDate]![RAxleFWDSer]"
    Name ="Units.UnitRAFSerial"
    Expression ="[UnitBuildDate]![RAxleREARSer]"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="StagingVendorTracking_Lastest"
    Expression ="Units.UnitID = StagingVendorTracking_Lastest.UnitId"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2205"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitBuildDate"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPilot"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRARSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFASerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRAFSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.RAxleFWDSer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.ImportDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =265
    Top =-2
    Right =1877
    Bottom =830
    Left =-1
    Top =-1
    Right =1580
    Bottom =481
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="StagingVendorTracking_Lastest"
        Name =""
    End
End
