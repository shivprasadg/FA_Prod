Operation =3
Name ="TrailerMilesHours"
Option =0
Begin InputTables
    Name ="FleetStudyImportBase"
    Name ="Units"
End
Begin OutputColumns
    Name ="UnitID"
    Expression ="Units.UnitID"
    Name ="UnitNumber"
    Expression ="Units.UnitUnitNum"
    Name ="Miles"
    Expression ="FleetStudyImportBase.Odometer"
    Name ="ReadDate"
    Expression ="FleetStudyImportBase.OdometerDate"
    Alias ="Expr1"
    Name ="Hours"
    Expression ="FleetStudyImportBase.ReeferHours"
    Alias ="LogActivityDate"
    Name ="LogActivityDate"
    Expression ="DatePart('m',[OdometerDate]) & \"/1/\" & DatePart('yyyy',[OdometerDate])"
End
Begin Joins
    LeftTable ="FleetStudyImportBase"
    RightTable ="Units"
    Expression ="FleetStudyImportBase.VIN = Units.UnitVIN"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.Odometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.OdometerDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.ReeferHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =91
    Top =73
    Right =1681
    Bottom =1024
    Left =-1
    Top =-1
    Right =1558
    Bottom =616
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =65
        Top =25
        Right =250
        Bottom =377
        Top =0
        Name ="FleetStudyImportBase"
        Name =""
    End
    Begin
        Left =352
        Top =123
        Right =522
        Bottom =507
        Top =0
        Name ="Units"
        Name =""
    End
End
