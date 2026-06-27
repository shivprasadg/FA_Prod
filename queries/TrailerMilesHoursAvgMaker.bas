Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="FleetStudyImportBase"
    Name ="TrailerLastMaxMilesHours"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitUnitNum"
    Expression ="FleetStudyImportBase.VIN"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitType"
    Expression ="Units.UnitInServ"
    Alias ="ReadDate"
    Expression ="FleetStudyImportBase.OdometerDate"
    Alias ="ReeferHours"
    Expression ="[FleetStudyImportBase].[ReeferHours]-[HoursLTD]"
    Alias ="TrailerMiles"
    Expression ="[FleetStudyImportBase].[Odometer]-[MilesLTD]"
    Alias ="Months"
    Expression ="DateDiff(\"m\",[UnitInServ],[OdometerDate])"
    Alias ="AvgMilesPM"
    Expression ="Round([Odometer]/[Months],1)"
    Alias ="AvgHoursPM"
    Expression ="Round([ReeferHours]/[months],1)"
    Alias ="LogActivityDate"
    Expression ="DatePart('m',[OdometerDate]) & \"/1/\" & DatePart('yyyy',[OdometerDate])"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="FleetStudyImportBase"
    Expression ="Units.UnitVIN = FleetStudyImportBase.VIN"
    Flag =1
    LeftTable ="Units"
    RightTable ="TrailerLastMaxMilesHours"
    Expression ="Units.UnitID = TrailerLastMaxMilesHours.UnitID"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReeferHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="months"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgMilesPM"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgHoursPM"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMiles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerLastMaxMilesHours.MaxOfMilesLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =134
    Top =60
    Right =1071
    Bottom =1011
    Left =-1
    Top =-1
    Right =905
    Bottom =497
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =67
        Top =15
        Right =269
        Bottom =428
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =299
        Top =47
        Right =570
        Bottom =496
        Top =0
        Name ="FleetStudyImportBase"
        Name =""
    End
    Begin
        Left =675
        Top =36
        Right =819
        Bottom =180
        Top =0
        Name ="TrailerLastMaxMilesHours"
        Name =""
    End
End
