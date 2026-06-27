Operation =2
Name ="Axle_Update_Temp"
Option =0
Where ="(((Axle_Update_PartA.Axle)=\"TBD\") AND ((Axle_Update_PartB.NewAxel) Like \"*/A\""
    "))"
Begin InputTables
    Name ="Axle_Update_PartA"
    Name ="Axle_Update_PartB"
End
Begin OutputColumns
    Expression ="Axle_Update_PartA.UnitID"
    Expression ="Axle_Update_PartA.UnitVIN"
    Expression ="Axle_Update_PartA.Axle"
    Expression ="Axle_Update_PartA.UnitMake"
    Expression ="Axle_Update_PartA.UnitModel"
    Expression ="Axle_Update_PartA.UnitYr"
    Expression ="Axle_Update_PartB.NewAxel"
End
Begin Joins
    LeftTable ="Axle_Update_PartA"
    RightTable ="Axle_Update_PartB"
    Expression ="Axle_Update_PartA.VinPrefix = Axle_Update_PartB.VinPreFix_Decoder"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.AXLE_CONFIGURATION"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVinDecoder.DRIVE_LINE_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewAxel"
        dbInteger "ColumnWidth" ="3225"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Axle_Update_PartA.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Axle_Update_PartA.UnitVIN"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Axle_Update_PartA.Axle"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Axle_Update_PartA.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Axle_Update_PartA.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Axle_Update_PartA.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Axle_Update_PartB.NewAxel"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1520
    Bottom =825
    Left =-1
    Top =-1
    Right =1488
    Bottom =468
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =49
        Top =7
        Right =206
        Bottom =405
        Top =0
        Name ="Axle_Update_PartA"
        Name =""
    End
    Begin
        Left =291
        Top =140
        Right =435
        Bottom =284
        Top =0
        Name ="Axle_Update_PartB"
        Name =""
    End
End
