Operation =1
Option =8
Begin InputTables
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Schedule.SchFAID"
    Expression ="SchGrp.SGrpNO"
    Expression ="Schedule.SchNo"
    Expression ="Units.UnitStatus"
    Alias ="AssetID"
    Expression ="Units.unitID"
    Alias ="UnitDesc"
    Expression ="[UnitYr] & ' ' & [UnitMake] & ' ' & [UnitSubType] & ' ' & [UnitType] & IIf(Len(["
        "UnitConfiguration])>0,\" \" & [UnitConfiguration],\"\") & IIf(Len([UnitDescripto"
        "r])>0,\" \" & [UnitDescriptor],\"\") & \" (\" & [UnitStatus] & \")\""
    Expression ="Units.UnitYr"
    Expression ="Units.unitmake"
    Expression ="Units.unittype"
    Expression ="Units.UnitVIN"
    Alias ="UnitNum"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitVINLocked"
    Expression ="SchGrp.SGrpID"
    Expression ="Schedule.SchID"
    Expression ="Units.UnitConfiguration"
    Expression ="Units.UnitDescriptor"
    Expression ="Units.DrawDownID"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFAID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="AssetID"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVINLocked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4935"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitConfiguration"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDescriptor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrawDownID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =201
    Top =94
    Right =1561
    Bottom =921
    Left =-1
    Top =-1
    Right =1336
    Bottom =310
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =147
        Top =6
        Right =435
        Bottom =294
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =499
        Top =98
        Right =643
        Bottom =339
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =696
        Top =19
        Right =840
        Bottom =392
        Top =0
        Name ="Units"
        Name =""
    End
End
