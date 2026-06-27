Operation =4
Option =8
Where ="(((Schedule.SchFAID)=\"14-652\") AND ((SchGrp.SGrpNO)=\"A\") AND ((Schedule.SchN"
    "o)=\"2\"))"
Begin InputTables
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Name ="Units.UnitModel"
    Expression ="\"CA125\""
    Name ="Units.VendorREF"
    Expression ="226"
    Name ="Units.UnitFinalCost"
    Expression ="113340"
    Name ="Units.UNITRENT"
    Expression ="1358.21"
    Name ="Units.UNITPDRENT"
    Expression ="45.27"
    Name ="Units.UnitSalesResid"
    Expression ="32000"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNoUnits"
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
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITPDRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITsRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitGVW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpLType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.FirstExtRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSalesResid"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =27
    Top =67
    Right =1387
    Bottom =899
    Left =-1
    Top =-1
    Right =1328
    Bottom =324
    Left =0
    Top =393
    ColumnsShown =579
    Begin
        Left =147
        Top =-387
        Right =435
        Bottom =-99
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =499
        Top =-295
        Right =643
        Bottom =-54
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =696
        Top =-374
        Right =840
        Bottom =-1
        Top =0
        Name ="Units"
        Name =""
    End
End
