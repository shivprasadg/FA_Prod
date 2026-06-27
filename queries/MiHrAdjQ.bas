Operation =1
Option =0
Where ="(((MstrLease.MLNo) Like \"costco*\") And ((Schedule.SchNo)=\"1 mira loma\") And "
    "((Mileage.MHDate)>=\"#8/1/2012#\" And (Mileage.MHDate)<=\"#9/1/2012#\") And ((Ms"
    "trLease.MLOrig)=\"fa\"))"
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Mileage"
    Name ="Units"
End
Begin OutputColumns
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.unitinserv"
    Expression ="Units.unitunitnum"
    Alias ="Expr1"
    Expression ="Mileage.MHDate"
    Alias ="Expr2"
    Expression ="Mileage.Mileage"
    Expression ="Mileage.Hours"
End
Begin Joins
    LeftTable ="Mileage"
    RightTable ="Units"
    Expression ="Mileage.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="Units.unitunitnum"
    Flag =0
    Expression ="Mileage.MHDate"
    Flag =0
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
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.MHDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.Mileage"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="0"
    End
    Begin
        dbText "Name" ="Mileage.Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitinserv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =381
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =5
        Top =24
        Right =149
        Bottom =168
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =203
        Top =23
        Right =347
        Bottom =167
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =373
        Top =65
        Right =517
        Bottom =209
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =728
        Top =55
        Right =872
        Bottom =199
        Top =0
        Name ="Mileage"
        Name =""
    End
    Begin
        Left =551
        Top =42
        Right =695
        Bottom =186
        Top =0
        Name ="Units"
        Name =""
    End
End
