Operation =1
Option =0
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="MaxDelTRFq"
End
Begin OutputColumns
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Alias ="ProjBLCD"
    Expression ="IIf(IsNull([BLCD_Projected]),DateAdd(\"m\",1,[projIns])+(25-DatePart(\"d\",DateA"
        "dd(\"m\",1,[projIns]))),[BLCD_Projected])"
End
Begin Joins
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="MaxDelTRFq"
    Expression ="SchGrp.SGrpID = MaxDelTRFq.SGrpID"
    Flag =1
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
End
Begin Groups
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="IIf(IsNull([BLCD_Projected]),DateAdd(\"m\",1,[projIns])+(25-DatePart(\"d\",DateA"
        "dd(\"m\",1,[projIns]))),[BLCD_Projected])"
    GroupLevel =0
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
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjBLCD"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =80
    Right =1014
    Bottom =661
    Left =-1
    Top =-1
    Right =961
    Bottom =315
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =209
        Top =7
        Right =353
        Bottom =151
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =396
        Top =14
        Right =540
        Bottom =158
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =621
        Top =1
        Right =765
        Bottom =145
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =422
        Top =215
        Right =566
        Bottom =359
        Top =0
        Name ="MaxDelTRFq"
        Name =""
    End
End
