Operation =1
Option =0
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="schgrpcntq"
    Name ="SchGrp"
    Name ="Units"
    Name ="ChildPartSummaryQuery"
End
Begin OutputColumns
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="Schedule.SchShortDesc"
    Expression ="SchGrp.SGrpNO"
    Alias ="Unit No"
    Expression ="Units.UNITUNITNUM"
    Alias ="VIN"
    Expression ="Units.UNITVIN"
    Alias ="Description"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType]"
    Alias ="In-Service Date"
    Expression ="Units.UNITINSERV"
    Alias ="Lessor's Cost(per Unit)"
    Expression ="[unitfinalcost]+IIf(IsNull([sumofchildfinalcost]),0,[sumofchildfinalcost])"
    Alias ="Item #"
    Expression ="1"
    Alias ="vin4"
    Expression ="Right([unitvin],6)"
    Alias ="GrpCnt"
    Expression ="schgrpcntq.CountOfSGrpID"
    Expression ="SchGrp.SgrpDesc"
    Expression ="Schedule.SchID"
    Alias ="ClientID"
    Expression ="MstrLease.Client"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="schgrpcntq"
    Expression ="Schedule.SchID = schgrpcntq.SchID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="ChildPartSummaryQuery"
    Expression ="Units.UnitID = ChildPartSummaryQuery.UnitREF"
    Flag =2
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
Begin
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchShortDesc"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="2625"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="In-Service Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lessor's Cost(per Unit)"
        dbText "Format" ="$#,##0.00;($#,##0.00)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item #"
        dbText "Format" ="Fixed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrpCnt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
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
    Bottom =227
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="schgrpcntq"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1008
        Top =12
        Right =1152
        Bottom =156
        Top =0
        Name ="ChildPartSummaryQuery"
        Name =""
    End
End
