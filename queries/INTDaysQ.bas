Operation =1
Option =0
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="BillStEndq"
    Name ="BillUnitq"
End
Begin OutputColumns
    Expression ="MstrLease.MLID"
    Expression ="BillStEndq.BillID"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpID"
    Expression ="BillUnitq.UnitID"
    Alias ="SlvDays"
    Expression ="IIf([billstart]>[schcrexp],0,IIf([sgrpnoslv]=No,0,IIf([ubstart]>[slvend],0,IIf(["
        "ubstart]>[billenddt],0,DateDiff(\"d\",IIf([ubstart]>[billstart],[ubstart],[bills"
        "tart]),IIf([billenddt]<[slvend],[billenddt],[slvend])+1)))))"
    Alias ="UBSlvS"
    Expression ="IIf([ubstart]>[billstart],[ubstart],[billstart])"
    Alias ="ubslvend"
    Expression ="IIf([billenddt]<[slvend],[billenddt],[slvend])"
    Expression ="BillUnitq.ubstart"
    Expression ="BillStEndq.Billstart"
    Expression ="BillStEndq.slvend"
    Expression ="BillStEndq.billenddt"
End
Begin Joins
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="BillStEndq"
    Expression ="SchGrp.SGrpID = BillStEndq.SGrpID"
    Flag =1
    LeftTable ="BillStEndq"
    RightTable ="BillUnitq"
    Expression ="BillStEndq.SGrpID = BillUnitq.SGrpID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
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
        dbText "Name" ="BillUnitq.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SlvDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UBSlvS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ubslvend"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillUnitq.ubstart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.slvend"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.billenddt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.Billstart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.BillID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =128
    Top =69
    Right =1323
    Bottom =812
    Left =-1
    Top =-1
    Right =1163
    Bottom =159
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
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="BillStEndq"
        Name =""
    End
    Begin
        Left =816
        Top =6
        Right =960
        Bottom =150
        Top =0
        Name ="BillUnitq"
        Name =""
    End
End
