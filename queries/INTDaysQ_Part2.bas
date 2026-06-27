Operation =1
Option =0
Begin InputTables
    Name ="INTDaysQ_Part1"
    Name ="BillStEndq"
    Name ="BillUnitq"
End
Begin OutputColumns
    Expression ="INTDaysQ_Part1.*"
    Expression ="BillStEndq.Billstart"
    Expression ="BillStEndq.BillEndDt"
    Expression ="BillStEndq.SlvEnd"
    Expression ="BillStEndq.PDEND"
    Expression ="BillUnitq.UnitID"
    Alias ="SlvDays"
    Expression ="IIf([billstart]>[schcrexp],0,IIf([sgrpnoslv]=No,0,IIf([ubstart]>[slvend],0,IIf(["
        "ubstart]>[billenddt],0,DateDiff(\"d\",IIf([ubstart]>[billstart],[ubstart],[bills"
        "tart]),IIf([billenddt]<[slvend],[billenddt],[slvend])+1)))))"
    Alias ="UBSlvS"
    Expression ="IIf(IsNull([billstart]),[ubstart],IIf([ubstart]>[billstart],[ubstart],[billstart"
        "]))"
    Alias ="ubslvend"
    Expression ="IIf([billenddt]<[slvend],[billenddt],[slvend])"
End
Begin Joins
    LeftTable ="INTDaysQ_Part1"
    RightTable ="BillStEndq"
    Expression ="INTDaysQ_Part1.MLNo = BillStEndq.MLNo"
    Flag =1
    LeftTable ="INTDaysQ_Part1"
    RightTable ="BillStEndq"
    Expression ="INTDaysQ_Part1.SchNo = BillStEndq.SchNo"
    Flag =1
    LeftTable ="INTDaysQ_Part1"
    RightTable ="BillUnitq"
    Expression ="INTDaysQ_Part1.SchNo = BillUnitq.SchNo"
    Flag =1
    LeftTable ="INTDaysQ_Part1"
    RightTable ="BillUnitq"
    Expression ="INTDaysQ_Part1.MLNo = BillUnitq.MLNo"
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
        dbText "Name" ="INTDaysQ_Part1.MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INTDaysQ_Part1.Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INTDaysQ_Part1.Schedule.SchCRExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INTDaysQ_Part1.SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INTDaysQ_Part1.SchGrp.SGrpNoSlv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.PDEND"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.BillEndDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.Billstart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.SlvEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INTDaysQ_Part1.MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ubslvend"
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
        dbText "Name" ="BillUnitq.UnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =131
    Top =23
    Right =1436
    Bottom =792
    Left =-1
    Top =-1
    Right =1273
    Bottom =350
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =88
        Top =51
        Right =228
        Bottom =346
        Top =0
        Name ="INTDaysQ_Part1"
        Name =""
    End
    Begin
        Left =367
        Top =141
        Right =519
        Bottom =382
        Top =0
        Name ="BillStEndq"
        Name =""
    End
    Begin
        Left =614
        Top =157
        Right =758
        Bottom =368
        Top =0
        Name ="BillUnitq"
        Name =""
    End
End
