Operation =1
Option =0
Where ="(((SchGrp.LoanBank) Is Not Null) And ((SchGrp.LoanDate) Is Null Or (SchGrp.LoanD"
    "ate)<forms!loanparmf!asofdt) And ((SchGrp.LoanAmount)>0.001) And ((SchGrp.LoanTy"
    "pe)=forms!loanparmf!loantype) And ((SchGrp.SgrpPayoffDt) Is Null Or (SchGrp.Sgrp"
    "PayoffDt)>forms!loanparmf!asofdt) And ((Clients.ClientGroupID)<> 18)) Or (((SchG"
    "rp.LoanAmount)>0.001) And ((SchGrp.LoanType)=forms!loanparmf!loantype) And ((Uni"
    "ts.UnitIFDate) Is Not Null Or (Units.UnitIFDate)<forms!LOANPARMF!ASOFDT) And ((C"
    "lients.ClientGroupID)<> 18))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="ChildPartSummaryQuery"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.LoanBank"
    Expression ="SchGrp.LoanDate"
    Expression ="SchGrp.LoanAmount"
    Expression ="SchGrp.LoanType"
    Expression ="SchGrp.SGrpExpDt"
    Expression ="SchGrp.SgrpPayoffDt"
    Expression ="Units.UnitVIN"
    Expression ="Units.unitstatus"
    Alias ="Cost"
    Expression ="[unitfinalcost]+IIf(IsNull([sumofchildfinalcost]),0,[sumofchildfinalcost])"
    Expression ="Units.UnitIFDate"
    Expression ="Units.UnitIFBank"
    Expression ="Units.UnitIFAmt"
    Expression ="Units.UnitPayOff"
    Alias ="ASOF DT"
    Expression ="[FORMS]![LOANPARMF]![ASOFDT]"
    Expression ="SchGrp.SgrpPayoffDt"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
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
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="SchGrp.SGrpNO"
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
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnWidth" ="1185"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpExpDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpPayoffDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPayOff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnWidth" ="1155"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ASOF DT"
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
    Bottom =270
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =81
        Top =27
        Right =225
        Bottom =171
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =250
        Top =36
        Right =394
        Bottom =180
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =474
        Top =42
        Right =618
        Bottom =186
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =666
        Top =12
        Right =830
        Bottom =410
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =906
        Top =56
        Right =1050
        Bottom =200
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1085
        Top =60
        Right =1229
        Bottom =204
        Top =0
        Name ="ChildPartSummaryQuery"
        Name =""
    End
End
