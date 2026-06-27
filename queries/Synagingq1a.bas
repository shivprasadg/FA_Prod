Operation =1
Option =0
Where ="(((Schedule.SchExSyn)=No) AND ((Units.unitacceptdate) Is Not Null) AND ((SchGrp."
    "LoanType) Is Null Or (SchGrp.LoanType)<>\"perm\") AND ((MstrLease.MLOrig)=\"fa\""
    ") AND ((Units.unitstatus)=\"a\"))"
Having ="(((Max(Units.unitinserv)) Is Not Null) AND ((Clients.ClientShNm) Not Like \"demo"
    "*\") AND ((SchGrp.AssignmentDate) Is Null) AND ((SchGrp.FundDate) Is Null))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Schedule.SchNo"
    Alias ="MaxOfunitinserv"
    Expression ="Max(Units.unitinserv)"
    Expression ="Schedule.SchID"
    Alias ="INserVIN"
    Expression ="Count(Units.UnitVIN)"
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
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
End
Begin Groups
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.AssignmentDate"
    GroupLevel =0
    Expression ="SchGrp.UnitsInGroup"
    GroupLevel =0
    Expression ="SchGrp.FundDate"
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
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="3570"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfunitinserv"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="INservin"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =925
    Bottom =753
    Left =-1
    Top =-1
    Right =893
    Bottom =233
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =859
        Bottom =261
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =912
        Top =24
        Right =1056
        Bottom =168
        Top =0
        Name ="Units"
        Name =""
    End
End
