Operation =1
Option =0
Where ="(((Schedule.SchExSyn)=No) AND ((SchGrp.LoanType) Is Null Or (SchGrp.LoanType)<>\""
    "perm\") AND ((MstrLease.MLOrig)=\"fa\") AND ((SchGrp.FundDate) Is Null) AND ((Un"
    "its.unitstatus)=\"a\"))"
Having ="(((Clients.ClientShNm) Not Like \"demo*\") AND ((SchGrp.AssignmentDate) Is Null)"
    " AND ((Max(Units.unitacceptdate)) Is Not Null))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.AssignmentDate"
    Expression ="SchGrp.UnitsInGroup"
    Alias ="CountOfunitvin"
    Expression ="Count(Units.unitvin)"
    Alias ="MaxOfunitacceptdate"
    Expression ="Max(Units.unitacceptdate)"
    Expression ="Schedule.SchID"
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
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.AssignmentDate"
    GroupLevel =0
    Expression ="SchGrp.UnitsInGroup"
    GroupLevel =0
    Expression ="Schedule.SchID"
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
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="2115"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnWidth" ="1755"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="1935"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpblcd"
        dbInteger "ColumnOrder" ="8"
        dbText "Format" ="mm/dd/yy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpBA"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfunitvin"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpnounits"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfunitacceptdate"
        dbInteger "ColumnOrder" ="10"
        dbText "Format" ="mm/dd/yy"
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
    Right =925
    Bottom =753
    Left =-1
    Top =-1
    Right =893
    Bottom =301
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
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
