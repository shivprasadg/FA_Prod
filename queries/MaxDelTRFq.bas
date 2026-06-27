Operation =1
Option =0
Having ="(((MstrLease.MLOrig)=\"fa\" Or (MstrLease.MLOrig)=\"other\") AND ((Max(IIf(IsNul"
    "l([unitpdstartdate]),IIf(IsNull([unitacceptdate]),IIf(IsNull([unitinserv]),IIf(I"
    "sNull([unitestdeldate]),[unitbuilddate],[unitestdeldate]),[unitinserv]),[unitacc"
    "eptdate]),[unitpdstartdate]))) Is Not Null) AND ((Avg(IIf(IsNull([unitpdstartdat"
    "e]),IIf(IsNull([unitacceptdate]),IIf(IsNull([unitinserv]),IIf(IsNull([unitestdel"
    "date]),[unitbuilddate],[unitestdeldate]),[unitinserv]),[unitacceptdate]),[unitpd"
    "startdate]))) Is Not Null))"
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLOrig"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.SGrpID"
    Alias ="ProjIns"
    Expression ="Max(IIf(IsNull([unitpdstartdate]),IIf(IsNull([unitacceptdate]),IIf(IsNull([uniti"
        "nserv]),IIf(IsNull([unitestdeldate]),[unitbuilddate],[unitestdeldate]),[unitinse"
        "rv]),[unitacceptdate]),[unitpdstartdate]))"
    Alias ="AVGIns"
    Expression ="Avg(IIf(IsNull([unitpdstartdate]),IIf(IsNull([unitacceptdate]),IIf(IsNull([uniti"
        "nserv]),IIf(IsNull([unitestdeldate]),[unitbuilddate],[unitestdeldate]),[unitinse"
        "rv]),[unitacceptdate]),[unitpdstartdate]))"
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
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
End
Begin Groups
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
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
dbMemo "Filter" ="([MaxDelTRFq].[MLNo] In (\"AMI2011\",\"Daimler\"))"
Begin
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ProjIns"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AVGIns"
        dbText "Format" ="mm/dd/yy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =152
    Top =224
    Right =1189
    Bottom =793
    Left =-1
    Top =-1
    Right =1005
    Bottom =116
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =24
        Top =3
        Right =168
        Bottom =147
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =225
        Top =22
        Right =369
        Bottom =166
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =392
        Top =22
        Right =536
        Bottom =166
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =595
        Top =20
        Right =834
        Bottom =164
        Top =0
        Name ="Units"
        Name =""
    End
End
