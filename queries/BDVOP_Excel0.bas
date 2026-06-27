Operation =1
Option =0
Having ="(((IIf(IsNull([sgrpfaid]),[schfaid],[sgrpfaid])) Is Not Null) AND ((MstrLease.ML"
    "Orig)=\"fa\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
End
Begin OutputColumns
    Alias ="LFAID"
    Expression ="IIf(IsNull([sgrpfaid]),[schfaid],[sgrpfaid])"
    Expression ="MstrLease.MLOrig"
    Alias ="OrigEquipCost"
    Expression ="Sum(SchGrp.OrigEquipCost)"
    Alias ="UnitsInGroup"
    Expression ="Sum(SchGrp.UnitsInGroup)"
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
End
Begin Groups
    Expression ="IIf(IsNull([sgrpfaid]),[schfaid],[sgrpfaid])"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
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
        dbText "Name" ="LFAID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsInGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrigEquipCost"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1190
    Bottom =793
    Left =-1
    Top =-1
    Right =1158
    Bottom =238
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
        Right =768
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
