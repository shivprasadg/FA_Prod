Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
End
Begin OutputColumns
    Alias ="SchId"
    Expression ="SchGrp.SGrpSchID"
    Alias ="GroupId"
    Expression ="SchGrp.SGrpID"
    Alias ="BLCD"
    Expression ="SchGrp.BLCD"
    Alias ="FeeTerm"
    Expression ="SchGrp.ScheduleGroupPaymentTerm"
    Alias ="FSLFeeX"
    Expression ="SchGrp.SGrpFSLPmt"
    Alias ="DateFeesEnd"
    Expression ="DateAdd(\"m\",Nz([ScheduleGroupPaymentTerm],0),[BLCD]-1)"
End
Begin Groups
    Expression ="SchGrp.SGrpSchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.ScheduleGroupPaymentTerm"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
    GroupLevel =0
    Expression ="DateAdd(\"m\",Nz([ScheduleGroupPaymentTerm],0),[BLCD]-1)"
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
        dbText "Name" ="DateFeesEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeeTerm"
        dbInteger "ColumnWidth" ="1260"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLFeeX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =161
    Top =82
    Right =1387
    Bottom =805
    Left =-1
    Top =-1
    Right =1208
    Bottom =321
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
