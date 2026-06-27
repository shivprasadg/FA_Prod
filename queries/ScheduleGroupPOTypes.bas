Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
End
Begin OutputColumns
    Alias ="ScheduleID"
    Expression ="SchGrp.SGrpSchID"
    Alias ="GroupID"
    Expression ="SchGrp.SGrpID"
    Alias ="POType"
    Expression ="IIf([CustPO1]=0 And [CustPO2]=0 And [CustPO3]=0 And [CustPO4]=0,\"FA\",\"Cust\")"
    Alias ="Ordered"
    Expression ="IIf(IsNull([PONum1]),\"TBD\",[PONum1])"
    Alias ="OrderedDate"
    Expression ="IIf([POType]=\"Cust\",\"Cust\",IIf(IsNull([POSentDt]),\"TBD\",[POSentDt]))"
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
        dbText "Name" ="POType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.CustPO1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ordered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PONum1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1272
    Bottom =1029
    Left =-1
    Top =-1
    Right =1240
    Bottom =655
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =193
        Top =85
        Right =393
        Bottom =673
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
