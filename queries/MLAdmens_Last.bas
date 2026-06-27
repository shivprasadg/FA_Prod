Operation =1
Option =0
Begin InputTables
    Name ="MLAmends"
End
Begin OutputColumns
    Alias ="ID"
    Expression ="Max(MLAmends.ID)"
    Alias ="MLAmdDt"
    Expression ="Max(MLAmends.MLAmdDt)"
    Alias ="MLAmtNO"
    Expression ="Max(MLAmends.MLAmtNO)"
    Expression ="MLAmends.MLRef"
End
Begin Groups
    Expression ="MLAmends.MLRef"
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
        dbText "Name" ="MLAmends.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmends.MLRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmends.MLAmdDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmends.MLAmdDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmends_1.MLAmtNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfMLAmdDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmdDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmends_1.MLAmdDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmends_1.MLAmdDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfMLRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmends.MLAmtNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfMLAmtNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLAmtNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =233
    Top =311
    Right =1197
    Bottom =962
    Left =-1
    Top =-1
    Right =946
    Bottom =374
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =186
        Bottom =246
        Top =0
        Name ="MLAmends"
        Name =""
    End
End
