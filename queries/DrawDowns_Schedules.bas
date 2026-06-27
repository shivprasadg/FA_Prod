Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="SchGrp"
    Name ="Schedule"
    Name ="UnitTotalOEC"
End
Begin OutputColumns
    Alias ="SchID"
    Expression ="SchGrp.SGrpSchID"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Alias ="OpenAmt"
    Expression ="Format(Sum(IIf(Nz([Units].[DrawDownId],0)=0,[UnitTotalCost],0)),\"$#,###\")"
    Alias ="DrawnAmt"
    Expression ="Format(Sum(IIf(Nz([Units].[DrawDownId],0)=0,0,[unitIFAmt])),\"$#,###\")"
    Alias ="MLID"
    Expression ="Schedule.SchMLRef"
    Alias ="OpenUnits"
    Expression ="Sum(IIf(Nz([Units].[DrawDownId],0)=0,1,0))"
    Alias ="DrawnUnits"
    Expression ="Sum(IIf(Nz([Units].[DrawDownId],0)=0,0,1))"
    Alias ="IFBank"
    Expression ="Min(Units.UnitIfBank)"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Schedule"
    Expression ="SchGrp.SGrpSchID = Schedule.SchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="UnitTotalOEC"
    Expression ="Units.UnitID = UnitTotalOEC.AssetID"
    Flag =1
End
Begin Groups
    Expression ="SchGrp.SGrpSchID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="Schedule.SchMLRef"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="OpenAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawnAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IFBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpenUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawnUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OEC_Rents_LRF_Cost.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OEC_Rents_LRF_Cost.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.Schedule"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.UnitTotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchMLRef"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1481
    Bottom =1021
    Left =-1
    Top =-1
    Right =1413
    Bottom =423
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =158
        Top =182
        Right =302
        Bottom =326
        Top =0
        Name ="Units"
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
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =431
        Top =168
        Right =575
        Bottom =412
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
End
