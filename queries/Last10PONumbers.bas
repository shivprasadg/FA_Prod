Operation =1
Option =16
RowCount ="10"
Having ="(((SchGrp.PONum1) Is Not Null))"
Begin InputTables
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="SchGrp.PONum1"
    Expression ="SchGrp.SGrpFAID"
    Alias ="PoSeq"
    Expression ="CInt(splitstring([PONum1],\"-\",2))"
End
Begin OrderBy
    Expression ="SchGrp.SGrpFAID"
    Flag =1
End
Begin Groups
    Expression ="SchGrp.PONum1"
    GroupLevel =0
    Expression ="SchGrp.SGrpFAID"
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
        dbText "Name" ="SchGrp.PONum1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="POSEQ"
        dbInteger "ColumnWidth" ="3315"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="3315"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfSGrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1486
    Bottom =1026
    Left =-1
    Top =-1
    Right =1468
    Bottom =692
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =87
        Top =32
        Right =303
        Bottom =348
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
