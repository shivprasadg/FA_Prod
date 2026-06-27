Operation =1
Option =0
Where ="(((Schedule.SchBnkClNO) Is Not Null))"
Having ="(((Schedule.SchID)=845))"
Begin InputTables
    Name ="Schedule"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="Schedule.SchID"
    Alias ="BankContract"
    Expression ="[SchBnkClNO] & \"-\" & [SgrpBnkNo]"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
End
Begin Groups
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="[SchBnkClNO] & \"-\" & [SgrpBnkNo]"
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
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankContract"
        dbInteger "ColumnWidth" ="4410"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =62
    Top =6
    Right =1462
    Bottom =838
    Left =-1
    Top =-1
    Right =1368
    Bottom =532
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =90
        Top =114
        Right =234
        Bottom =549
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =265
        Top =179
        Right =452
        Bottom =549
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
