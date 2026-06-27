Operation =4
Option =0
Where ="(((SchGrp.SgrpDesc) Is Null Or (SchGrp.SgrpDesc)=\"\"))"
Begin InputTables
    Name ="SchGrp"
    Name ="UnitDescByGroupTemp"
End
Begin OutputColumns
    Name ="SchGrp.SgrpDesc"
    Expression ="[UnitDescByGroupTemp].[GroupDesc]"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="UnitDescByGroupTemp"
    Expression ="SchGrp.SGrpID = UnitDescByGroupTemp.GroupID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="SchGrp.SgrpDesc"
        dbInteger "ColumnWidth" ="7650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDescByGroup_B.GroupDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDescByGroup_B.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDescByGroupTemp.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDescByGroupTemp.GroupDesc"
        dbInteger "ColumnWidth" ="7170"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =171
    Top =68
    Right =1598
    Bottom =1066
    Left =-1
    Top =-1
    Right =1395
    Bottom =680
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =86
        Top =106
        Right =384
        Bottom =588
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =446
        Top =187
        Right =590
        Bottom =506
        Top =0
        Name ="UnitDescByGroupTemp"
        Name =""
    End
End
