Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
    Name ="Units"
    Name ="ChildPartSummaryQuery"
End
Begin OutputColumns
    Alias ="UTGrpRef"
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.SGrpID"
    Alias ="UnitCost"
    Expression ="Avg([unitfinalcost]+IIf(IsNull([SumOfchildFinalCost]),0,[sumofchildfinalcost]))"
    Alias ="SumOfUnitSynResid"
    Expression ="Sum(Units.UnitSynResid)"
    Alias ="AvgOfUNITRENT"
    Expression ="Avg(Units.UNITRENT)"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ChildPartSummaryQuery"
    Expression ="Units.UnitID = ChildPartSummaryQuery.UnitREF"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin Groups
    Expression ="SchGrp.SGrpID"
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
Begin
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTGrpRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgOfUNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitSynResid"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =164
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
    Begin
        Left =226
        Top =23
        Right =504
        Bottom =176
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =510
        Top =14
        Right =750
        Bottom =158
        Top =0
        Name ="ChildPartSummaryQuery"
        Name =""
    End
End
