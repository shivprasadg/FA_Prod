Operation =1
Option =0
Where ="(((Units.UnitTitleRecd)>=[enter start dt] And (Units.UnitTitleRecd)<=[enter end "
    "dt]) AND ((CLASS.TITLED)=Yes) AND ((MstrLease.mlorig)=\"fa\"))"
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="CLASS"
End
Begin OutputColumns
    Expression ="Units.UnitTitleRecd"
    Expression ="Units.UnitTASent"
    Expression ="Schedule.SchTitleToTrust"
    Expression ="Units.unitvin"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="Units.UnitAcceptDate"
    Expression ="CLASS.TITLED"
    Expression ="MstrLease.mlorig"
    Expression ="Units.UnitBuildDate"
    Expression ="Units.UnitPOA"
    Expression ="Units.UnitAddDt"
    Expression ="SchGrp.sgrpplb"
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
    LeftTable ="Units"
    RightTable ="CLASS"
    Expression ="Units.UnitType = CLASS.CLASS"
    Flag =2
End
Begin OrderBy
    Expression ="Units.UnitTitleRecd"
    Flag =0
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
        dbText "Name" ="Units.UnitTitleRecd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTASent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchTitleToTrust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CLASS.TITLED"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.mlorig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPOA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAddDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpplb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1248
    Bottom =772
    Left =-1
    Top =-1
    Right =1178
    Bottom =318
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Schedule"
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
        Left =610
        Top =25
        Right =797
        Bottom =297
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =843
        Top =163
        Right =987
        Bottom =307
        Top =0
        Name ="CLASS"
        Name =""
    End
End
