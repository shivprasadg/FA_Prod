Operation =1
Option =0
Where ="(((Units.GroupREF)=[forms]![groupform]![sgrpid]))"
Begin InputTables
    Name ="Units"
    Name ="Rmkt"
End
Begin OutputColumns
    Expression ="Units.UNITMAKE"
    Expression ="Units.UNITVIN"
    Expression ="Units.UNITMODEL"
    Expression ="Units.UNITUNITNUM"
    Expression ="Units.UNITTYPE"
    Expression ="Rmkt.[SUF Sent]"
    Expression ="Units.GroupREF"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Rmkt.[SUF Sent]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITUNITNUM"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1785"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UNITVIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2070"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UNITMAKE"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1635"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UNITMODEL"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Units.UNITTYPE"
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
    Bottom =155
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
