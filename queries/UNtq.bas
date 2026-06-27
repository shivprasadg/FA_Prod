Operation =1
Option =0
Where ="(((UnitNotes.UnitRef)=[forms]![UnitsForm]![unitid]))"
Begin InputTables
    Name ="UnitNotes"
End
Begin OutputColumns
    Expression ="UnitNotes.*"
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
        dbText "Name" ="UnitNotes.UnitNtid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNotes.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNotes.UnitNtdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNotes.UnitNt"
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
    Bottom =373
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="UnitNotes"
        Name =""
    End
End
