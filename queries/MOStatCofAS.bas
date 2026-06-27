Operation =1
Option =0
Where ="(((Units.UnitCOASent) Is Not Null And (Units.UnitCOASent)>=[enter start date] An"
    "d (Units.UnitCOASent)<=[enter end date]))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Alias ="CountOfUnitVIN"
    Expression ="Count(Units.UnitVIN)"
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
        dbText "Name" ="CountOfUnitVIN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1248
    Bottom =772
    Left =-1
    Top =-1
    Right =1140
    Bottom =323
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
