Operation =2
Name ="TitleLogTemp"
Option =0
Where ="(((LastTitleEntry_In.TitleNumber) Is Not Null) AND ((Units.UnitTitleNumber) Is N"
    "ull))"
Begin InputTables
    Name ="Units"
    Name ="LastTitleEntry_In"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="LastTitleEntry_In.TitleNumber"
    Alias ="DateLog"
    Expression ="LastTitleEntry_In.LastOfUnitTitleLogDate"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="LastTitleEntry_In"
    Expression ="Units.UnitID = LastTitleEntry_In.UnitRef"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleRecd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.TitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.Description"
        dbInteger "ColumnWidth" ="4575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry_In.LastOfUnitTitleLogDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateLog"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =172
    Top =131
    Right =1776
    Bottom =963
    Left =-1
    Top =-1
    Right =1572
    Bottom =515
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =29
        Top =106
        Right =233
        Bottom =430
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =306
        Top =147
        Right =546
        Bottom =400
        Top =0
        Name ="LastTitleEntry_In"
        Name =""
    End
End
