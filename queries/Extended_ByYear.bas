Operation =1
Option =0
Where ="(((UnitExitOption.OptionTypeID)=3))"
Begin InputTables
    Name ="UnitExitOption"
End
Begin OutputColumns
    Expression ="UnitExitOption.UnitID"
    Expression ="UnitExitOption.OptionTypeDesc"
    Expression ="UnitExitOption.EndOrActualDate"
    Expression ="UnitExitOption.OrginalEOLDate"
    Alias ="EXTendYear"
    Expression ="CInt(Nz(DatePart(\"yyyy\",[DatePosted]),0))"
    Expression ="UnitExitOption.Rent"
    Expression ="UnitExitOption.TermMonths"
End
Begin OrderBy
    Expression ="UnitExitOption.UnitID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="UnitExitOption.EndOrActualDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1950"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="EXTendYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="UnitExitOption.Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OrginalEOLDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="UnitExitOption.TermMonths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="UnitExitOption.TerminateDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.DatePosted"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1324
    Bottom =969
    Left =-1
    Top =-1
    Right =1306
    Bottom =618
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =366
        Top =167
        Right =611
        Bottom =477
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
End
