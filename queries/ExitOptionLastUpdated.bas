Operation =1
Option =0
Begin InputTables
    Name ="UnitExitOption"
End
Begin OutputColumns
    Expression ="UnitExitOption.UnitID"
    Alias ="MaxOfDatePosted"
    Expression ="Max(UnitExitOption.DatePosted)"
    Expression ="UnitExitOption.LeaseOptionID"
End
Begin Groups
    Expression ="UnitExitOption.UnitID"
    GroupLevel =0
    Expression ="UnitExitOption.LeaseOptionID"
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
        dbText "Name" ="UnitExitOption.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfDatePosted"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.LeaseOptionID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1143
    Bottom =1009
    Left =-1
    Top =-1
    Right =1111
    Bottom =668
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =309
        Top =120
        Right =823
        Bottom =513
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
End
