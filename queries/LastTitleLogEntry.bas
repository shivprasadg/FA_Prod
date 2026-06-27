Operation =1
Option =0
Begin InputTables
    Name ="LastTitleEntry"
End
Begin OutputColumns
    Expression ="LastTitleEntry.AssetID"
    Expression ="LastTitleEntry.TitleLogId"
    Expression ="LastTitleEntry.LogDate"
    Expression ="LastTitleEntry.Action"
    Alias ="TitleNumber"
    Expression ="LastTitleEntry.LastOfTitleNumber"
    Alias ="Description"
    Expression ="LastTitleEntry.LastOfReason"
    Expression ="LastTitleEntry.ExpectedReturn"
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
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleNumber"
        dbInteger "ColumnWidth" ="1620"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry.LastOfTitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry.Action"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry.ExpectedReturn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry.LogDate"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="2175"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry.TitleLogId"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry.LastOfReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleEntry.AssetID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =162
    Top =307
    Right =1490
    Bottom =1338
    Left =-1
    Top =-1
    Right =1304
    Bottom =751
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =175
        Top =149
        Right =370
        Bottom =420
        Top =0
        Name ="LastTitleEntry"
        Name =""
    End
End
