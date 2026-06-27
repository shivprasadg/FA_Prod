Operation =1
Option =0
Begin InputTables
    Name ="DrawDowns_Units_Drawn"
End
Begin OutputColumns
    Expression ="DrawDowns_Units_Drawn.DrawDownId"
    Alias ="Request#"
    Expression ="DrawDowns_Units_Drawn.RequestNumber"
    Expression ="DrawDowns_Units_Drawn.Draw"
    Expression ="DrawDowns_Units_Drawn.DrawName"
    Alias ="Units"
    Expression ="Count(DrawDowns_Units_Drawn.AssetID)"
    Alias ="UnitAmt"
    Expression ="Format(Sum([UnitCost]),\"$#,###\")"
    Expression ="DrawDowns_Units_Drawn.Invoice"
    Expression ="DrawDowns_Units_Drawn.SGrpID"
End
Begin OrderBy
    Expression ="DrawDowns_Units_Drawn.Draw"
    Flag =0
End
Begin Groups
    Expression ="DrawDowns_Units_Drawn.DrawDownId"
    GroupLevel =0
    Expression ="DrawDowns_Units_Drawn.RequestNumber"
    GroupLevel =0
    Expression ="DrawDowns_Units_Drawn.Draw"
    GroupLevel =0
    Expression ="DrawDowns_Units_Drawn.DrawName"
    GroupLevel =0
    Expression ="DrawDowns_Units_Drawn.Invoice"
    GroupLevel =0
    Expression ="DrawDowns_Units_Drawn.SGrpID"
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
        dbText "Name" ="DrawDowns_Units_Drawn.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns_Units_Drawn.Draw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns_Units_Drawn.Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Request#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns_Units_Drawn.DrawName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DrawDowns_Units_Drawn.SGrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =294
    Top =54
    Right =1911
    Bottom =1005
    Left =-1
    Top =-1
    Right =1593
    Bottom =599
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =70
        Top =64
        Right =214
        Bottom =531
        Top =0
        Name ="DrawDowns_Units_Drawn"
        Name =""
    End
End
