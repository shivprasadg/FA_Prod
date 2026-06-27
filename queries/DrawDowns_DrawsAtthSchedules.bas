Operation =1
Option =0
Where ="((Not (DDU.DrawDownId) Is Null))"
Begin InputTables
    Name ="DrawDowns_Units_All"
    Alias ="DDU"
End
Begin OutputColumns
    Expression ="DDU.DrawDownId"
    Alias ="Request#"
    Expression ="DDU.RequestNumber"
    Expression ="DDU.Draw"
    Expression ="DDU.DrawName"
    Alias ="Units"
    Expression ="Count(DDU.AssetID)"
    Alias ="UnitAmt"
    Expression ="Sum(DDU.UnitCost)"
    Expression ="DDU.SchID"
End
Begin OrderBy
    Expression ="DDU.Draw"
    Flag =0
End
Begin Groups
    Expression ="DDU.DrawDownId"
    GroupLevel =0
    Expression ="DDU.RequestNumber"
    GroupLevel =0
    Expression ="DDU.Draw"
    GroupLevel =0
    Expression ="DDU.DrawName"
    GroupLevel =0
    Expression ="DDU.SchID"
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
        dbText "Name" ="Units"
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
        dbText "Name" ="DDU.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DDU.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DDU.Draw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DDU.DrawName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1617
    Bottom =991
    Left =-1
    Top =-1
    Right =1593
    Bottom =548
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="DDU"
        Name =""
    End
End
