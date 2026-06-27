Operation =1
Option =0
Begin InputTables
    Name ="UnitHistory"
End
Begin OutputColumns
    Expression ="UnitHistory.AssetId"
    Alias ="EventDate"
    Expression ="Max(UnitHistory.EventDate)"
    Alias ="UnitHistoryId"
    Expression ="Last(UnitHistory.UnitHistoryId)"
    Expression ="UnitHistory.EventGroup"
End
Begin Groups
    Expression ="UnitHistory.AssetId"
    GroupLevel =0
    Expression ="UnitHistory.EventGroup"
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
        dbText "Name" ="EventDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitHistoryId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitHistory.AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitHistory.EventGroup"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =6
    Top =18
    Right =1227
    Bottom =999
    Left =-1
    Top =-1
    Right =1197
    Bottom =647
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =213
        Top =72
        Right =543
        Bottom =357
        Top =0
        Name ="UnitHistory"
        Name =""
    End
End
