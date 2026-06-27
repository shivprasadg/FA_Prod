Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.MLID"
    Expression ="vw_SixKeys.MLNo"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.AssetID)"
    Alias ="Has Draw"
    Expression ="Max(IIf(Nz([vw_SixKeys].[DrawDownId],0)<>0,'Yes','No'))"
    Expression ="vw_SixKeys.ClientID"
End
Begin OrderBy
    Expression ="vw_SixKeys.MLNo"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.MLID"
    GroupLevel =0
    Expression ="vw_SixKeys.MLNo"
    GroupLevel =0
    Expression ="vw_SixKeys.ClientID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
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
        dbText "Name" ="vw_SixKeys.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HasDraws"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Has Draw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Has Draw]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1481
    Bottom =1021
    Left =-1
    Top =-1
    Right =1405
    Bottom =559
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =43
        Top =31
        Right =263
        Bottom =265
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
