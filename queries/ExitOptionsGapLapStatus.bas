Operation =1
Option =0
Begin InputTables
    Name ="EXITOptionsGAPReport"
End
Begin OutputColumns
    Expression ="EXITOptionsGAPReport.AssetID"
    Alias ="GapLapMsg"
    Expression ="IIf(IsNull([EarlyGapMsg]),Null,[EarlyGapMsg]) & \" \" & IIf(IsNull([OverLapMsg])"
        ",Null,[OverLapMsg])"
    Expression ="EXITOptionsGAPReport.Early"
    Expression ="EXITOptionsGAPReport.GAP"
    Expression ="EXITOptionsGAPReport.MonthsDiff"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([ExitOptionsGapLapStatus].[AssetID]=135)"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="EXITOptionsGAPReport.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GapLapMsg"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.Early"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.GAP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXITOptionsGAPReport.MonthsDiff"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =27
    Top =70
    Right =975
    Bottom =979
    Left =-1
    Top =-1
    Right =916
    Bottom =421
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =267
        Bottom =458
        Top =0
        Name ="EXITOptionsGAPReport"
        Name =""
    End
End
