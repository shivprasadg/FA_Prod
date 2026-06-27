Operation =1
Option =0
Where ="(((vw_SixKeys.Status)=\"TR\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="ExitOptionsGroup_EOL"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.Status"
    Expression ="ExitOptionsGroup_EOL.EOLDate"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="ExitOptionsGroup_EOL"
    Expression ="vw_SixKeys.AssetID = ExitOptionsGroup_EOL.AssetID"
    Flag =2
End
Begin OrderBy
    Expression ="ExitOptionsGroup_EOL.EOLDate"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="[Query7].[EOLDate] DESC"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExitOptionsGroup_EOL.EOLDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =36
    Top =11
    Right =1595
    Bottom =1009
    Left =-1
    Top =-1
    Right =1527
    Bottom =714
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =443
        Bottom =455
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =511
        Top =140
        Right =840
        Bottom =502
        Top =0
        Name ="ExitOptionsGroup_EOL"
        Name =""
    End
End
