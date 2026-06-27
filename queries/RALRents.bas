Operation =1
Option =0
Begin InputTables
    Name ="UnitTotalOEC"
End
Begin OutputColumns
    Expression ="UnitTotalOEC.GroupID"
    Expression ="UnitTotalOEC.SyndRent"
    Alias ="Units"
    Expression ="Count(UnitTotalOEC.AssetID)"
    Alias ="RentLine"
    Expression ="Format(Nz([SyndRent],0),\"$#,###.##\") & \" per unit (x \" & Count([assetID]) & "
        "\" units)\""
End
Begin Groups
    Expression ="UnitTotalOEC.GroupID"
    GroupLevel =0
    Expression ="UnitTotalOEC.SyndRent"
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
        dbText "Name" ="RentLine"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfAssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1171
    Bottom =1017
    Left =-1
    Top =-1
    Right =1153
    Bottom =700
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =173
        Top =86
        Right =317
        Bottom =426
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
End
