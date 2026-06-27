Operation =1
Option =0
Where ="(((vw_SixKeys.SubType)<>\"Auto Hauler\"))"
Having ="(((vw_SixKeys.MLOrig)=\"FA\") AND ((IIf([UnitsInGroup]<>Count([AssetID]),\"Y\",N"
    "ull))=\"Y\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.MLOrig"
    Alias ="ActualUnits"
    Expression ="Count(vw_SixKeys.AssetID)"
    Expression ="SchGrp.UnitsInGroup"
    Alias ="Problem"
    Expression ="IIf([UnitsInGroup]<>Count([AssetID]),\"Y\",Null)"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
End
Begin Groups
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.MLNo"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
    GroupLevel =0
    Expression ="vw_SixKeys.UnitGroup"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="SchGrp.UnitsInGroup"
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
        dbText "Name" ="Problem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.UnitsInGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualUnits"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-39
    Top =58
    Right =1394
    Bottom =1056
    Left =-1
    Top =-1
    Right =1401
    Bottom =714
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =89
        Top =37
        Right =330
        Bottom =572
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =401
        Top =166
        Right =545
        Bottom =310
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
