Operation =1
Option =0
Where ="(((Units.ObcUnitNumber) Is Null) AND ((Units.UnitStatus)<>\"B\"))"
Begin InputTables
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.UnitDescShortVer"
    Expression ="Units.ObcUnitNumber"
    Expression ="vw_SixKeys.AcceptedDate"
    Expression ="vw_SixKeys.InServiceDate"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitID"
    Flag =1
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
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ObcUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescShortVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AcceptedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =155
    Top =6
    Right =1627
    Bottom =1005
    Left =-1
    Top =-1
    Right =1440
    Bottom =664
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =116
        Top =237
        Right =456
        Bottom =716
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =484
        Top =106
        Right =763
        Bottom =483
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
