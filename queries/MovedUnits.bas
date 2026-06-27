Operation =1
Option =0
Where ="(((vw_SixKeys.SchID)=846))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="UnitLocationChangeHistory"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.VIN"
    Expression ="UnitLocationChangeHistory.LocationChangeDate"
    Expression ="UnitLocationChangeHistory.PrevLocation"
    Expression ="UnitLocationChangeHistory.NewLocation"
    Expression ="UnitLocationChangeHistory.Reason"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="UnitLocationChangeHistory"
    Expression ="vw_SixKeys.AssetID = UnitLocationChangeHistory.AssetId"
    Flag =1
End
Begin OrderBy
    Expression ="UnitLocationChangeHistory.LocationChangeDate"
    Flag =0
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
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLocationChangeHistory.NewLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLocationChangeHistory.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLocationChangeHistory.PrevLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLocationChangeHistory.LocationChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =165
    Top =25
    Right =1280
    Bottom =1003
    Left =-1
    Top =-1
    Right =1091
    Bottom =698
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =224
        Bottom =593
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =279
        Top =131
        Right =856
        Bottom =528
        Top =0
        Name ="UnitLocationChangeHistory"
        Name =""
    End
End
