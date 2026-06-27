Operation =1
Option =0
Where ="(((vw_UnitVinDetails.unitid) In (SELECT [UNITID] FROM [vw_UnitVinDetails] As Tmp"
    " GROUP BY [UNITID] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="vw_UnitVinDetails"
End
Begin OutputColumns
    Expression ="vw_UnitVinDetails.unitid"
    Expression ="vw_UnitVinDetails.VIN"
    Expression ="vw_UnitVinDetails.modelyear"
    Expression ="vw_UnitVinDetails.MAKE"
    Expression ="vw_UnitVinDetails.MODEL"
End
Begin OrderBy
    Expression ="vw_UnitVinDetails.unitid"
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
        dbText "Name" ="vw_UnitVinDetails.[MODELYEAR]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UnitVinDetails.modelyear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UnitVinDetails.MAKE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UnitVinDetails.UNItID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UnitVinDetails.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UnitVinDetails.MODEL"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =68
    Top =46
    Right =1177
    Bottom =799
    Left =-1
    Top =-1
    Right =1077
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_UnitVinDetails"
        Name =""
    End
End
