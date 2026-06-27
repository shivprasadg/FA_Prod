Operation =1
Option =0
Where ="(((Units.UnitUnitNum) In (SELECT [UnitUnitNum] FROM [Units] As Tmp GROUP BY [Uni"
    "tUnitNum] HAVING Count(*)>1 )) AND ((Units.ClientID) In (SELECT [ClientId] FROM "
    "[Units] As Tmp GROUP BY [ClientId] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitUnitNum"
    Expression ="Units.ClientID"
    Expression ="Units.UnitID"
    Expression ="Units.GroupREF"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.ObcUnitNumber"
    Expression ="Units.unitstatus"
End
Begin OrderBy
    Expression ="Units.UnitUnitNum"
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
dbMemo "Filter" ="([Find_Dup_Units_UnitNumbers].[ClientID]=3622)"
Begin
    Begin
        dbText "Name" ="Units.[UnitMake]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitModel]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[GroupREF]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[ClientID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitYr]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[ObcUnitNumber]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitUnitNum]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =198
    Top =178
    Right =1636
    Bottom =1106
    Left =-1
    Top =-1
    Right =1406
    Bottom =254
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
